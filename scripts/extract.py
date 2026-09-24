"""Wave-aware weighted BICS time series from the official latest-wave workbook."""

from __future__ import annotations

import hashlib
import io
import logging
import math
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime, time
from typing import Any
from urllib.parse import urljoin

import httpx
import openpyxl

from scripts.config import (
    MAX_DOWNLOAD_BYTES,
    MAX_STALE_MONTHS,
    MIN_HISTORY_YEARS,
    REQUEST_TIMEOUT,
    USER_AGENT,
)
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

logger = logging.getLogger(__name__)


# -- series_id contract (GUIDELINES.md 4) ---------------------------------
# series_id is uppercase, underscore-separated and ordered coarse -> fine. The
# pair below is the canonical public surface: parse splits an id into its
# components, build rejoins them, and build(*parse(sid)) == sid for every id
# this collector emits. Only economic identity is encoded -- never a delivery
# provider or any other detail of how the value reached us.


def parse_series_id(series_id: str) -> tuple[str, ...]:
    """Split a series_id into its underscore-delimited components.

    Raises ValueError on anything this collector would not have produced:
    lowercase, empty components, or an id with no structure at all.
    """
    if not series_id or series_id != series_id.upper():
        raise ValueError(f"series_id must be uppercase: {series_id!r}")
    components = tuple(series_id.split("_"))
    if any(not component for component in components):
        raise ValueError(f"series_id has an empty component: {series_id!r}")
    return components


def build_series_id(*components: str) -> str:
    """Rejoin the tuple parse_series_id returned into the original id."""
    if not components:
        raise ValueError("series_id needs at least one component")
    if any(not component or component != component.upper() for component in components):
        raise ValueError(f"invalid series_id components: {components!r}")
    return "_".join(components)


# -- 5.1 usable-series filtering ------------------------------------------


@dataclass(frozen=True)
class UsabilityReport:
    """What the filter removed, for logging and for tests to assert on."""

    kept: tuple[str, ...]
    stale: tuple[str, ...]
    short_history: tuple[str, ...]
    empty: tuple[str, ...]

    @property
    def dropped(self) -> tuple[str, ...]:
        return tuple(sorted(set(self.stale) | set(self.short_history) | set(self.empty)))


def _months_between(earlier: date, later: date) -> int:
    """Whole months from ``earlier`` to ``later``, day-of-month aware."""
    months = (later.year - earlier.year) * 12 + (later.month - earlier.month)
    if later.day < earlier.day:
        months -= 1
    return months


def _is_valid(value: Any) -> bool:
    """A real observation: present, numeric and finite."""
    if value is None:
        return False
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return False
    return math.isfinite(numeric)


def assess_series(
    reference_dates: list[date],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> str:
    """Classify one series from the reference dates of its valid observations.

    Returns ``"keep"``, ``"empty"``, ``"stale"`` or ``"short_history"``.
    Recency is judged at the period end and over non-null values only: a source
    that keeps listing a discontinued series with empty recent cells must not
    look live because of those blanks.
    """
    if not reference_dates:
        return "empty"
    first, last = min(reference_dates), max(reference_dates)
    if _months_between(last, today) > max_stale_months:
        return "stale"
    if _months_between(first, last) < round(min_history_years * 12):
        return "short_history"
    return "keep"


def filter_usable_series(
    observations: list[Any],
    catalog: dict[str, dict[str, Any]],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> tuple[list[Any], dict[str, dict[str, Any]], UsabilityReport]:
    """Drop obsolete and history-less series before anything is persisted.

    Runs after parsing and before the time_series / metadata upsert, so the
    standardized tables never carry a dead or stub series, and prunes the
    catalog alongside the observations so metadata can never describe a series
    the database does not hold (GUIDELINES.md 5.1).
    """
    valid_dates: dict[str, list[date]] = {}
    for observation in observations:
        if _is_valid(observation.value):
            valid_dates.setdefault(observation.series_id, []).append(observation.reference_date)

    verdicts: dict[str, str] = {}
    for series_id in set(catalog) | {o.series_id for o in observations}:
        verdicts[series_id] = assess_series(
            valid_dates.get(series_id, []), today, max_stale_months, min_history_years
        )

    keep = {series_id for series_id, verdict in verdicts.items() if verdict == "keep"}
    report = UsabilityReport(
        kept=tuple(sorted(keep)),
        stale=tuple(sorted(s for s, v in verdicts.items() if v == "stale")),
        short_history=tuple(sorted(s for s, v in verdicts.items() if v == "short_history")),
        empty=tuple(sorted(s for s, v in verdicts.items() if v == "empty")),
    )

    if report.dropped:
        logger.info(
            "Usable-series filter: kept %d, dropped %d "
            "(stale=%d short_history=%d empty=%d; max_stale_months=%d min_history_years=%s)",
            len(report.kept),
            len(report.dropped),
            len(report.stale),
            len(report.short_history),
            len(report.empty),
            max_stale_months,
            min_history_years,
        )
        for series_id in report.stale:
            logger.info(
                "Dropped %s: last valid observation older than %d months",
                series_id,
                max_stale_months,
            )
        for series_id in report.short_history:
            logger.info(
                "Dropped %s: valid history shorter than %s years", series_id, min_history_years
            )
        for series_id in report.empty:
            logger.info("Dropped %s: no valid observations", series_id)
    else:
        logger.info("Usable-series filter: all %d series usable", len(report.kept))

    kept_observations = [o for o in observations if o.series_id in keep]
    kept_catalog = {sid: fields for sid, fields in catalog.items() if sid in keep}
    return kept_observations, kept_catalog, report


LANDING = "https://www.ons.gov.uk/economy/economicoutputandproductivity/output/datasets/businessinsightsandimpactontheukeconomy"
SHEETS = {
    "Staffing Costs TS (WTD) (1)": {"Costs have increased"},
    "Employment Costs TS (WTD)": {"Increase prices"},
    "Energy Prices Concern TS (WTD)": {"Very concerned", "Somewhat concerned"},
    "Supply Chain Concerns TS (WTD)": {
        "International conflict",
        "Shipping disruption",
        "Increased barriers to trade",
    },
}


@dataclass(frozen=True)
class ExtractedData:
    observations: list[Observation]
    snapshots: list[Snapshot]
    catalog: dict[str, dict[str, Any]]
    releases: list[datetime]
    availability_by_key: dict[tuple[str, date], tuple[datetime, str, date | None]]
    min_lag_days: int = 0
    max_lag_days: int = 0
    inferred_lag_days: int | None = None


def _slug(x: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", x.upper()).strip("_")


def _end_date(x: str) -> date:
    m = re.search(r"to (\d{1,2} \w+ \d{4})$", x.strip())
    if not m:
        raise ValueError(f"BICS date range drifted: {x!r}")
    return datetime.strptime(m.group(1), "%d %B %Y").replace(tzinfo=UTC).date()


def parse_xlsx(
    body: bytes, snapshot_id: str, url: str, collected: datetime
) -> tuple[
    list[Observation],
    dict[str, dict[str, Any]],
    datetime,
    dict[tuple[str, date], tuple[datetime, str, date | None]],
]:
    book = openpyxl.load_workbook(io.BytesIO(body), data_only=True, read_only=False)
    observations = []
    catalog = {}
    availability = {}
    keys = set()
    release = None
    if "Weighted Time Series" not in book.sheetnames:
        raise ValueError("BICS weighted time-series index missing")
    intro = book["Weighted Time Series"]
    for row in intro.iter_rows(values_only=True):
        if row and isinstance(row[0], str) and row[0].startswith("Released "):
            release = datetime.combine(
                datetime.strptime(row[0][9:], "%d %B %Y").replace(tzinfo=UTC).date(),
                time(7),
                tzinfo=UTC,
            )
    if release is None:
        raise ValueError("BICS release date missing")
    # BICS is a rotating-module survey: an individual wave workbook carries only
    # the question sheets asked in that wave, and each sheet it does carry holds
    # that question's full time series. A configured sheet being absent from one
    # wave is therefore normal and must not fail the run -- the next wave that
    # asks the question restores its history in full. Structural drift is still
    # caught below: if the workbook carries none of the configured questions, the
    # layout itself has changed and the run fails loudly.
    present = [name for name in SHEETS if name in book.sheetnames]
    if not present:
        raise ValueError(
            "BICS workbook carries none of the configured question sheets "
            f"({sorted(SHEETS)}); the source layout has changed"
        )
    for sheet_name in sorted(set(SHEETS) - set(present)):
        logger.info("BICS question %r not asked in this wave; skipping", sheet_name)
    for sheet_name in present:
        responses = SHEETS[sheet_name]
        sheet = book[sheet_name]
        header = [str(sheet.cell(6, c).value or "").strip() for c in range(1, sheet.max_column + 1)]
        columns = {name: header.index(name) + 1 for name in responses if name in header}
        if set(columns) != responses:
            raise ValueError(f"BICS response categories drifted in {sheet_name}")
        question = str(sheet.cell(1, 1).value or "").removeprefix("Question: ")
        for row in range(7, sheet.max_row + 1):
            period = sheet.cell(row, 1).value
            wave = str(sheet.cell(row, 2).value or "")
            dimension = str(sheet.cell(row, 3).value or "").strip()
            if not isinstance(period, str) or not wave.startswith("Wave ") or not dimension:
                continue
            reference = _end_date(period)
            for response, col in columns.items():
                raw = sheet.cell(row, col).value
                if not isinstance(raw, (int, float)):
                    continue
                series_id = f"ONS_BICS_{_slug(sheet_name.replace(' TS (WTD)', ''))}_{_slug(response)}_{_slug(dimension)}"
                key = (series_id, reference)
                if key in keys:
                    raise ValueError(f"Duplicate BICS key {key}")
                keys.add(key)
                observations.append(Observation(series_id, reference, float(raw), snapshot_id))
                catalog[series_id] = {
                    "source_id": "ons_bics",
                    "name": f"{dimension}: {response}",
                    "description": f"Raw weighted share. Question wording: {question}. Wave is retained by the source snapshot and reference date; series is not stitched across changed question sheets.",
                    "frequency": "biweekly",
                    "unit": "percent",
                    "eco_group": "surveys",
                    "source_url": url,
                    "last_publish_date": release.date(),
                }
    # Scale the volume floor by the number of question sheets this wave actually
    # carries. A rotating module means a legitimate wave can hold a single
    # question, so a fixed fleet-wide floor would reject good data; a per-sheet
    # floor still catches a truncated or mis-parsed sheet. Observed yields are
    # ~400-470 observations and ~35-50 series per question sheet.
    if len(observations) < 300 * len(present) or len(catalog) < 25 * len(present):
        raise ValueError(
            f"BICS selected history unexpectedly short: {len(observations)} observations "
            f"and {len(catalog)} series across {len(present)} question sheet(s)"
        )
    latest = max(o.reference_date for o in observations)
    for o in observations:
        current = o.reference_date == latest
        availability[(o.series_id, o.reference_date)] = (
            release if current else collected,
            "official_timestamp" if current else "first_seen",
            release.date() if current else None,
        )
    return observations, catalog, release, availability


def collect() -> ExtractedData:
    fetched = datetime.now(UTC)
    with httpx.Client(
        timeout=REQUEST_TIMEOUT, headers={"User-Agent": USER_AGENT}, follow_redirects=True
    ) as client:
        page = client.get(LANDING)
        page.raise_for_status()
        links = re.findall(
            r'href=["\']([^"\']*file\?uri=[^"\']*/bicswave(\d+)/[^"\']+\.xlsx)["\']',
            page.text,
            re.IGNORECASE,
        )
        if not links:
            raise ValueError("No BICS wave workbook discovered")
        path, wave = max(links, key=lambda item: int(item[1]))
        url = urljoin(LANDING, path)
        response = client.get(url)
        response.raise_for_status()
    body = response.content
    if not body or len(body) > MAX_DOWNLOAD_BYTES:
        raise ValueError(f"Invalid BICS artifact size {len(body)}")
    digest = hashlib.sha256(body).hexdigest()
    obs, catalog, release, availability = parse_xlsx(body, digest, url, fetched)
    snapshot = build_snapshot(
        "ons_bics",
        url,
        f"bics_wave_{wave}.xlsx",
        body,
        digest,
        response.headers.get("etag"),
        response.headers.get("last-modified"),
        fetched,
        release.date(),
    )
    return ExtractedData(obs, [snapshot], catalog, [release], availability)
