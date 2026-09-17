"""Wave-aware weighted BICS time series from the official latest-wave workbook."""

from __future__ import annotations

import hashlib
import io
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime, time
from typing import Any
from urllib.parse import urljoin

import httpx
import openpyxl

from scripts.config import MAX_DOWNLOAD_BYTES, REQUEST_TIMEOUT, USER_AGENT
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

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
    for sheet_name, responses in SHEETS.items():
        if sheet_name not in book.sheetnames:
            raise ValueError(f"BICS question sheet missing: {sheet_name}")
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
    latest = max(o.reference_date for o in observations)
    for o in observations:
        current = o.reference_date == latest
        availability[(o.series_id, o.reference_date)] = (
            release if current else collected,
            "official_timestamp" if current else "first_seen",
            release.date() if current else None,
        )
    if len(observations) < 500 or len(catalog) < 50:
        raise ValueError("BICS selected history unexpectedly short")
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
