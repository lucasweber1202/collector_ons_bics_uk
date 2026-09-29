"""Historical BICS waves must preserve question identity and publication order."""

from __future__ import annotations

from datetime import UTC, date, datetime

from scripts.extract import stitch_waves
from scripts.time_series import Observation


def test_overlapping_waves_keep_newer_revision_and_extend_history() -> None:
    sid = "ONS_BICS_ENERGY_PRICE_CONCERN_VERY_CONCERNED_ALL"
    old = date(2024, 1, 1)
    overlap = date(2026, 1, 1)
    newest = date(2026, 9, 1)
    newer_at = datetime(2026, 9, 20, tzinfo=UTC)
    older_at = datetime(2026, 2, 1, tzinfo=UTC)
    newer = (
        [Observation(sid, overlap, 22.0, "new"), Observation(sid, newest, 24.0, "new")],
        {sid: {"description": "Question wording: energy prices"}},
        {
            (sid, overlap): (newer_at, "official_timestamp", newest),
            (sid, newest): (newer_at, "official_timestamp", newest),
        },
    )
    older = (
        [Observation(sid, old, 20.0, "old"), Observation(sid, overlap, 21.0, "old")],
        {sid: {"description": "Question wording: energy prices"}},
        {
            (sid, old): (older_at, "first_seen", None),
            (sid, overlap): (older_at, "official_timestamp", overlap),
        },
    )
    observations, catalog, availability = stitch_waves([newer, older])
    assert len(observations) == 3
    assert {(o.reference_date, o.value) for o in observations} == {
        (old, 20.0),
        (overlap, 22.0),
        (newest, 24.0),
    }
    assert catalog[sid] == newer[1][sid]
    assert availability[(sid, overlap)][0] == newer_at


def test_changed_question_wording_does_not_create_false_history() -> None:
    sid = "ONS_BICS_ENERGY_PRICE_CONCERN_VERY_CONCERNED_ALL"
    latest = date(2026, 9, 1)
    old = date(2024, 1, 1)
    at = datetime(2026, 9, 20, tzinfo=UTC)
    waves = [
        (
            [Observation(sid, latest, 24.0, "new")],
            {sid: {"description": "Question wording: energy prices"}},
            {(sid, latest): (at, "official_timestamp", latest)},
        ),
        (
            [Observation(sid, old, 20.0, "old")],
            {sid: {"description": "Question wording: fuel costs"}},
            {(sid, old): (at, "first_seen", None)},
        ),
    ]
    observations, _, availability = stitch_waves(waves)
    assert [o.reference_date for o in observations] == [latest]
    assert (sid, old) not in availability
