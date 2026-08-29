"""Pure unit tests for bulk header snapshots and merged-value resolution."""

from __future__ import annotations

from src.header_discovery import (
    columns_matching,
    effective_merged_value,
    normalize_header_value,
    value_at,
)
from src.models import HeaderSnapshot, MergedArea


def _snapshot(
    values, merges=(), first_row=1, first_column=1
) -> HeaderSnapshot:
    rows = tuple(tuple(row) for row in values)
    return HeaderSnapshot(
        sheet="VD Total",
        first_column=first_column,
        last_column=first_column + len(rows[0]) - 1 if rows else first_column,
        first_row=first_row,
        last_row=first_row + len(rows) - 1,
        values=rows,
        merged_areas=tuple(merges),
    )


def test_normalize_header_value_strips_and_uppercases():
    assert normalize_header_value("  August ") == "AUGUST"
    assert normalize_header_value(None) == ""
    assert normalize_header_value("t08") == "T08"


def test_value_at_returns_none_outside_bounds():
    snapshot = _snapshot([["A", "B"], ["C", "D"]])
    assert value_at(snapshot, 1, 2) == "B"
    assert value_at(snapshot, 2, 2) == "D"
    assert value_at(snapshot, 0, 1) is None
    assert value_at(snapshot, 3, 1) is None
    assert value_at(snapshot, 1, 5) is None


def test_effective_merged_value_uses_merge_top_left():
    merge = MergedArea(
        first_row=1, first_column=4, row_count=2, column_count=2,
        top_left_value="August",
    )
    snapshot = _snapshot(
        [
            ["x", "x", "x", "August"],
            ["x", "x", "x", None],
        ],
        merges=(merge,),
    )
    # Every covered cell resolves to the merged label even though stored as None.
    assert effective_merged_value(snapshot, 1, 4) == "AUGUST"
    assert effective_merged_value(snapshot, 1, 5) == "AUGUST"
    assert effective_merged_value(snapshot, 2, 5) == "AUGUST"
    assert effective_merged_value(snapshot, 2, 1) == "X"


def test_columns_matching_finds_all_occurrences_case_insensitively():
    snapshot = _snapshot(
        [
            [None, "t08", None, "T08"],
            ["T08", None, None, None],
        ]
    )
    matches = columns_matching(snapshot, "T08")
    assert matches == [(1, 2), (1, 4), (2, 1)]


def test_columns_matching_no_match_returns_empty():
    snapshot = _snapshot([[None, "S08"]])
    assert columns_matching(snapshot, "A08") == []


def test_snapshot_with_offsets_preserves_real_sheet_geometry():
    merge = MergedArea(
        first_row=3, first_column=7, row_count=1, column_count=2,
        top_left_value="September",
    )
    snapshot = _snapshot(
        [["", ""], ["", ""], ["September", None]],
        merges=(merge,),
        first_row=2,
        first_column=6,
    )
    assert value_at(snapshot, 4, 6) == "September"
    assert effective_merged_value(snapshot, 3, 8) == "SEPTEMBER"
