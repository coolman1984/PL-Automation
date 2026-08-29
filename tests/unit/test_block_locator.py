import pytest

from src.block_locator import (
    find_month_block_candidates,
    select_unique_month_block,
)
from src.errors import AmbiguousMonthBlockError, ExistingActualColumnError, MissingMonthBlockError
from src.models import HeaderSnapshot, RunCodes


def codes():
    return RunCodes(2026, 8, "August", "2026.008", "T08", "S08", "A08")


def snapshot_for_headers(headers, *, first_row=1):
    width = max(col for row, col, _ in headers)
    height = max(row for row, _, _ in headers)
    values = [[None for _ in range(width)] for _ in range(height)]
    for row, col, value in headers:
        values[row - first_row][col - 1] = value
    return HeaderSnapshot(
        sheet="VD Total",
        first_column=1,
        last_column=width,
        first_row=first_row,
        last_row=height,
        values=tuple(tuple(row) for row in values),
        merged_areas=(),
    )


def test_finds_valid_business_month_block():
    snapshot = snapshot_for_headers(
        [
            (12, 10, "2026.008"),
            (12, 11, "2026.008"),
            (12, 12, "2026.008"),
            (12, 13, "2026.008"),
            (12, 16, "2026.009"),
            (15, 12, "T08"),
            (15, 13, "%"),
            (15, 14, "S08"),
            (15, 15, "%"),
            (14, 10, "August"),
            (14, 16, "September"),
        ]
    )

    candidates = find_month_block_candidates(snapshot, codes(), 487)

    assert len(candidates) == 1
    block = candidates[0]
    assert block.target_col == 12
    assert block.forecast_col == 14
    assert block.insert_at_col == 16
    assert block.september_start_col == 16


def test_two_valid_candidates_are_ambiguous():
    snapshot = snapshot_for_headers(
        [
            (12, 10, "2026.008"), (12, 11, "2026.008"), (12, 12, "2026.008"),
            (12, 13, "2026.008"), (12, 18, "2026.009"),
            (15, 12, "T08"), (15, 13, "%"), (15, 14, "S08"), (15, 15, "%"),
            (14, 10, "August"),
            (14, 16, "September"),
            (12, 20, "2026.008"), (12, 21, "2026.008"), (12, 22, "2026.008"),
            (12, 23, "2026.008"), (12, 26, "2026.009"),
            (15, 22, "T08"), (15, 23, "%"), (15, 24, "S08"), (15, 25, "%"),
            (14, 26, "September"),
        ]
    )
    candidates = find_month_block_candidates(snapshot, codes(), 487)

    with pytest.raises(AmbiguousMonthBlockError):
        select_unique_month_block(candidates)


def test_no_valid_candidate_fails():
    snapshot = snapshot_for_headers([(15, 12, "T08"), (15, 13, "%")])

    candidates = find_month_block_candidates(snapshot, codes(), 487)

    with pytest.raises(MissingMonthBlockError):
        select_unique_month_block(candidates)


def test_existing_actual_is_reported_in_block():
    snapshot = snapshot_for_headers(
        [
            (12, 10, "2026.008"), (12, 11, "2026.008"), (12, 12, "2026.008"),
            (12, 13, "2026.008"), (12, 18, "2026.009"),
            (15, 12, "T08"), (15, 13, "%"), (15, 14, "S08"), (15, 15, "%"),
            (14, 10, "August"),
            (15, 16, "A08"), (15, 17, "%"), (14, 10, "August"),
            (14, 18, "September"),
        ]
    )
    block = find_month_block_candidates(snapshot, codes(), 487)[0]

    with pytest.raises(ExistingActualColumnError):
        from src.block_locator import ensure_actual_absent
        ensure_actual_absent(snapshot, block, codes().actual_version)
