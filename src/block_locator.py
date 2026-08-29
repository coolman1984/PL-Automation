"""Semantic August block discovery with fail-closed ambiguity handling."""

from __future__ import annotations

from collections.abc import Sequence

from .errors import (
    AmbiguousMonthBlockError,
    ExistingActualColumnError,
    MissingMonthBlockError,
)
from .header_discovery import (
    columns_matching,
    effective_merged_value,
    normalize_header_value,
    value_at,
)
from .models import HeaderSnapshot, MonthBlock, RunCodes


_AUGUST_ALIASES = {"AUGUST", "AUGUT"}
_SEPTEMBER_ALIASES = {"SEPTEMBER", "SEPTEMPER", "SEP", "SEPT"}


def _row_value(snapshot: HeaderSnapshot, row: int, column: int) -> str:
    return normalize_header_value(effective_merged_value(snapshot, row, column))


def _find_period(snapshot: HeaderSnapshot, column: int, period: str, before_row: int) -> int | None:
    for row in range(snapshot.first_row, min(before_row, snapshot.last_row + 1)):
        if _row_value(snapshot, row, column) == normalize_header_value(period):
            return row
    return None


def _find_month_anchor(snapshot: HeaderSnapshot, aliases: set[str], column: int, before_row: int):
    for area in snapshot.merged_areas:
        if area.first_row < before_row and area.first_column <= column <= area.last_column:
            if normalize_header_value(area.top_left_value) in aliases:
                return area.first_row, area
    for row in range(snapshot.first_row, min(before_row, snapshot.last_row + 1)):
        for candidate_column in range(snapshot.first_column, column + 1):
            if _row_value(snapshot, row, candidate_column) in aliases:
                return row, None
    return None, None


def _september_boundary(
    snapshot: HeaderSnapshot,
    start_column: int,
    period: str,
    *,
    version_header_row: int,
    actual_version: str,
) -> int | None:
    next_period = f"{period.rsplit('.', 1)[0]}.009" if "." in period else None
    for column in range(start_column, snapshot.last_column + 1):
        # An already-updated workbook may have an A08 pair between S08/% and
        # September. Keep the candidate so the idempotency check can report the
        # specific existing pair rather than misclassifying the block.
        if (
            normalize_header_value(value_at(snapshot, version_header_row, column))
            == normalize_header_value(actual_version)
            and normalize_header_value(value_at(snapshot, version_header_row, column + 1))
            == "%"
        ):
            continue
        if column > start_column and normalize_header_value(
            value_at(snapshot, version_header_row, column - 1)
        ) == normalize_header_value(actual_version):
            continue
        if next_period and any(
            _row_value(snapshot, row, column) == next_period
            for row in range(snapshot.first_row, snapshot.last_row + 1)
        ):
            return column
        if any(
            _row_value(snapshot, row, column) in _SEPTEMBER_ALIASES
            for row in range(snapshot.first_row, snapshot.last_row + 1)
        ):
            return column
    return None


def find_month_block_candidates(
    snapshot: HeaderSnapshot,
    codes: RunCodes,
    last_used_row: int,
    *,
    require_period: bool | None = None,
    source_columns: dict[str, int] | None = None,
) -> list[MonthBlock]:
    """Return only candidates satisfying every mandatory anchor.

    `Total PL` in the supplied workbook has no visible 2026.008 period row. For
    that sheet the caller may set `require_period=False` and provide the
    business-sheet T08 columns; the lineage check then proves the current block.
    """
    if require_period is None:
        require_period = snapshot.sheet != "Total PL"
    candidates: list[MonthBlock] = []
    seen: set[tuple[int, int]] = set()
    for version_row, target_col in columns_matching(snapshot, codes.target_version):
        key = (version_row, target_col)
        if key in seen:
            continue
        seen.add(key)
        target_pct_col = target_col + 1
        forecast_col = target_col + 2
        forecast_pct_col = target_col + 3
        if normalize_header_value(value_at(snapshot, version_row, target_pct_col)) != "%":
            continue
        if normalize_header_value(value_at(snapshot, version_row, forecast_col)) != normalize_header_value(codes.forecast_version):
            continue
        if normalize_header_value(value_at(snapshot, version_row, forecast_pct_col)) != "%":
            continue

        period_row = _find_period(snapshot, target_col, codes.period, version_row)
        if require_period and period_row is None:
            continue
        month_row, month_merge = _find_month_anchor(snapshot, _AUGUST_ALIASES, target_col, version_row)
        if require_period and month_row is None:
            continue

        insert_at = forecast_pct_col + 1
        september_start = _september_boundary(
            snapshot,
            insert_at,
            codes.period,
            version_header_row=version_row,
            actual_version=codes.actual_version,
        )
        if september_start is None:
            continue
        actual_at_insert = (
            normalize_header_value(value_at(snapshot, version_row, insert_at))
            == normalize_header_value(codes.actual_version)
            and normalize_header_value(value_at(snapshot, version_row, insert_at + 1))
            == "%"
        )
        if september_start != insert_at and not actual_at_insert:
            # A valid block must transition directly from S08/% to September.
            continue

        evidence = [
            f"target={version_row}:{target_col}",
            f"target_pct={version_row}:{target_pct_col}",
            f"forecast={version_row}:{forecast_col}",
            f"forecast_pct={version_row}:{forecast_pct_col}",
            f"insert_at={insert_at}",
            f"september={september_start}",
        ]
        if period_row is not None:
            evidence.append(f"period={period_row}:{target_col}={codes.period}")
        if month_row is not None:
            evidence.append(f"month={month_row}:{target_col}=August")
        if source_columns:
            evidence.append("Total PL lineage columns supplied")

        candidates.append(
            MonthBlock(
                sheet=snapshot.sheet,
                year=codes.year,
                month=codes.month,
                period=codes.period,
                target_col=target_col,
                target_pct_col=target_pct_col,
                forecast_col=forecast_col,
                forecast_pct_col=forecast_pct_col,
                insert_at_col=insert_at,
                version_header_row=version_row,
                period_header_row=period_row,
                month_header_row=month_row,
                month_merge=month_merge,
                september_start_col=september_start,
                last_used_row=last_used_row,
                evidence=tuple(evidence),
            )
        )
    return candidates


def select_unique_month_block(
    candidates: Sequence[MonthBlock], *, require_unique: bool = True
) -> MonthBlock:
    if not candidates:
        raise MissingMonthBlockError("No unique August T08/S08 block satisfied all required anchors")
    if require_unique and len(candidates) != 1:
        raise AmbiguousMonthBlockError(
            f"{len(candidates)} August T08/S08 blocks satisfied the required anchors",
            evidence=[candidate.evidence for candidate in candidates],
        )
    return candidates[0]


def ensure_actual_absent(
    snapshot: HeaderSnapshot, block: MonthBlock, actual_version: str
) -> None:
    wanted = normalize_header_value(actual_version)
    matches = []
    for row in range(snapshot.first_row, snapshot.last_row + 1):
        for column in range(block.start_col, block.september_start_col):
            if normalize_header_value(value_at(snapshot, row, column)) == wanted:
                matches.append((row, column))
    if matches:
        raise ExistingActualColumnError(
            f"{actual_version} already exists in {block.sheet} August block",
            evidence={"matches": matches, "block": block.evidence},
        )


def detect_existing_actual(
    snapshot: HeaderSnapshot, block: MonthBlock, actual_version: str
) -> list[tuple[int, int]]:
    wanted = normalize_header_value(actual_version)
    return [
        (row, column)
        for row in range(snapshot.first_row, snapshot.last_row + 1)
        for column in range(block.start_col, block.september_start_col)
        if normalize_header_value(value_at(snapshot, row, column)) == wanted
    ]
