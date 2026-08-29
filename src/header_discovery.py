"""Bulk header snapshots and merged-header helpers."""

from __future__ import annotations

from typing import Any

from .constants import HEADER_FIRST_ROW, HEADER_LAST_ROW
from .models import HeaderSnapshot, MergedArea


def normalize_header_value(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().upper()


def _as_matrix(value: Any) -> tuple[tuple[Any, ...], ...]:
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return tuple(tuple(row) for row in value)
        return (tuple(value),)
    return ((value,),)


def _merged_area_from_cell(cell: object) -> MergedArea | None:
    try:
        if not bool(cell.MergeCells):
            return None
        area = cell.MergeArea
        top_left = area.Cells(1, 1).Value2
        return MergedArea(
            first_row=int(area.Row),
            first_column=int(area.Column),
            row_count=int(area.Rows.Count),
            column_count=int(area.Columns.Count),
            top_left_value=top_left,
        )
    except Exception:
        return None


def read_header_snapshot(
    worksheet: object,
    *,
    first_row: int = HEADER_FIRST_ROW,
    last_row: int = HEADER_LAST_ROW,
) -> HeaderSnapshot:
    used = worksheet.UsedRange
    first_column = max(1, int(used.Column))
    last_column = int(used.Column + used.Columns.Count - 1)
    if last_column < first_column:
        last_column = first_column
    header_range = worksheet.Range(
        worksheet.Cells(first_row, first_column),
        worksheet.Cells(last_row, last_column),
    )
    values = _as_matrix(header_range.Value2)

    # Only probe cells whose bulk value can plausibly be a merged month label.
    # This avoids a cell-by-cell scan across the million-cell workbook.
    month_words = {
        "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
        "AUGUST", "AUGUT", "SEPTEMBER", "SEP", "SEPTEMPER", "SEPTEMBER",
        "OCTOBER", "NOVEMBER", "DECEMBER",
    }
    merges: dict[tuple[int, int, int, int], MergedArea] = {}
    for row_offset, row in enumerate(values):
        for col_offset, value in enumerate(row):
            if normalize_header_value(value) not in month_words:
                continue
            row_number = first_row + row_offset
            column_number = first_column + col_offset
            area = _merged_area_from_cell(worksheet.Cells(row_number, column_number))
            if area:
                key = (area.first_row, area.first_column, area.row_count, area.column_count)
                merges[key] = area

    return HeaderSnapshot(
        sheet=str(worksheet.Name),
        first_column=first_column,
        last_column=last_column,
        first_row=first_row,
        last_row=last_row,
        values=values,
        merged_areas=tuple(merges.values()),
    )


def value_at(snapshot: HeaderSnapshot, row: int, column: int) -> object:
    if not (snapshot.first_row <= row <= snapshot.last_row):
        return None
    if not (snapshot.first_column <= column <= snapshot.last_column):
        return None
    return snapshot.values[row - snapshot.first_row][column - snapshot.first_column]


def effective_merged_value(snapshot: HeaderSnapshot, row: int, column: int) -> str:
    for area in snapshot.merged_areas:
        if area.first_row <= row <= area.last_row and area.first_column <= column <= area.last_column:
            return normalize_header_value(area.top_left_value)
    return normalize_header_value(value_at(snapshot, row, column))


def columns_matching(snapshot: HeaderSnapshot, exact_text: str) -> list[tuple[int, int]]:
    wanted = normalize_header_value(exact_text)
    matches: list[tuple[int, int]] = []
    for row_offset, row in enumerate(snapshot.values):
        for col_offset, value in enumerate(row):
            if normalize_header_value(value) == wanted:
                matches.append((snapshot.first_row + row_offset, snapshot.first_column + col_offset))
    return matches

