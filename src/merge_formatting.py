"""Formatting, column properties, and surgical August merge repair."""

from __future__ import annotations

from typing import Sequence

from .errors import MergeRepairError
from .models import ColumnProperties, MergedArea, MonthBlock


def capture_column_properties(
    worksheet: object, columns: Sequence[int]
) -> dict[int, ColumnProperties]:
    result: dict[int, ColumnProperties] = {}
    for column in columns:
        column_range = worksheet.Columns(int(column))
        try:
            width = float(column_range.ColumnWidth)
        except Exception:
            width = 0.0
        try:
            hidden = bool(column_range.Hidden)
        except Exception:
            hidden = False
        try:
            outline = int(column_range.OutlineLevel)
        except Exception:
            outline = None
        result[int(column)] = ColumnProperties(width=width, hidden=hidden, outline_level=outline)
    return result


def restore_column_properties(
    worksheet: object,
    source_properties: dict[int, ColumnProperties],
    destination_columns: Sequence[int],
) -> None:
    if len(source_properties) != len(destination_columns):
        raise MergeRepairError("Source and destination column-property counts do not match")
    for source_column, destination_column in zip(sorted(source_properties), destination_columns):
        source = source_properties[source_column]
        target = worksheet.Columns(int(destination_column))
        try:
            if source.width > 0:
                target.ColumnWidth = source.width
            target.Hidden = source.hidden
            if source.outline_level is not None:
                target.OutlineLevel = source.outline_level
        except Exception as exc:  # pragma: no cover - requires Excel
            raise MergeRepairError(
                f"Could not restore properties for inserted column {destination_column}: {exc}"
            ) from exc


def capture_month_merge(worksheet: object, block: MonthBlock) -> MergedArea | None:
    if block.month_header_row is None or block.month_merge is None:
        return None
    return block.month_merge


def _capture_format(cell: object) -> dict[str, object]:
    properties = (
        "HorizontalAlignment",
        "VerticalAlignment",
        "NumberFormat",
        "Interior",
        "Font",
        "Borders",
        "Locked",
    )
    result: dict[str, object] = {}
    for prop in properties:
        try:
            result[prop] = getattr(cell, prop)
        except Exception:
            pass
    return result


def _restore_format(target: object, saved: dict[str, object]) -> None:
    for prop, value in saved.items():
        try:
            setattr(target, prop, value)
        except Exception:
            # Formatting is best effort after the specific merge is proven safe;
            # native copy remains the primary formatting preservation mechanism.
            pass


def ensure_august_merge_extended(
    worksheet: object, block: MonthBlock, actual_pct_col: int
) -> bool:
    merge = block.month_merge
    if merge is None or block.month_header_row is None:
        return False
    first_row = merge.first_row
    first_col = merge.first_column
    desired_last_col = actual_pct_col
    if desired_last_col <= merge.last_column:
        return False

    try:
        anchor_cell = worksheet.Cells(first_row, first_col)
        current_area = anchor_cell.MergeArea if bool(anchor_cell.MergeCells) else None
        if current_area is None:
            raise MergeRepairError(
                f"August merge disappeared before repair on {block.sheet}"
            )
        current_address = str(current_area.Address)
        saved_value = anchor_cell.Value2
        saved_format = _capture_format(anchor_cell)

        # Inspect only the extension columns. Any unrelated merge makes repair
        # ambiguous, so the update must stop rather than unmerge it.
        for column in range(current_area.Column + current_area.Columns.Count, desired_last_col + 1):
            cell = worksheet.Cells(first_row, column)
            if bool(cell.MergeCells):
                other = cell.MergeArea
                if str(other.Address) != current_address:
                    raise MergeRepairError(
                        f"August merge extension intersects unrelated merge {other.Address}"
                    )

        current_area.UnMerge()
        desired_range = worksheet.Range(
            worksheet.Cells(first_row, first_col),
            worksheet.Cells(merge.last_row, desired_last_col),
        )
        desired_range.Merge()
        desired_range.Cells(1, 1).Value2 = saved_value
        _restore_format(desired_range.Cells(1, 1), saved_format)
        repaired_area = desired_range.Cells(1, 1).MergeArea
        if int(repaired_area.Column) != first_col or int(
            repaired_area.Column + repaired_area.Columns.Count - 1
        ) != desired_last_col:
            raise MergeRepairError(
                f"August merge repair did not cover columns {first_col}:{desired_last_col}"
            )
        return True
    except MergeRepairError:
        raise
    except Exception as exc:  # pragma: no cover - requires Excel
        raise MergeRepairError(f"Could not repair August merge on {block.sheet}: {exc}") from exc
