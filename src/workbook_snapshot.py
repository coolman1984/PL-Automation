"""Excel COM workbook inventory and optional full cell/style JSON snapshot."""

from __future__ import annotations

import hashlib
import json
import os
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .workbook_audit import collect_fingerprint


@dataclass(frozen=True)
class SnapshotOptions:
    mode: str = "auto"
    max_cells: int = 250_000


def json_safe(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, (list, tuple)):
        return [json_safe(item) for item in value]
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    return str(value)


def _safe_get(obj: object, attribute: str, default: Any = None) -> Any:
    try:
        return getattr(obj, attribute)
    except Exception:
        return default


def _style_payload(cell: object) -> dict[str, Any]:
    font = _safe_get(cell, "Font")
    interior = _safe_get(cell, "Interior")
    protection = _safe_get(cell, "Protection")
    borders: dict[str, Any] = {}
    border_collection = _safe_get(cell, "Borders")
    if border_collection is not None:
        for name, index in (("diagonal_down", 5), ("diagonal_up", 6), ("left", 7), ("top", 8), ("bottom", 9), ("right", 10)):
            try:
                border = border_collection(index)
                borders[name] = {
                    "line_style": json_safe(_safe_get(border, "LineStyle")),
                    "weight": json_safe(_safe_get(border, "Weight")),
                    "color": json_safe(_safe_get(border, "Color")),
                    "color_index": json_safe(_safe_get(border, "ColorIndex")),
                }
            except Exception:
                continue
    return {
        "style_name": json_safe(_safe_get(cell, "Style")),
        "number_format": json_safe(_safe_get(cell, "NumberFormat")),
        "horizontal_alignment": json_safe(_safe_get(cell, "HorizontalAlignment")),
        "vertical_alignment": json_safe(_safe_get(cell, "VerticalAlignment")),
        "wrap_text": json_safe(_safe_get(cell, "WrapText")),
        "orientation": json_safe(_safe_get(cell, "Orientation")),
        "indent_level": json_safe(_safe_get(cell, "IndentLevel")),
        "shrink_to_fit": json_safe(_safe_get(cell, "ShrinkToFit")),
        "font": {
            "name": json_safe(_safe_get(font, "Name")),
            "size": json_safe(_safe_get(font, "Size")),
            "bold": json_safe(_safe_get(font, "Bold")),
            "italic": json_safe(_safe_get(font, "Italic")),
            "underline": json_safe(_safe_get(font, "Underline")),
            "strikethrough": json_safe(_safe_get(font, "Strikethrough")),
            "color": json_safe(_safe_get(font, "Color")),
            "color_index": json_safe(_safe_get(font, "ColorIndex")),
        },
        "fill": {
            "color": json_safe(_safe_get(interior, "Color")),
            "color_index": json_safe(_safe_get(interior, "ColorIndex")),
            "pattern": json_safe(_safe_get(interior, "Pattern")),
            "pattern_color": json_safe(_safe_get(interior, "PatternColor")),
        },
        "borders": borders,
        "protection": {
            "locked": json_safe(_safe_get(protection, "Locked")),
            "formula_hidden": json_safe(_safe_get(protection, "FormulaHidden")),
        },
    }


def _style_id(payload: dict[str, Any]) -> str:
    canonical = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:16]


def _collection_count(obj: object, attribute: str) -> int | None:
    try:
        collection = getattr(obj, attribute)
        if callable(collection):
            collection = collection()
        return int(collection.Count)
    except Exception:
        return None


def _collection(obj: object, attribute: str) -> list[object]:
    try:
        collection = getattr(obj, attribute)
        if callable(collection):
            collection = collection()
        count = int(collection.Count)
        result = []
        for index in range(1, count + 1):
            try:
                result.append(collection(index))
            except Exception:
                result.append(collection.Item(index))
        return result
    except Exception:
        return []


def resolve_snapshot_mode(requested: str, total_cells: int, max_cells: int) -> tuple[str, list[str]]:
    if requested not in {"auto", "inventory", "full"}:
        raise ValueError("Snapshot mode must be 'auto', 'inventory', or 'full'")
    if max_cells <= 0:
        raise ValueError("Snapshot cell limit must be positive")
    if requested == "full" and total_cells > max_cells:
        raise ValueError(
            f"Full snapshot requires {total_cells:,} cells, above the safe limit of "
            f"{max_cells:,}. Use inventory mode or raise the explicit limit."
        )
    if requested == "auto" and total_cells > max_cells:
        return "inventory", [
            f"Full cell snapshot skipped because {total_cells:,} cells exceed the "
            f"automatic limit of {max_cells:,}; dimensions and object inventory remain captured."
        ]
    return ("full" if requested == "auto" else requested), []


def _defined_names(workbook: object) -> list[dict[str, Any]]:
    return [
        {
            "name": json_safe(_safe_get(item, "Name")),
            "refers_to": json_safe(_safe_get(item, "RefersTo")),
            "visible": json_safe(_safe_get(item, "Visible")),
            "macro_type": json_safe(_safe_get(item, "MacroType")),
            "category": json_safe(_safe_get(item, "Category")),
        }
        for item in _collection(workbook, "Names")
    ]


def _connections(workbook: object) -> list[dict[str, Any]]:
    return [
        {
            "name": json_safe(_safe_get(item, "Name")),
            "type": json_safe(_safe_get(item, "Type")),
            "description": json_safe(_safe_get(item, "Description")),
            "refresh_with_refresh_all": json_safe(_safe_get(item, "RefreshWithRefreshAll")),
        }
        for item in _collection(workbook, "Connections")
    ]


def _object_range_address(obj: object, attribute: str) -> Any:
    target = _safe_get(obj, attribute)
    return json_safe(_safe_get(target, "Address"))


def _shape_inventory(sheet: object) -> list[dict[str, Any]]:
    return [
        {
            "name": json_safe(_safe_get(item, "Name")),
            "type": json_safe(_safe_get(item, "Type")),
            "left": json_safe(_safe_get(item, "Left")),
            "top": json_safe(_safe_get(item, "Top")),
            "width": json_safe(_safe_get(item, "Width")),
            "height": json_safe(_safe_get(item, "Height")),
            "visible": json_safe(_safe_get(item, "Visible")),
            "alternative_text": json_safe(_safe_get(item, "AlternativeText")),
            "on_action": json_safe(_safe_get(item, "OnAction")),
        }
        for item in _collection(sheet, "Shapes")
    ]


def _chart_inventory(sheet: object) -> list[dict[str, Any]]:
    result = []
    for item in _collection(sheet, "ChartObjects"):
        chart = _safe_get(item, "Chart")
        title = _safe_get(chart, "ChartTitle")
        result.append({
            "name": json_safe(_safe_get(item, "Name")),
            "left": json_safe(_safe_get(item, "Left")),
            "top": json_safe(_safe_get(item, "Top")),
            "width": json_safe(_safe_get(item, "Width")),
            "height": json_safe(_safe_get(item, "Height")),
            "chart_type": json_safe(_safe_get(chart, "ChartType")),
            "has_title": json_safe(_safe_get(chart, "HasTitle")),
            "title": json_safe(_safe_get(title, "Text")),
        })
    return result


def _table_inventory(sheet: object) -> list[dict[str, Any]]:
    return [
        {
            "name": json_safe(_safe_get(item, "Name")),
            "display_name": json_safe(_safe_get(item, "DisplayName")),
            "range": _object_range_address(item, "Range"),
            "show_headers": json_safe(_safe_get(item, "ShowHeaders")),
            "show_totals": json_safe(_safe_get(item, "ShowTotals")),
            "table_style": json_safe(_safe_get(item, "TableStyle")),
        }
        for item in _collection(sheet, "ListObjects")
    ]


def _pivot_inventory(sheet: object) -> list[dict[str, Any]]:
    return [
        {
            "name": json_safe(_safe_get(item, "Name")),
            "source_data": json_safe(_safe_get(item, "SourceData")),
            "table_range": _object_range_address(item, "TableRange2"),
            "refresh_on_open": json_safe(_safe_get(item, "RefreshOnFileOpen")),
        }
        for item in _collection(sheet, "PivotTables")
    ]


def _conditional_formats(used: object) -> list[dict[str, Any]]:
    result = []
    for item in _collection(used, "FormatConditions"):
        result.append({
            "type": json_safe(_safe_get(item, "Type")),
            "operator": json_safe(_safe_get(item, "Operator")),
            "formula1": json_safe(_safe_get(item, "Formula1")),
            "formula2": json_safe(_safe_get(item, "Formula2")),
            "priority": json_safe(_safe_get(item, "Priority")),
            "stop_if_true": json_safe(_safe_get(item, "StopIfTrue")),
            "font_color": json_safe(_safe_get(_safe_get(item, "Font"), "Color")),
            "fill_color": json_safe(_safe_get(_safe_get(item, "Interior"), "Color")),
        })
    return result


def _sheet_inventory(sheet: object) -> dict[str, Any]:
    used = sheet.UsedRange
    first_row = int(used.Row)
    first_col = int(used.Column)
    row_count = int(used.Rows.Count)
    col_count = int(used.Columns.Count)
    rows = []
    columns = []
    for row in range(first_row, first_row + row_count):
        row_range = sheet.Rows(row)
        rows.append({"row": row, "height": json_safe(_safe_get(row_range, "RowHeight")), "hidden": json_safe(_safe_get(row_range, "Hidden")), "outline_level": json_safe(_safe_get(row_range, "OutlineLevel"))})
    for col in range(first_col, first_col + col_count):
        column_range = sheet.Columns(col)
        columns.append({"column": col, "width": json_safe(_safe_get(column_range, "ColumnWidth")), "hidden": json_safe(_safe_get(column_range, "Hidden")), "outline_level": json_safe(_safe_get(column_range, "OutlineLevel"))})
    return {
        "name": str(sheet.Name),
        "visible": json_safe(_safe_get(sheet, "Visible")),
        "protected": bool(_safe_get(sheet, "ProtectContents", False)),
        "used_range": {
            "first_row": first_row,
            "first_column": first_col,
            "row_count": row_count,
            "column_count": col_count,
            "cell_count": row_count * col_count,
        },
        "merged_area_count": _collection_count(used, "MergeAreas"),
        "merged_areas": [json_safe(_safe_get(item, "Address")) for item in _collection(used, "MergeAreas")],
        "shape_count": _collection_count(sheet, "Shapes"),
        "shapes": _shape_inventory(sheet),
        "chart_count": _collection_count(sheet, "ChartObjects"),
        "charts": _chart_inventory(sheet),
        "table_count": _collection_count(sheet, "ListObjects"),
        "tables": _table_inventory(sheet),
        "pivot_count": _collection_count(sheet, "PivotTables"),
        "pivots": _pivot_inventory(sheet),
        "conditional_format_count": _collection_count(used, "FormatConditions"),
        "conditional_formats": _conditional_formats(used),
        "hyperlink_count": _collection_count(sheet, "Hyperlinks"),
        "comment_count": _collection_count(sheet, "Comments"),
        "auto_filter_mode": json_safe(_safe_get(sheet, "AutoFilterMode")),
        "filter_mode": json_safe(_safe_get(sheet, "FilterMode")),
        "rows": rows,
        "columns": columns,
        "page_setup": {
            "orientation": json_safe(_safe_get(_safe_get(sheet, "PageSetup"), "Orientation")),
            "print_area": json_safe(_safe_get(_safe_get(sheet, "PageSetup"), "PrintArea")),
            "print_title_rows": json_safe(_safe_get(_safe_get(sheet, "PageSetup"), "PrintTitleRows")),
            "print_title_columns": json_safe(_safe_get(_safe_get(sheet, "PageSetup"), "PrintTitleColumns")),
        },
    }


def _full_sheet_snapshot(sheet: object, inventory: dict[str, Any], styles: dict[str, Any]) -> dict[str, Any]:
    used_info = inventory["used_range"]
    first_row = int(used_info["first_row"])
    first_col = int(used_info["first_column"])
    row_count = int(used_info["row_count"])
    col_count = int(used_info["column_count"])
    cells: list[dict[str, Any]] = []

    for row in range(first_row, first_row + row_count):
        for col in range(first_col, first_col + col_count):
            cell = sheet.Cells(row, col)
            style = _style_payload(cell)
            sid = _style_id(style)
            styles.setdefault(sid, style)
            validation = _safe_get(cell, "Validation")
            comment = _safe_get(cell, "Comment")
            merge_area = _safe_get(cell, "MergeArea") if bool(_safe_get(cell, "MergeCells", False)) else None
            cells.append({
                "row": row,
                "column": col,
                "address": json_safe(_safe_get(cell, "Address")),
                "value": json_safe(_safe_get(cell, "Value2")),
                "formula": json_safe(_safe_get(cell, "Formula")),
                "data_type": json_safe(_safe_get(cell, "Value2").__class__.__name__ if _safe_get(cell, "Value2") is not None else "empty"),
                "style_id": sid,
                "has_formula": bool(_safe_get(cell, "HasFormula", False)),
                "merged": bool(_safe_get(cell, "MergeCells", False)),
                "merge_area": json_safe(_safe_get(merge_area, "Address")),
                "comment": json_safe(_safe_get(comment, "Text")),
                "validation": {
                    "type": json_safe(_safe_get(validation, "Type")),
                    "operator": json_safe(_safe_get(validation, "Operator")),
                    "formula1": json_safe(_safe_get(validation, "Formula1")),
                    "formula2": json_safe(_safe_get(validation, "Formula2")),
                    "ignore_blank": json_safe(_safe_get(validation, "IgnoreBlank")),
                    "in_cell_dropdown": json_safe(_safe_get(validation, "InCellDropdown")),
                },
            })
    return {"cells": cells}


def build_workbook_snapshot(workbook: object, path: Path, options: SnapshotOptions) -> dict[str, Any]:
    names = [
        str(workbook.Worksheets(index).Name)
        for index in range(1, int(workbook.Worksheets.Count) + 1)
    ]
    inventories = [_sheet_inventory(workbook.Worksheets(name)) for name in names]
    total_cells = sum(int(item["used_range"]["cell_count"]) for item in inventories)
    effective_mode, warnings = resolve_snapshot_mode(options.mode, total_cells, options.max_cells)

    styles: dict[str, Any] = {}
    sheets: list[dict[str, Any]] = []
    for inventory in inventories:
        item = dict(inventory)
        if effective_mode == "full":
            item.update(_full_sheet_snapshot(workbook.Worksheets(item["name"]), item, styles))
        sheets.append(item)
    fingerprint = collect_fingerprint(workbook, path)
    return {
        "schema_version": "1.0",
        "requested_mode": options.mode,
        "snapshot_mode": effective_mode,
        "created_utc": datetime.now().astimezone().isoformat(),
        "source": {
            "path": str(Path(path).resolve()),
            "sha256": fingerprint.sha256,
            "size_bytes": fingerprint.size_bytes,
            "file_format": fingerprint.file_format,
        },
        "workbook": {
            "sheet_count": fingerprint.sheet_count,
            "worksheet_count": len(names),
            "sheet_names": fingerprint.sheet_names,
            "defined_name_count": fingerprint.defined_name_count,
            "defined_names": _defined_names(workbook),
            "external_links": fingerprint.external_links,
            "connection_count": fingerprint.connection_count,
            "connections": _connections(workbook),
            "pivot_counts": fingerprint.pivot_counts,
            "has_vba_project": fingerprint.has_vba_project,
            "calculation_mode": json_safe(_safe_get(_safe_get(workbook, "Application"), "Calculation")),
            "date_1904": json_safe(_safe_get(workbook, "Date1904")),
            "read_only": json_safe(_safe_get(workbook, "ReadOnly")),
            "saved": json_safe(_safe_get(workbook, "Saved")),
        },
        "cell_count": total_cells,
        "styles": styles,
        "sheets": sheets,
        "warnings": warnings,
        "limitations": [
            "The byte-for-byte workbook backup is the recovery source of truth.",
            "Binary VBA, pivot caches, embedded objects, and chart internals are inventoried but not reconstructed from JSON alone.",
        ],
    }


def write_snapshot_json(snapshot: dict[str, Any], output_path: Path) -> Path:
    output_path = Path(output_path).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_suffix(output_path.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(snapshot, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temporary, output_path)
    return output_path
