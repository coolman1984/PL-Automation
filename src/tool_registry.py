"""Stable, machine-readable capability catalogue for an Excel AI agent."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ToolSpec:
    name: str
    category: str
    description: str
    status: str
    mutates_workbook: bool
    requires_excel: bool
    requires_backup: bool
    input_schema: dict[str, Any]
    risk: str = "low"
    requires_approval: bool = False
    safe_for_dry_run: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _path_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["file"],
        "properties": {"file": {"type": "string", "minLength": 1}},
        "additionalProperties": False,
    }


def _read_range_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["file", "sheet", "address"],
        "properties": {
            "file": {"type": "string", "minLength": 1},
            "sheet": {"type": "string", "minLength": 1},
            "address": {"type": "string", "minLength": 1},
            "mode": {"enum": ["auto", "attach", "open"]},
        },
        "additionalProperties": False,
    }


def _clear_range_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["expected_cell_count"],
        "properties": {
            "expected_cell_count": {"type": "integer", "minimum": 1},
        },
        "additionalProperties": False,
    }


def _write_range_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["values"],
        "properties": {
            "values": {"type": "array", "items": {"type": "array"}},
        },
        "additionalProperties": False,
    }


def _fill_formula_down_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["template", "expected_template_formulas", "expected_target_row_count"],
        "properties": {
            "template": {
                "type": "object",
                "required": ["workbook_id"],
                "properties": {
                    "workbook_id": {"type": "string"},
                    "sheet": {"type": "string"},
                    "address": {"type": "string"},
                },
            },
            "expected_template_formulas": {"type": "array", "items": {"type": "string"}, "minItems": 1},
            "expected_target_row_count": {"type": "integer", "minimum": 1},
        },
        "additionalProperties": False,
    }


def _insert_columns_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["count", "expected_anchor_column"],
        "properties": {
            "count": {"type": "integer", "minimum": 1},
            "expected_anchor_column": {"type": "string", "minLength": 1},
        },
        "additionalProperties": False,
    }


def _update_pivot_source_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["expected_current_source", "new_source"],
        "properties": {
            "expected_current_source": {"type": "string", "minLength": 1},
            "new_source": {"type": "string", "minLength": 1},
            "allow_shared_cache_replacement": {"type": "boolean"},
        },
        "additionalProperties": False,
    }


def _copy_range_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["source"],
        "properties": {
            "source": {
                "type": "object",
                "required": ["workbook_id"],
                "properties": {
                    "workbook_id": {"type": "string"},
                    "sheet": {"type": "string"},
                    "address": {"type": "string"},
                    "object_name": {"type": "string"},
                },
            },
            "mode": {"enum": ["all", "values", "formulas", "formats"]},
        },
        "additionalProperties": False,
    }


def _object_operation_schema(
    operations: tuple[str, ...], *, extra: dict[str, Any] | None = None
) -> dict[str, Any]:
    properties: dict[str, Any] = {
        "operation": {"enum": list(operations)},
    }
    properties.update(extra or {})
    return {
        "type": "object",
        "required": ["operation"],
        "properties": properties,
        "additionalProperties": False,
    }


def _format_range_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["format"],
        "properties": {
            "format": {
                "type": "object",
                "minProperties": 1,
                "additionalProperties": False,
                "properties": {
                    "font_name": {"type": "string"},
                    "font_size": {"type": "number", "minimum": 1, "maximum": 409},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "font_color": {"type": "integer", "minimum": 0, "maximum": 16777215},
                    "fill_color": {"type": "integer", "minimum": 0, "maximum": 16777215},
                    "number_format": {"type": "string"},
                    "horizontal_alignment": {"enum": ["general", "left", "center", "right"]},
                    "vertical_alignment": {"enum": ["top", "center", "bottom"]},
                    "wrap_text": {"type": "boolean"},
                    "row_height": {"type": "number", "minimum": 0, "maximum": 409.5},
                    "column_width": {"type": "number", "minimum": 0, "maximum": 255}
                }
            }
        },
        "additionalProperties": False,
    }


def _calculate_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "properties": {
            "full_rebuild": {"type": "boolean"},
            "timeout_seconds": {"type": "integer", "minimum": 1, "maximum": 600}
        },
        "additionalProperties": False,
    }


TOOLS: tuple[ToolSpec, ...] = (
    ToolSpec("list_tools", "system", "List every known tool and its readiness.", "available", False, False, False, {"type": "object", "additionalProperties": False}),
    ToolSpec("inspect_file", "safety", "Detect the Excel container, protection, complexity, and safest engine.", "available", False, False, False, _path_schema()),
    ToolSpec("create_backup", "safety", "Create and verify a byte-for-byte backup plus JSON evidence.", "available", False, False, False, _path_schema()),
    ToolSpec("snapshot_workbook", "safety", "Create a structured workbook inventory or full cell/style snapshot through Excel.", "available", False, True, False, _path_schema()),
    ToolSpec("prepare_workbook", "safety", "Back up, verify, open read-only, and snapshot before any edit.", "available", False, True, False, _path_schema()),
    ToolSpec("pnl_update_a08", "recipe", "Run the guarded August P&L update recipe.", "available", True, True, True, _path_schema(), requires_approval=True),
    ToolSpec("read_range", "cells", "Read values, formulas, and formats from an explicit range without changing the workbook.", "available", False, True, False, _read_range_schema(), risk="low", safe_for_dry_run=True),
    ToolSpec("clear_range", "cells", "Clear cell contents only (never formatting, never rows/columns) from an exact bounded range on a working copy.", "available", True, True, True, _clear_range_schema(), risk="low", requires_approval=True, safe_for_dry_run=True),
    ToolSpec("write_range", "cells", "Bulk-write values to a configured range on a working copy.", "available", True, True, True, _write_range_schema(), risk="low", requires_approval=True, safe_for_dry_run=True),
    ToolSpec("set_formula", "cells", "Set an exact rectangular formula matrix on a working copy.", "available", True, True, True, {"type": "object", "required": ["formulas"], "properties": {"formulas": {"type": "array", "items": {"type": "array"}}}, "additionalProperties": False}, requires_approval=True),
    ToolSpec("format_range", "formatting", "Apply a whitelisted format patch to one bounded range.", "available", True, True, True, _format_range_schema(), requires_approval=True),
    ToolSpec("copy_range", "cells", "Copy values, formulas, formats, or all content through Excel within the same open workbook.", "available", True, True, True, _copy_range_schema(), risk="low", requires_approval=True, safe_for_dry_run=True),
    ToolSpec("fill_formula_down", "cells", "Propagate a single-row formula template down through an exact, contiguous, same-column target row range using Excel's own fill semantics.", "available", True, True, True, _fill_formula_down_schema(), risk="low", requires_approval=True, safe_for_dry_run=True),
    ToolSpec("insert_rows", "structure", "Insert an exact bounded row count at an exact anchor.", "available", True, True, True, _object_operation_schema(("insert",), extra={"count": {"type": "integer", "minimum": 1, "maximum": 10000}, "expected_anchor_row": {"type": "integer", "minimum": 1}}), requires_approval=True),
    ToolSpec("insert_columns", "structure", "Insert an exact column count at an exact anchor on a working copy, inheriting neighboring formatting through Excel's own insert semantics.", "available", True, True, True, _insert_columns_schema(), risk="medium", requires_approval=True, safe_for_dry_run=True),
    ToolSpec("manage_sheet", "structure", "Create, rename, show, hide, or safely remove an empty sheet.", "available", True, True, True, _object_operation_schema(("create", "rename", "set_visibility", "delete_empty"), extra={"name": {"type": "string"}, "new_name": {"type": "string"}, "visibility": {"enum": ["visible", "hidden", "very_hidden"]}, "expected_empty": {"type": "boolean"}}), risk="medium", requires_approval=True),
    ToolSpec("manage_table", "objects", "Create, resize, or remove a named Excel Table while preserving cell data.", "available", True, True, True, _object_operation_schema(("create", "resize", "unlist"), extra={"name": {"type": "string"}, "expected_current_address": {"type": "string"}, "has_headers": {"type": "boolean"}}), risk="medium", requires_approval=True),
    ToolSpec("manage_filter", "objects", "Apply, clear, or remove an AutoFilter on one exact range.", "available", True, True, True, _object_operation_schema(("apply", "clear", "remove"), extra={"field": {"type": "integer", "minimum": 1}, "criteria1": {}, "criteria2": {}, "operator": {"type": "integer"}}), requires_approval=True),
    ToolSpec("manage_validation", "objects", "Set or remove data validation on one exact range.", "available", True, True, True, _object_operation_schema(("set", "delete"), extra={"validation_type": {"enum": ["list", "whole", "decimal", "date", "custom"]}, "formula1": {"type": "string"}, "formula2": {"type": "string"}, "operator": {"type": "integer"}, "replace": {"type": "boolean"}}), requires_approval=True),
    ToolSpec("manage_comment", "objects", "Set or remove a legacy cell note on one exact cell.", "available", True, True, True, _object_operation_schema(("set", "delete"), extra={"text": {"type": "string"}, "expected_current_text": {"type": ["string", "null"]}}), requires_approval=True),
    ToolSpec("manage_hyperlink", "objects", "Set or remove a safe web, mail, or internal hyperlink on one exact cell.", "available", True, True, True, _object_operation_schema(("set", "delete"), extra={"address": {"type": "string"}, "sub_address": {"type": "string"}, "display_text": {"type": "string"}, "expected_current_address": {"type": ["string", "null"]}}), requires_approval=True),
    ToolSpec("update_pivot_source", "objects", "Change one named PivotTable's worksheet/table source and perform a targeted refresh (never RefreshAll). External and Data Model sources stay locked.", "available", True, True, True, _update_pivot_source_schema(), risk="medium", requires_approval=True, safe_for_dry_run=True),
    ToolSpec("manage_chart", "objects", "Create, update, or delete one named chart from exact ranges.", "available", True, True, True, _object_operation_schema(("create", "update", "delete"), extra={"name": {"type": "string"}, "expected_exists": {"type": "boolean"}, "source_address": {"type": "string"}, "anchor_address": {"type": "string"}, "chart_type": {"enum": ["column", "bar", "line", "pie", "area", "scatter"]}, "title": {"type": "string"}, "width": {"type": "number"}, "height": {"type": "number"}}), risk="medium", requires_approval=True),
    ToolSpec("manage_name", "objects", "Create, update, or delete one exact workbook name with a reference fingerprint.", "available", True, True, True, _object_operation_schema(("set", "delete"), extra={"name": {"type": "string"}, "refers_to": {"type": "string"}, "expected_current_refers_to": {"type": ["string", "null"]}}), requires_approval=True),
    ToolSpec("manage_connection", "calculation", "Refresh one exact workbook connection with a bounded timeout.", "available", True, True, True, _object_operation_schema(("refresh",), extra={"name": {"type": "string"}, "timeout_seconds": {"type": "integer", "minimum": 1, "maximum": 600}}), risk="medium", requires_approval=True),
    ToolSpec("refresh_workbook", "calculation", "Refresh only explicitly named connections and PivotTables in order.", "available", True, True, True, {"type": "object", "required": ["connection_names", "pivot_tables"], "properties": {"connection_names": {"type": "array", "items": {"type": "string"}}, "pivot_tables": {"type": "array", "items": {"type": "object"}}, "timeout_seconds": {"type": "integer", "minimum": 1, "maximum": 600}}, "additionalProperties": False}, risk="medium", requires_approval=True),
    ToolSpec("calculate_workbook", "calculation", "Calculate in Excel and wait for completion with a timeout.", "available", True, True, True, _calculate_schema(), requires_approval=True),
    ToolSpec("validate_workbook", "safety", "Validate expected sheets, required objects, and formula-error health.", "available", False, True, True, {"type": "object", "properties": {"expected_sheet_names": {"type": "array", "items": {"type": "string"}}, "require_no_formula_errors": {"type": "boolean"}, "required_tables": {"type": "array", "items": {"type": "object"}}, "required_charts": {"type": "array", "items": {"type": "object"}}, "required_names": {"type": "array", "items": {"type": "string"}}}, "additionalProperties": False}),
    ToolSpec("publish_workbook", "safety", "Atomically publish only a closed and validated working copy.", "planned", False, False, True, {}),
    ToolSpec("restore_backup", "safety", "Restore a selected verified backup to a new recovery path.", "planned", False, False, False, {}),
)


def tool_catalog(*, include_planned: bool = True) -> dict[str, Any]:
    selected = TOOLS if include_planned else tuple(item for item in TOOLS if item.status == "available")
    return {
        "schema_version": "1.0",
        "tool_count": len(selected),
        "available_count": sum(item.status == "available" for item in selected),
        "tools": [item.to_dict() for item in selected],
    }


def describe_tool(name: str) -> dict[str, Any] | None:
    """Return one tool definition without exposing executable internals."""
    for item in TOOLS:
        if item.name == name:
            return item.to_dict()
    return None


def write_catalog_json(path: Path, *, include_planned: bool = True) -> Path:
    """Write the generated catalogue for agents and repository tooling."""
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(tool_catalog(include_planned=include_planned), handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")
        handle.flush()
    temporary.replace(destination)
    return destination
