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


TOOLS: tuple[ToolSpec, ...] = (
    ToolSpec("list_tools", "system", "List every known tool and its readiness.", "available", False, False, False, {"type": "object", "additionalProperties": False}),
    ToolSpec("inspect_file", "safety", "Detect the Excel container, protection, complexity, and safest engine.", "available", False, False, False, _path_schema()),
    ToolSpec("create_backup", "safety", "Create and verify a byte-for-byte backup plus JSON evidence.", "available", False, False, False, _path_schema()),
    ToolSpec("snapshot_workbook", "safety", "Create a structured workbook inventory or full cell/style snapshot through Excel.", "available", False, True, False, _path_schema()),
    ToolSpec("prepare_workbook", "safety", "Back up, verify, open read-only, and snapshot before any edit.", "available", False, True, False, _path_schema()),
    ToolSpec("pnl_update_a08", "recipe", "Run the guarded August P&L update recipe.", "available", True, True, True, _path_schema()),
    ToolSpec("read_range", "cells", "Read values, formulas, and formats from an explicit range without changing the workbook.", "available", False, True, False, _read_range_schema(), risk="low", safe_for_dry_run=True),
    ToolSpec("clear_range", "cells", "Clear cell contents only (never formatting, never rows/columns) from an exact bounded range on a working copy.", "available", True, True, True, _clear_range_schema(), risk="low", safe_for_dry_run=True),
    ToolSpec("write_range", "cells", "Bulk-write values to a configured range on a working copy.", "available", True, True, True, _write_range_schema(), risk="low", safe_for_dry_run=True),
    ToolSpec("set_formula", "cells", "Set one or more formulas on a working copy.", "planned", True, True, True, {}),
    ToolSpec("format_range", "formatting", "Apply fonts, fills, borders, alignment, and number formats.", "planned", True, True, True, {}),
    ToolSpec("copy_range", "cells", "Copy values, formulas, formats, or all content through Excel within the same open workbook.", "available", True, True, True, _copy_range_schema(), risk="low", safe_for_dry_run=True),
    ToolSpec("fill_formula_down", "cells", "Propagate a single-row formula template down through an exact, contiguous, same-column target row range using Excel's own fill semantics.", "available", True, True, True, _fill_formula_down_schema(), risk="low", safe_for_dry_run=True),
    ToolSpec("insert_rows", "structure", "Insert rows while preserving neighboring workbook behavior.", "planned", True, True, True, {}),
    ToolSpec("insert_columns", "structure", "Insert an exact column count at an exact anchor on a working copy, inheriting neighboring formatting through Excel's own insert semantics.", "available", True, True, True, _insert_columns_schema(), risk="medium", safe_for_dry_run=True),
    ToolSpec("manage_sheet", "structure", "Create, rename, move, hide, or safely delete sheets.", "planned", True, True, True, {}),
    ToolSpec("manage_table", "objects", "Create, resize, or update an Excel table.", "planned", True, True, True, {}),
    ToolSpec("update_pivot_source", "objects", "Change one named PivotTable's worksheet/table source and perform a targeted refresh (never RefreshAll). External and Data Model sources stay locked.", "available", True, True, True, _update_pivot_source_schema(), risk="medium", safe_for_dry_run=True),
    ToolSpec("manage_chart", "objects", "Create or modify a chart using Excel's object model.", "planned", True, True, True, {}),
    ToolSpec("refresh_workbook", "calculation", "Refresh approved connections in a controlled order.", "planned", True, True, True, {}),
    ToolSpec("calculate_workbook", "calculation", "Calculate and wait for Excel completion with a timeout.", "planned", True, True, True, {}),
    ToolSpec("validate_workbook", "safety", "Compare required workbook facts and formulas before publication.", "planned", False, True, True, {}),
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
