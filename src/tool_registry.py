"""Stable, machine-readable capability catalogue for an Excel AI agent."""

from __future__ import annotations

from dataclasses import asdict, dataclass
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

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _path_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "required": ["file"],
        "properties": {"file": {"type": "string", "minLength": 1}},
        "additionalProperties": False,
    }


TOOLS: tuple[ToolSpec, ...] = (
    ToolSpec("list_tools", "system", "List every known tool and its readiness.", "available", False, False, False, {"type": "object", "additionalProperties": False}),
    ToolSpec("inspect_file", "safety", "Detect the Excel container, protection, complexity, and safest engine.", "available", False, False, False, _path_schema()),
    ToolSpec("create_backup", "safety", "Create and verify a byte-for-byte backup plus JSON evidence.", "available", False, False, False, _path_schema()),
    ToolSpec("snapshot_workbook", "safety", "Create a structured workbook inventory or full cell/style snapshot through Excel.", "available", False, True, False, _path_schema()),
    ToolSpec("prepare_workbook", "safety", "Back up, verify, open read-only, and snapshot before any edit.", "available", False, True, False, _path_schema()),
    ToolSpec("pnl_update_a08", "recipe", "Run the guarded August P&L update recipe.", "available", True, True, True, _path_schema()),
    ToolSpec("read_range", "cells", "Read values, formulas, and formats from a configured range.", "planned", False, True, False, {}),
    ToolSpec("write_range", "cells", "Bulk-write values to a configured range on a working copy.", "planned", True, True, True, {}),
    ToolSpec("set_formula", "cells", "Set one or more formulas on a working copy.", "planned", True, True, True, {}),
    ToolSpec("format_range", "formatting", "Apply fonts, fills, borders, alignment, and number formats.", "planned", True, True, True, {}),
    ToolSpec("copy_range", "cells", "Copy values, formulas, formats, or all content through Excel.", "planned", True, True, True, {}),
    ToolSpec("insert_rows", "structure", "Insert rows while preserving neighboring workbook behavior.", "planned", True, True, True, {}),
    ToolSpec("insert_columns", "structure", "Insert columns while preserving neighboring workbook behavior.", "planned", True, True, True, {}),
    ToolSpec("manage_sheet", "structure", "Create, rename, move, hide, or safely delete sheets.", "planned", True, True, True, {}),
    ToolSpec("manage_table", "objects", "Create, resize, or update an Excel table.", "planned", True, True, True, {}),
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
