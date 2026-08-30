"""Bounded executor for advanced Excel features.

The agent can select only the named tools below.  Validation happens before
the engine is called; the engine never receives free-form Python or COM code.
"""

from __future__ import annotations

import time
from collections.abc import Mapping
from typing import Any
from urllib.parse import urlparse

from .agent_contracts import TargetRef, ToolError, ToolMetrics, ToolRequest, ToolResult


ADVANCED_TOOLS = frozenset(
    {
        "format_range",
        "insert_rows",
        "manage_sheet",
        "manage_table",
        "manage_filter",
        "manage_validation",
        "manage_comment",
        "manage_hyperlink",
        "manage_chart",
        "manage_name",
        "manage_connection",
        "refresh_workbook",
        "calculate_workbook",
        "validate_workbook",
    }
)

_MAX_CELLS = 200_000
_MAX_INSERT_ROWS = 10_000
_FORMAT_KEYS = frozenset(
    {
        "font_name",
        "font_size",
        "bold",
        "italic",
        "font_color",
        "fill_color",
        "number_format",
        "horizontal_alignment",
        "vertical_alignment",
        "wrap_text",
        "row_height",
        "column_width",
    }
)


def _failure(tool: str, code: str, message: str, *, details: dict[str, Any] | None = None) -> ToolResult:
    return ToolResult.failure(
        tool,
        ToolError(code=code, message=message, details=details or {}, recoverable=True),
    )


def _working_target(request: ToolRequest, *, address: bool = False, object_name: bool = False) -> ToolResult | None:
    target = request.target
    if target is None or target.workbook_id not in {"working-copy", ""}:
        return _failure(request.tool, "invalid_target_workbook", "This operation may only use the approved working copy.")
    if not target.sheet:
        return _failure(request.tool, "sheet_required", "An exact worksheet name is required.")
    if address and not target.address:
        return _failure(request.tool, "address_required", "An exact bounded range address is required.")
    if object_name and not target.object_name:
        return _failure(request.tool, "object_name_required", "An exact object name is required.")
    return None


def _bounded(engine: Any, target: TargetRef) -> tuple[int, int, int]:
    first_row, first_col, last_row, last_col = engine.validate_bounded_range(target)
    rows = last_row - first_row + 1
    columns = last_col - first_col + 1
    cells = rows * columns
    if cells < 1 or cells > _MAX_CELLS:
        raise ValueError(f"Range contains {cells} cells; the limit is {_MAX_CELLS}")
    return rows, columns, cells


def _require_operation(request: ToolRequest, allowed: set[str]) -> tuple[str | None, ToolResult | None]:
    operation = request.arguments.get("operation")
    if not isinstance(operation, str) or operation not in allowed:
        return None, _failure(
            request.tool,
            "invalid_operation",
            f"arguments.operation must be one of {sorted(allowed)}.",
        )
    return operation, None


def _validate_request(request: ToolRequest, engine: Any) -> ToolResult | None:
    tool = request.tool
    args = request.arguments

    if tool == "validate_workbook":
        if not args:
            return _failure(tool, "validation_checks_required", "At least one workbook validation check is required.")
        return None
    if tool in {"calculate_workbook", "refresh_workbook"}:
        if request.target is not None and request.target.workbook_id not in {"working-copy", ""}:
            return _failure(tool, "invalid_target_workbook", "Workbook operation may only use the approved working copy.")
        timeout = args.get("timeout_seconds", 120)
        if not isinstance(timeout, int) or isinstance(timeout, bool) or not 1 <= timeout <= 600:
            return _failure(tool, "invalid_timeout", "timeout_seconds must be an integer from 1 to 600.")
        if tool == "calculate_workbook" and not isinstance(args.get("full_rebuild", False), bool):
            return _failure(tool, "invalid_calculation_mode", "full_rebuild must be true or false.")
        if tool == "refresh_workbook":
            connections = args.get("connection_names")
            pivots = args.get("pivot_tables")
            if not isinstance(connections, list) or not all(isinstance(name, str) and name for name in connections):
                return _failure(tool, "invalid_connection_list", "connection_names must be an explicit list of names.")
            if not isinstance(pivots, list) or not all(
                isinstance(item, Mapping)
                and isinstance(item.get("sheet"), str) and item.get("sheet")
                and isinstance(item.get("name"), str) and item.get("name")
                for item in pivots
            ):
                return _failure(tool, "invalid_pivot_list", "pivot_tables must contain exact sheet/name objects.")
        return None

    if tool == "format_range":
        error = _working_target(request, address=True)
        if error:
            return error
        patch = args.get("format")
        if not isinstance(patch, Mapping) or not patch or set(patch) - _FORMAT_KEYS:
            return _failure(tool, "invalid_format_patch", "Only the declared non-empty format properties are allowed.")
        boolean_keys = {"bold", "italic", "wrap_text"}
        numeric_limits = {
            "font_size": (1, 409), "font_color": (0, 16777215),
            "fill_color": (0, 16777215), "row_height": (0, 409.5),
            "column_width": (0, 255),
        }
        if any(key in patch and not isinstance(patch[key], bool) for key in boolean_keys):
            return _failure(tool, "invalid_format_value", "Boolean format properties must be true or false.")
        for key, (minimum, maximum) in numeric_limits.items():
            value = patch.get(key)
            if key in patch and (not isinstance(value, (int, float)) or isinstance(value, bool) or not minimum <= value <= maximum):
                return _failure(tool, "invalid_format_value", f"{key} is outside its supported range.")
        for key in {"font_name", "number_format"}:
            if key in patch and not isinstance(patch[key], str):
                return _failure(tool, "invalid_format_value", f"{key} must be text.")
        if patch.get("horizontal_alignment", "general") not in {"general", "left", "center", "right"}:
            return _failure(tool, "invalid_format_value", "Unsupported horizontal alignment.")
        if patch.get("vertical_alignment", "top") not in {"top", "center", "bottom"}:
            return _failure(tool, "invalid_format_value", "Unsupported vertical alignment.")
        _bounded(engine, request.target)
        return None

    if tool == "insert_rows":
        error = _working_target(request, address=True)
        if error:
            return error
        operation, error = _require_operation(request, {"insert"})
        if error:
            return error
        count = args.get("count")
        expected = args.get("expected_anchor_row")
        if not isinstance(count, int) or isinstance(count, bool) or not 1 <= count <= _MAX_INSERT_ROWS:
            return _failure(tool, "invalid_row_count", f"Row count must be 1..{_MAX_INSERT_ROWS}.")
        first_row, _, _, _ = engine.resolve_bounds(request.target)
        if not isinstance(expected, int) or expected != first_row:
            return _failure(tool, "anchor_row_mismatch", f"Resolved anchor row is {first_row}, expected {expected!r}.")
        return None

    if tool == "manage_sheet":
        error = _working_target(request)
        if error:
            return error
        operation, error = _require_operation(request, {"create", "rename", "set_visibility", "delete_empty"})
        if error:
            return error
        if operation == "create" and args.get("name") != request.target.sheet:
            return _failure(tool, "sheet_name_mismatch", "Create name must exactly match target.sheet.")
        if operation == "rename" and (not isinstance(args.get("new_name"), str) or not args["new_name"]):
            return _failure(tool, "new_name_required", "Rename requires arguments.new_name.")
        if operation == "set_visibility" and args.get("visibility") not in {"visible", "hidden", "very_hidden"}:
            return _failure(tool, "visibility_required", "A declared visibility is required.")
        if operation == "delete_empty" and args.get("expected_empty") is not True:
            return _failure(tool, "empty_ack_required", "Deleting a sheet requires expected_empty=true.")
        return None

    if tool == "manage_table":
        error = _working_target(request, address=True, object_name=True)
        if error:
            return error
        operation, error = _require_operation(request, {"create", "resize", "unlist"})
        if error:
            return error
        if operation in {"resize", "unlist"} and "expected_current_address" not in args:
            return _failure(tool, "table_fingerprint_required", "Resize and unlist require expected_current_address.")
        _bounded(engine, request.target)
        return None

    if tool in {"manage_filter", "manage_validation"}:
        error = _working_target(request, address=True)
        if error:
            return error
        allowed = {"apply", "clear", "remove"} if tool == "manage_filter" else {"set", "delete"}
        operation, error = _require_operation(request, allowed)
        if error:
            return error
        rows, columns, _ = _bounded(engine, request.target)
        if tool == "manage_filter" and operation == "apply":
            field = args.get("field")
            if not isinstance(field, int) or isinstance(field, bool) or not 1 <= field <= columns:
                return _failure(tool, "filter_field_out_of_range", f"Filter field must be between 1 and {columns}.")
            if "criteria1" not in args:
                return _failure(tool, "criteria_required", "Applying a filter requires criteria1.")
        if tool == "manage_validation" and operation == "set":
            if args.get("validation_type") not in {"list", "whole", "decimal", "date", "custom"}:
                return _failure(tool, "validation_type_required", "A supported validation_type is required.")
            if not isinstance(args.get("formula1"), str) or not args["formula1"]:
                return _failure(tool, "formula1_required", "Setting validation requires formula1.")
        return None

    if tool in {"manage_comment", "manage_hyperlink"}:
        error = _working_target(request, address=True)
        if error:
            return error
        operation, error = _require_operation(request, {"set", "delete"})
        if error:
            return error
        _, _, cells = _bounded(engine, request.target)
        if cells != 1:
            return _failure(tool, "single_cell_required", "This operation requires exactly one cell.")
        expected_key = "expected_current_text" if tool == "manage_comment" else "expected_current_address"
        if expected_key not in args:
            return _failure(tool, "current_fingerprint_required", f"{expected_key} is required, including null for an empty cell.")
        if tool == "manage_comment" and operation == "set" and not isinstance(args.get("text"), str):
            return _failure(tool, "comment_text_required", "Setting a comment requires text.")
        if tool == "manage_hyperlink" and operation == "set":
            address = args.get("address", "")
            sub_address = args.get("sub_address", "")
            if not address and not sub_address:
                return _failure(tool, "hyperlink_destination_required", "A web/mail address or internal sub-address is required.")
            parsed = urlparse(str(address)) if address else None
            if address and parsed.scheme.casefold() not in {"http", "https", "mailto"}:
                return _failure(tool, "unsafe_hyperlink_scheme", "Only http, https, and mailto links are allowed.")
            if parsed and parsed.scheme.casefold() in {"http", "https"} and not parsed.netloc:
                return _failure(tool, "invalid_hyperlink", "Web links require a host name.")
        return None

    if tool == "manage_chart":
        error = _working_target(request)
        if error:
            return error
        operation, error = _require_operation(request, {"create", "update", "delete"})
        if error:
            return error
        if not isinstance(args.get("name"), str) or not args["name"]:
            return _failure(tool, "chart_name_required", "An exact chart name is required.")
        if not isinstance(args.get("expected_exists"), bool):
            return _failure(tool, "chart_fingerprint_required", "expected_exists must explicitly declare whether the chart already exists.")
        if (operation == "create") == args["expected_exists"]:
            return _failure(tool, "chart_operation_mismatch", "Create requires expected_exists=false; update/delete require true.")
        if operation in {"create", "update"}:
            if args.get("chart_type") not in {"column", "bar", "line", "pie", "area", "scatter"}:
                return _failure(tool, "chart_type_required", "A supported chart type is required.")
            if not isinstance(args.get("source_address"), str):
                return _failure(tool, "chart_source_required", "An exact source_address is required.")
            source = TargetRef("working-copy", sheet=request.target.sheet, address=args["source_address"])
            _bounded(engine, source)
            if "anchor_address" in args:
                anchor = TargetRef("working-copy", sheet=request.target.sheet, address=args["anchor_address"])
                _, _, anchor_cells = _bounded(engine, anchor)
                if anchor_cells != 1:
                    return _failure(tool, "chart_anchor_must_be_cell", "anchor_address must resolve to one cell.")
            for dimension in {"width", "height"}:
                value = args.get(dimension, 1)
                if not isinstance(value, (int, float)) or isinstance(value, bool) or value <= 0:
                    return _failure(tool, "invalid_chart_size", f"{dimension} must be positive.")
        return None

    if tool in {"manage_name", "manage_connection"}:
        if request.target is not None and request.target.workbook_id not in {"working-copy", ""}:
            return _failure(tool, "invalid_target_workbook", "This operation may only use the approved working copy.")
        allowed = {"set", "delete"} if tool == "manage_name" else {"refresh"}
        operation, error = _require_operation(request, allowed)
        if error:
            return error
        if not isinstance(args.get("name"), str) or not args["name"]:
            return _failure(tool, "name_required", "An exact name is required.")
        if tool == "manage_connection":
            timeout = args.get("timeout_seconds", 120)
            if not isinstance(timeout, int) or isinstance(timeout, bool) or not 1 <= timeout <= 600:
                return _failure(tool, "invalid_timeout", "timeout_seconds must be an integer from 1 to 600.")
        if tool == "manage_name" and "expected_current_refers_to" not in args:
            return _failure(tool, "current_fingerprint_required", "expected_current_refers_to is required, including null when creating a name.")
        if tool == "manage_name" and operation == "set" and not isinstance(args.get("refers_to"), str):
            return _failure(tool, "refers_to_required", "Setting a name requires refers_to.")
        return None

    return _failure(tool, "advanced_handler_missing", f"No advanced validator exists for {tool}.")


def execute_advanced_tool(request: ToolRequest, *, engine: Any | None) -> ToolResult | None:
    """Execute an advanced declared tool, or return ``None`` for other tools."""
    if request.tool not in ADVANCED_TOOLS:
        return None
    if engine is None:
        return _failure(request.tool, "engine_required", "An open compatible workbook engine is required.")
    started = time.perf_counter()
    try:
        validation_error = _validate_request(request, engine)
    except Exception as exc:
        return _failure(request.tool, "precondition_failed", str(exc))
    if validation_error is not None:
        return validation_error

    try:
        before = engine.inspect_advanced(request.tool, request.target, request.arguments)
    except Exception as exc:
        return _failure(request.tool, "before_fingerprint_failed", str(exc))

    fingerprint_fields = {
        "manage_table": ("expected_current_address", "address"),
        "manage_comment": ("expected_current_text", "text"),
        "manage_hyperlink": ("expected_current_address", "address"),
        "manage_name": ("expected_current_refers_to", "refers_to"),
    }
    fingerprint = fingerprint_fields.get(request.tool)
    if fingerprint and fingerprint[0] in request.arguments:
        current = before.get("current") or {}
        actual = current.get(fingerprint[1])
        expected = request.arguments[fingerprint[0]]
        if actual != expected:
            return _failure(
                request.tool,
                "current_fingerprint_mismatch",
                f"Current {fingerprint[1]} does not match the declared expectation.",
                details={"expected": expected, "actual": actual},
            )
    if request.tool in {"manage_chart", "manage_connection"}:
        exists = before.get("current") is not None
        expected_exists = request.arguments.get("expected_exists", True)
        if exists != expected_exists:
            return _failure(
                request.tool,
                "object_existence_mismatch",
                "The object's current existence does not match the declared expectation.",
                details={"expected_exists": expected_exists, "actual_exists": exists},
            )

    if request.tool == "validate_workbook":
        passed = bool(before.get("passed"))
        if not passed:
            return _failure(request.tool, "workbook_validation_failed", "Workbook validation checks failed.", details=before)
        return ToolResult.success(request.tool, after_evidence=before)

    if request.dry_run:
        return ToolResult.success(
            request.tool,
            before_evidence=before,
            after_evidence={"would_apply": dict(request.arguments)},
            warnings=("Dry run: workbook was not changed.",),
        )

    try:
        after = engine.execute_advanced(request.tool, request.target, request.arguments)
    except Exception as exc:
        return _failure(request.tool, "engine_operation_failed", str(exc))
    return ToolResult.success(
        request.tool,
        changed=True,
        before_evidence=before,
        after_evidence=after,
        metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
    )
