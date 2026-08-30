"""Deterministic dispatcher for the declared agent tools.

Only tools with an actual handler can execute.  Planned catalogue entries are
explicitly rejected so an AI agent cannot accidentally invent capabilities.
"""

from __future__ import annotations

import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping

from .agent_contracts import TargetRef, ToolError, ToolRequest, ToolResult, ToolMetrics
from .backup_bundle import create_backup_bundle
from .file_probe import probe_excel_file
from .tool_registry import tool_catalog
from .file_probe import resolve_com_mode
from .advanced_tools import execute_advanced_tool

# No chunking is implemented yet (V2 plan section 8: "Add configurable
# chunking only after measuring the real payload"). This bounds a single
# bulk range operation instead of silently attempting an unbounded one.
_MAX_CELLS_PER_RANGE_OPERATION = 200_000
_MAX_COLUMNS_PER_INSERT = 256


def _error(tool: str, code: str, message: str, *, recoverable: bool = True, action: str | None = None, details: Mapping[str, Any] | None = None) -> ToolResult:
    return ToolResult.failure(
        tool,
        ToolError(
            code=code,
            message=message,
            recoverable=recoverable,
            details=dict(details or {}),
            suggested_action=action,
        ),
    )


def _tool_spec(tool: str) -> dict[str, Any] | None:
    return next((item for item in tool_catalog()["tools"] if item["name"] == tool), None)


def _file_from_request(request: ToolRequest) -> Path | None:
    value = request.arguments.get("file")
    if value is None and request.target is not None:
        value = request.target.workbook_id
    if not isinstance(value, str) or not value.strip():
        return None
    return Path(value).expanduser().resolve()


def _target_label(target: TargetRef | None) -> str:
    if target is None:
        return ""
    pieces = [target.sheet, target.address, target.object_name]
    return "!".join(item for item in pieces if item)


def _column_letter(number: int) -> str:
    letters = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _shape(values: Any) -> tuple[int, int] | None:
    if not isinstance(values, (list, tuple)):
        return None
    rows = [row for row in values if isinstance(row, (list, tuple))]
    if len(rows) != len(values):
        return None
    width = len(rows[0]) if rows else 0
    if any(len(row) != width for row in rows):
        return None
    return len(rows), width


def _snapshot_with_excel(
    source_path: Path,
    *,
    artifact_root: Path,
    mode: str,
    snapshot_mode: str,
    max_cells: int,
) -> dict[str, Any]:
    """Run the existing read-only snapshot path without exposing COM objects."""
    from .excel_session import ExcelSession
    from .workbook_snapshot import SnapshotOptions, build_workbook_snapshot, write_snapshot_json

    probe = probe_excel_file(source_path)
    if not probe.recognized:
        raise ValueError("Workbook format is not recognized safely")
    effective_mode = resolve_com_mode(mode, probe)
    if effective_mode == "auto":
        effective_mode = "open"
    if probe.protection in {"nasca_drm", "office_encrypted"} and effective_mode != "attach":
        raise PermissionError("Protected workbook requires authorized Excel attach mode")
    if effective_mode not in {"attach", "open"}:
        raise ValueError(f"No safe Excel mode selected: {effective_mode}")

    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    output_dir = artifact_root / "snapshots" / f"{stamp}_{uuid.uuid4().hex[:8]}"
    output_dir.mkdir(parents=True, exist_ok=False)
    session = ExcelSession.attach(source_path) if effective_mode == "attach" else ExcelSession.create(visible=False)
    with session:
        workbook = getattr(session, "source_workbook", None)
        if workbook is None:
            workbook = session.open_workbook(source_path, read_only=True, update_links=False)
        if effective_mode == "attach" and not bool(workbook.Saved):
            raise ValueError("Attached workbook has unsaved changes; save it before snapshotting")
        snapshot = build_workbook_snapshot(workbook, source_path, SnapshotOptions(mode=snapshot_mode, max_cells=max_cells))
        snapshot_path = write_snapshot_json(snapshot, output_dir / "workbook_snapshot.json")
    return {
        "snapshot_file": str(snapshot_path),
        "snapshot_mode": snapshot["snapshot_mode"],
        "cell_count": snapshot["cell_count"],
        "style_count": len(snapshot["styles"]),
        "engine": "excel_com",
        "mode": effective_mode,
    }


def _read_range_with_excel(
    source_path: Path,
    target: TargetRef,
    *,
    mode: str,
    include_formulas: bool,
) -> dict[str, Any]:
    """Open a workbook read-only and return JSON-safe range evidence."""
    from .engines.excel_com import ExcelComEngine
    from .excel_session import ExcelSession

    probe = probe_excel_file(source_path)
    if not probe.recognized:
        raise ValueError("Workbook format is not recognized safely")
    effective_mode = resolve_com_mode(mode, probe)
    if effective_mode == "auto":
        effective_mode = "open"
    if probe.protection in {"nasca_drm", "office_encrypted"} and effective_mode != "attach":
        raise PermissionError("Protected workbook requires authorized Excel attach mode")
    session = ExcelSession.attach(source_path) if effective_mode == "attach" else ExcelSession.create(visible=False)
    with session:
        workbook = getattr(session, "source_workbook", None)
        if workbook is None:
            workbook = session.open_workbook(source_path, read_only=True, update_links=False)
        engine = ExcelComEngine(workbook, session=None, workbook_id="source", read_only=True)
        result: dict[str, Any] = {
            "engine": "excel_com",
            "mode": effective_mode,
            "target": target.to_dict(),
            "values": engine.read_values(target),
            "workbook": engine.inspect(),
        }
        if include_formulas:
            result["formulas"] = engine.read_formulas(target)
        return result


def _require_target(request: ToolRequest) -> ToolResult | None:
    if request.target is None:
        return _error(
            request.tool,
            "target_required",
            "This tool needs an explicit workbook, sheet, and/or range target.",
            action="Resolve the exact target before calling the tool.",
        )
    return None


def _require_working_copy_target(request: ToolRequest) -> ToolResult | None:
    if request.target is None or request.target.workbook_id not in {"working-copy", ""}:
        return _error(
            request.tool,
            "invalid_target_workbook",
            f"{request.tool} may only target the Excel-created working copy, not the source.",
            action="Resolve a working-copy target through the approved coordinator transaction.",
        )
    return None


def _bounded_shape(engine: Any, target: TargetRef) -> tuple[int, int, int]:
    resolver = getattr(engine, "validate_bounded_range", engine.resolve_bounds)
    first_row, first_col, last_row, last_col = resolver(target)
    rows = last_row - first_row + 1
    columns = last_col - first_col + 1
    if rows < 1 or columns < 1:
        raise ValueError("Resolved range is empty")
    return rows, columns, rows * columns


def execute_tool(
    request: ToolRequest | Mapping[str, Any],
    *,
    project_root: Path | None = None,
    engine: Any | None = None,
) -> ToolResult:
    """Execute one declared request or return a stable, actionable error."""
    started = time.perf_counter()
    try:
        normalized = request if isinstance(request, ToolRequest) else ToolRequest.from_dict(request)
    except (TypeError, ValueError) as exc:
        return _error("request", "invalid_request", str(exc), recoverable=True)

    tool = normalized.tool
    spec = _tool_spec(tool)
    if spec is None:
        return _error(tool, "unknown_tool", f"Tool is not in the catalogue: {tool}", action="List tools and choose a declared capability.")
    if spec["status"] != "available":
        return _error(
            tool,
            "tool_not_available",
            f"Tool is declared but locked until its acceptance gates pass: {tool}",
            action="Choose an available tool or implement and test this capability first.",
        )

    if tool == "list_tools":
        return ToolResult.success(tool, after_evidence=tool_catalog())

    source_path = _file_from_request(normalized)
    if tool in {"inspect_file", "create_backup"}:
        if source_path is None:
            return _error(tool, "file_required", "A real local workbook path is required.", action="Set arguments.file to the exact source path.")
        if not source_path.exists() or not source_path.is_file():
            return _error(tool, "file_not_found", f"Workbook was not found: {source_path}", action="Check the path and permissions.")

    if tool == "inspect_file":
        try:
            probe = probe_excel_file(source_path)  # type: ignore[arg-type]
        except OSError as exc:
            return _error(tool, "file_probe_failed", str(exc), details={"path": str(source_path)})
        return ToolResult.success(
            tool,
            after_evidence={
                "path": str(source_path),
                "container": probe.container,
                "workbook_format": probe.workbook_format,
                "protection": probe.protection,
                "recognized": probe.recognized,
                "recommended_engine": probe.recommended_engine,
                "recommended_com_mode": probe.recommended_com_mode,
                "fast_read_candidate": probe.fast_read_candidate,
                "fast_edit_candidate": probe.fast_edit_candidate,
                "reason": probe.reason,
                "warnings": list(probe.warnings),
            },
            metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
        )

    if tool == "create_backup":
        if normalized.dry_run:
            return ToolResult.success(
                tool,
                warnings=("Dry run: no backup was created.",),
                after_evidence={"source": str(source_path), "would_create_backup": True},
                metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
            )
        root = project_root or Path.cwd()
        backup_root = Path(normalized.arguments.get("backup_root", root / "backups"))
        try:
            bundle = create_backup_bundle(source_path, backup_root, reason=normalized.arguments.get("reason", "agent_tool"))  # type: ignore[arg-type]
        except (OSError, ValueError) as exc:
            return _error(tool, "backup_failed", str(exc), details={"source": str(source_path)}, action="Keep the original unchanged and resolve the backup error.")
        return ToolResult.success(
            tool,
            changed=False,
            after_evidence=bundle.to_dict(),
            metrics=ToolMetrics(
                elapsed_ms=round((time.perf_counter() - started) * 1000),
                bytes_read=source_path.stat().st_size,
                bytes_written=bundle.backup_file.stat().st_size,
            ),
        )

    if tool in {"snapshot_workbook", "prepare_workbook"}:
        if source_path is None:
            return _error(tool, "file_required", "A real local workbook path is required.", action="Set arguments.file to the exact source path.")
        if normalized.dry_run:
            return ToolResult.success(
                tool,
                warnings=("Dry run: no backup or snapshot artifact was created.",),
                after_evidence={
                    "source": str(source_path),
                    "would_backup": tool == "prepare_workbook",
                    "would_snapshot": True,
                    "mode": normalized.arguments.get("mode", "auto"),
                },
                metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
            )
        root = Path(normalized.arguments.get("artifact_root", project_root or Path.cwd())).expanduser().resolve()
        try:
            bundle = None
            if tool == "prepare_workbook":
                bundle = create_backup_bundle(source_path, root / "backups", reason="agent_prepare")
            snapshot = _snapshot_with_excel(
                source_path,
                artifact_root=root,
                mode=str(normalized.arguments.get("mode", "auto")),
                snapshot_mode=str(normalized.arguments.get("snapshot_mode", "auto")),
                max_cells=int(normalized.arguments.get("max_snapshot_cells", 250000)),
            )
        except ImportError as exc:
            return _error(tool, "excel_dependency_missing", f"Excel execution dependencies are unavailable: {exc}", action="Run the private runtime self-check on Windows.")
        except (OSError, PermissionError, ValueError) as exc:
            return _error(tool, "preparation_failed", str(exc), details={"source": str(source_path)}, action="Keep the source unchanged and resolve the preparation error.")
        evidence: dict[str, Any] = {"snapshot": snapshot}
        if bundle is not None:
            evidence["backup"] = bundle.to_dict()
        return ToolResult.success(
            tool,
            after_evidence=evidence,
            metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
        )

    if tool == "pnl_update_a08":
        if source_path is None:
            return _error(tool, "file_required", "A real P&L workbook path is required.")
        try:
            from .config import load_config
            from .file_transaction import assert_source_candidate
            from .workflow import run_dry_run, run_execute

            config_path = Path(normalized.arguments.get("config", (project_root or Path.cwd()) / "config.yaml"))
            if not config_path.is_absolute():
                config_path = (project_root or Path.cwd()) / config_path
            config = load_config(config_path.resolve(), year=normalized.arguments.get("year"), month=normalized.arguments.get("month"), execution=not normalized.dry_run)
            assert_source_candidate(source_path)
            probe = probe_excel_file(source_path)
            effective_mode = resolve_com_mode(str(normalized.arguments.get("mode", "auto")), probe)
            exit_code = run_dry_run(source_path, config, effective_mode) if normalized.dry_run else run_execute(source_path, config, effective_mode, project_root or Path.cwd())
        except ImportError as exc:
            return _error(tool, "excel_dependency_missing", f"Excel execution dependencies are unavailable: {exc}", action="Run the private runtime self-check on Windows.")
        except Exception as exc:
            return _error(tool, "recipe_failed", str(exc), action="Review the recipe report and keep the original workbook unchanged.")
        if int(exit_code) != 0:
            return _error(tool, "recipe_not_ready", f"The P&L recipe stopped with exit code {exit_code}.", details={"exit_code": int(exit_code)})
        return ToolResult.success(tool, changed=not normalized.dry_run, after_evidence={"exit_code": int(exit_code), "dry_run": normalized.dry_run}, metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)))

    if tool == "read_range" and normalized.target is None:
        sheet = normalized.arguments.get("sheet")
        address = normalized.arguments.get("address")
        if source_path is not None and isinstance(sheet, str) and isinstance(address, str):
            normalized = ToolRequest(
                schema_version=normalized.schema_version,
                transaction_id=normalized.transaction_id,
                tool=normalized.tool,
                target=TargetRef("source", sheet=sheet, address=address),
                arguments=normalized.arguments,
                preconditions=normalized.preconditions,
                expected_effect=normalized.expected_effect,
                dry_run=normalized.dry_run,
            )

    if tool == "read_range" and engine is None:
        if source_path is None:
            return _error(tool, "file_required", "A real local workbook path is required.", action="Set arguments.file to the exact source path.")
        target_error = _require_target(normalized)
        if target_error is not None:
            return target_error
        try:
            evidence = _read_range_with_excel(
                source_path,
                normalized.target,  # type: ignore[arg-type]
                mode=str(normalized.arguments.get("mode", "auto")),
                include_formulas=bool(normalized.arguments.get("include_formulas", True)),
            )
        except ImportError as exc:
            return _error(tool, "excel_dependency_missing", f"Excel execution dependencies are unavailable: {exc}", action="Run the private runtime self-check on Windows.")
        except (OSError, PermissionError, ValueError) as exc:
            return _error(tool, "read_failed", str(exc), details={"source": str(source_path)}, action="Check the exact workbook, sheet, range, and Excel authorization.")
        return ToolResult.success(
            tool,
            affected_ranges=(_target_label(normalized.target),),
            after_evidence=evidence,
            metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
        )

    advanced_result = execute_advanced_tool(normalized, engine=engine)
    if advanced_result is not None:
        return advanced_result

    target_error = _require_target(normalized)
    if target_error is not None:
        return target_error

    if tool == "clear_range":
        target_guard = _require_working_copy_target(normalized)
        if target_guard is not None:
            return target_guard
        if engine is None:
            return _error(tool, "engine_required", "No workbook engine was attached to this request.", action="Select and open a compatible engine before range operations.")
        expected_cell_count = normalized.arguments.get("expected_cell_count")
        if not isinstance(expected_cell_count, int) or isinstance(expected_cell_count, bool) or expected_cell_count < 1:
            return _error(tool, "expected_cell_count_required", "arguments.expected_cell_count must name the exact positive number of cells being cleared.")
        try:
            range_rows, range_columns, actual_cell_count = _bounded_shape(engine, normalized.target)
            if actual_cell_count > _MAX_CELLS_PER_RANGE_OPERATION:
                return _error(
                    tool,
                    "range_too_large",
                    f"{actual_cell_count} cells exceeds the current limit of {_MAX_CELLS_PER_RANGE_OPERATION}.",
                    details={"cell_count": actual_cell_count, "limit": _MAX_CELLS_PER_RANGE_OPERATION},
                )
            before_values = engine.read_values(normalized.target)
        except Exception as exc:
            return _error(tool, "before_fingerprint_failed", str(exc), details={"target": normalized.target.to_dict()})
        observed_shape = _shape(before_values)
        if observed_shape != (range_rows, range_columns):
            # A COM union range can return only its first area; the resolved
            # bounded shape and returned matrix must therefore agree exactly.
            return _error(tool, "range_shape_mismatch", "Resolved range evidence is incomplete; refusing to clear it.")
        if actual_cell_count != expected_cell_count:
            return _error(
                tool,
                "cell_count_mismatch",
                f"Resolved range has {actual_cell_count} cell(s) but expected_cell_count={expected_cell_count}.",
                details={"actual_cell_count": actual_cell_count, "expected_cell_count": expected_cell_count},
                action="Recompute the exact expected cell count for this range before retrying.",
            )
        if normalized.dry_run:
            return ToolResult.success(
                tool,
                affected_ranges=(_target_label(normalized.target),),
                warnings=("Dry run: no cells were cleared.",),
                before_evidence={"values": before_values},
                after_evidence={"would_clear_cell_count": actual_cell_count},
                metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
            )
        try:
            engine.clear_range(normalized.target)
        except Exception as exc:
            return _error(tool, "engine_operation_failed", str(exc), details={"target": normalized.target.to_dict()})
        return ToolResult.success(
            tool,
            changed=True,
            affected_ranges=(_target_label(normalized.target),),
            before_evidence={"values": before_values},
            after_evidence={"cleared_cell_count": actual_cell_count},
            metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000), cells_touched=actual_cell_count),
        )

    if tool == "fill_formula_down":
        if normalized.target.workbook_id not in {"working-copy", ""}:  # type: ignore[union-attr]
            return _error(
                tool,
                "invalid_target_workbook",
                "fill_formula_down may only target the Excel-created working copy, not the source.",
            )
        if engine is None:
            return _error(tool, "engine_required", "No workbook engine was attached to this request.", action="Select and open a compatible engine before range operations.")
        template_data = normalized.arguments.get("template")
        if not isinstance(template_data, Mapping):
            return _error(tool, "template_required", "fill_formula_down needs an explicit template target naming the single-row formula source.")
        template_target = TargetRef.from_dict(template_data)
        if template_target.sheet != normalized.target.sheet:
            return _error(tool, "template_target_sheet_mismatch", "The template and target ranges must be on the same sheet.")
        expected_formulas = normalized.arguments.get("expected_template_formulas")
        if not isinstance(expected_formulas, list) or not expected_formulas or not all(isinstance(item, str) for item in expected_formulas):
            return _error(tool, "expected_template_formulas_required", "arguments.expected_template_formulas must be a non-empty list of exact formula strings.")
        expected_row_count = normalized.arguments.get("expected_target_row_count")
        if not isinstance(expected_row_count, int) or isinstance(expected_row_count, bool) or expected_row_count < 1:
            return _error(tool, "expected_target_row_count_required", "arguments.expected_target_row_count must be a positive integer.")

        try:
            resolver = getattr(engine, "validate_bounded_range", engine.resolve_bounds)
            t_first_row, t_first_col, t_last_row, t_last_col = resolver(template_target)
            d_first_row, d_first_col, d_last_row, d_last_col = resolver(normalized.target)
        except Exception as exc:
            return _error(tool, "range_resolution_failed", str(exc))

        if t_first_row != t_last_row:
            return _error(tool, "template_must_be_one_row", "The formula template must be exactly one row.")
        if (t_first_col, t_last_col) != (d_first_col, d_last_col):
            return _error(tool, "column_mismatch", "The template and target ranges must span the exact same columns.")
        if d_first_row != t_first_row + 1:
            return _error(
                tool,
                "not_contiguous",
                f"The target range must start immediately below the template row (expected row {t_first_row + 1}, got {d_first_row}).",
            )
        actual_row_count = d_last_row - d_first_row + 1
        target_cell_count = actual_row_count * (d_last_col - d_first_col + 1)
        if target_cell_count > _MAX_CELLS_PER_RANGE_OPERATION:
            return _error(
                tool,
                "range_too_large",
                f"{target_cell_count} cells exceeds the current limit of {_MAX_CELLS_PER_RANGE_OPERATION}.",
            )
        if actual_row_count != expected_row_count:
            return _error(
                tool,
                "row_count_mismatch",
                f"Target range has {actual_row_count} row(s) but expected_target_row_count={expected_row_count}.",
                details={"actual_row_count": actual_row_count, "expected_row_count": expected_row_count},
            )

        try:
            actual_template_formulas = list(engine.read_formulas(template_target)[0])
        except Exception as exc:
            return _error(tool, "template_read_failed", str(exc))
        if actual_template_formulas != expected_formulas:
            return _error(
                tool,
                "template_fingerprint_mismatch",
                "The template row's actual formulas do not match expected_template_formulas.",
                details={"expected": expected_formulas, "actual": actual_template_formulas},
                action="Read the template row first and pass back its exact current formulas.",
            )

        try:
            before_values = engine.read_values(normalized.target)
            before_formulas = engine.read_formulas(normalized.target)
        except Exception as exc:
            return _error(tool, "before_fingerprint_failed", str(exc))

        if normalized.dry_run:
            return ToolResult.success(
                tool,
                affected_ranges=(_target_label(normalized.target),),
                warnings=("Dry run: no formulas were filled.",),
                before_evidence={"values": before_values, "formulas": before_formulas},
                after_evidence={"would_fill_row_count": actual_row_count},
            )

        try:
            engine.fill_formula_down(template_target, normalized.target)
        except Exception as exc:
            return _error(tool, "engine_operation_failed", str(exc))

        # Defense in depth: FillDown must never overwrite its own source row.
        try:
            template_after = list(engine.read_formulas(template_target)[0])
        except Exception:
            template_after = None
        if template_after is not None and template_after != expected_formulas:
            return _error(
                tool,
                "template_row_was_modified",
                "The fill operation unexpectedly changed the template row; the working copy may be inconsistent.",
                recoverable=False,
            )

        return ToolResult.success(
            tool,
            changed=True,
            affected_ranges=(_target_label(normalized.target),),
            before_evidence={"values": before_values, "formulas": before_formulas},
            after_evidence={"filled_row_count": actual_row_count},
            metrics=ToolMetrics(
                elapsed_ms=round((time.perf_counter() - started) * 1000),
                cells_touched=actual_row_count * (d_last_col - d_first_col + 1),
            ),
        )

    if tool == "insert_columns":
        if normalized.target.workbook_id not in {"working-copy", ""}:  # type: ignore[union-attr]
            return _error(
                tool,
                "invalid_target_workbook",
                "insert_columns may only target the Excel-created working copy, not the source.",
            )
        if engine is None:
            return _error(tool, "engine_required", "No workbook engine was attached to this request.", action="Select and open a compatible engine before range operations.")
        count = normalized.arguments.get("count")
        if not isinstance(count, int) or isinstance(count, bool) or count < 1:
            return _error(tool, "count_required", "arguments.count must be a positive integer.")
        if count > _MAX_COLUMNS_PER_INSERT:
            return _error(
                tool,
                "column_count_too_large",
                f"Column insertion count {count} exceeds the safety limit of {_MAX_COLUMNS_PER_INSERT}.",
            )
        expected_anchor = normalized.arguments.get("expected_anchor_column")
        if not isinstance(expected_anchor, str) or not expected_anchor.strip():
            return _error(tool, "expected_anchor_column_required", "arguments.expected_anchor_column must name the exact column letter (e.g. 'D').")
        try:
            _, actual_anchor_col, _, _ = engine.resolve_bounds(normalized.target)
        except Exception as exc:
            return _error(tool, "range_resolution_failed", str(exc))
        actual_anchor_letter = _column_letter(actual_anchor_col)
        if actual_anchor_letter != expected_anchor.strip().upper():
            return _error(
                tool,
                "anchor_column_mismatch",
                f"Resolved anchor column is {actual_anchor_letter} but expected_anchor_column={expected_anchor!r}.",
                details={"actual": actual_anchor_letter, "expected": expected_anchor},
                action="Recompute the exact expected anchor column before retrying.",
            )
        try:
            engine.calculate_sheet(str(normalized.target.sheet))
            errors_before = engine.count_formula_errors(str(normalized.target.sheet))
        except Exception as exc:
            return _error(tool, "before_fingerprint_failed", str(exc))
        if normalized.dry_run:
            return ToolResult.success(
                tool,
                affected_ranges=(_target_label(normalized.target),),
                warnings=("Dry run: no columns were inserted.",),
                before_evidence={"formula_error_count": errors_before},
                after_evidence={"would_insert_count": count, "anchor_column": actual_anchor_letter},
            )
        try:
            engine.insert_columns(normalized.target, count)
        except Exception as exc:
            return _error(tool, "engine_operation_failed", str(exc))
        try:
            engine.calculate_sheet(str(normalized.target.sheet))
            errors_after = engine.count_formula_errors(str(normalized.target.sheet))
        except Exception as exc:
            return _error(tool, "after_fingerprint_failed", str(exc))
        if errors_after > errors_before:
            return _error(
                tool,
                "reference_errors_introduced",
                f"Column insertion introduced {errors_after - errors_before} new formula error cell(s) on {normalized.target.sheet}.",
                details={"errors_before": errors_before, "errors_after": errors_after},
                recoverable=False,
            )
        return ToolResult.success(
            tool,
            changed=True,
            affected_ranges=(_target_label(normalized.target),),
            before_evidence={"formula_error_count": errors_before},
            after_evidence={"inserted_count": count, "anchor_column": actual_anchor_letter, "formula_error_count": errors_after},
            metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
        )

    if tool == "update_pivot_source":
        if normalized.target.workbook_id not in {"working-copy", ""}:  # type: ignore[union-attr]
            return _error(
                tool,
                "invalid_target_workbook",
                "update_pivot_source may only target the Excel-created working copy, not the source.",
            )
        if not normalized.target.sheet or not normalized.target.object_name:  # type: ignore[union-attr]
            return _error(tool, "pivot_target_required", "update_pivot_source needs target.sheet and target.object_name naming the exact PivotTable.")
        if engine is None:
            return _error(tool, "engine_required", "No workbook engine was attached to this request.", action="Select and open a compatible engine before range operations.")
        expected_current_source = normalized.arguments.get("expected_current_source")
        if not isinstance(expected_current_source, str) or not expected_current_source.strip():
            return _error(tool, "expected_current_source_required", "arguments.expected_current_source must name the exact current source you have already inspected.")
        new_source = normalized.arguments.get("new_source")
        if not isinstance(new_source, str) or not new_source.strip():
            return _error(tool, "new_source_required", "arguments.new_source must name the exact new worksheet/table source.")
        allow_shared = bool(normalized.arguments.get("allow_shared_cache_replacement", False))

        sheet = normalized.target.sheet
        pivot_name = normalized.target.object_name
        try:
            info = engine.inspect_pivot_table(sheet, pivot_name)
        except Exception as exc:
            return _error(tool, "pivot_table_not_found", str(exc))

        if info.get("source_type") != "database":
            return _error(
                tool,
                "unsupported_pivot_source",
                f"PivotTable source type {info.get('source_type')!r} is not a worksheet/table source; "
                "external and Data Model sources are locked as unsupported.",
                recoverable=False,
            )
        try:
            expected_bounds = engine.resolve_source_bounds(expected_current_source)
        except Exception as exc:
            return _error(tool, "expected_current_source_unresolvable", str(exc))
        actual_bounds = info.get("source_bounds")
        if expected_bounds is None:
            return _error(
                tool,
                "expected_current_source_unresolvable",
                "The expected current PivotTable source could not be resolved safely.",
            )
        if actual_bounds is None:
            return _error(
                tool,
                "actual_current_source_unresolvable",
                f"The actual PivotTable source {info.get('source_data')!r} could not be resolved safely.",
                recoverable=False,
            )
        if actual_bounds != expected_bounds:
            return _error(
                tool,
                "current_source_mismatch",
                f"Actual current source {info.get('source_data')!r} does not match "
                f"expected_current_source={expected_current_source!r}.",
                details={"actual": info.get("source_data"), "expected": expected_current_source},
            )
        shared_with = info.get("shared_with") or []
        if shared_with and not allow_shared:
            return _error(
                tool,
                "shared_cache_not_acknowledged",
                f"This PivotTable's cache is shared with {shared_with}; changing it would also change "
                "those PivotTables. Set arguments.allow_shared_cache_replacement=true to proceed.",
                details={"shared_with": shared_with},
            )

        if normalized.dry_run:
            return ToolResult.success(
                tool,
                affected_ranges=(_target_label(normalized.target),),
                warnings=("Dry run: no PivotTable source was changed.",),
                before_evidence=info,
                after_evidence={"would_set_source": new_source},
            )

        try:
            new_bounds = engine.resolve_source_bounds(new_source)
        except Exception as exc:
            return _error(tool, "new_source_unresolvable", str(exc))
        if new_bounds is None:
            return _error(
                tool,
                "new_source_unresolvable",
                "The requested new PivotTable source could not be resolved safely.",
            )

        try:
            engine.update_pivot_source(sheet, pivot_name, new_source)
        except Exception as exc:
            return _error(tool, "engine_operation_failed", str(exc))

        try:
            after_info = engine.inspect_pivot_table(sheet, pivot_name)
        except Exception as exc:
            return _error(tool, "after_fingerprint_failed", str(exc))
        after_bounds = after_info.get("source_bounds")
        if after_bounds is None or after_bounds != new_bounds:
            return _error(
                tool,
                "source_update_did_not_apply",
                "The PivotTable source does not reflect the requested change after refresh.",
                details={"actual": after_info.get("source_data"), "expected": new_source},
                recoverable=False,
            )

        return ToolResult.success(
            tool,
            changed=True,
            affected_ranges=(_target_label(normalized.target),),
            before_evidence=info,
            after_evidence=after_info,
            metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
        )

    if tool in {"read_range", "write_range", "set_formula", "copy_range"}:
        if engine is None:
            return _error(tool, "engine_required", "No workbook engine was attached to this request.", action="Select and open a compatible engine before range operations.")
        if tool != "read_range":
            target_guard = _require_working_copy_target(normalized)
            if target_guard is not None:
                return target_guard
        try:
            if tool == "read_range":
                values = engine.read_values(normalized.target)
                formulas = engine.read_formulas(normalized.target)
                return ToolResult.success(
                    tool,
                    affected_ranges=(_target_label(normalized.target),),
                    after_evidence={"values": values, "formulas": formulas},
                    metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)),
                )
            if tool == "copy_range":
                source = normalized.arguments.get("source")
                if not isinstance(source, Mapping):
                    return _error(tool, "source_required", "copy_range needs an explicit source target.")
                source_target = TargetRef.from_dict(source)
                mode = str(normalized.arguments.get("mode", "all"))
                if source_target.workbook_id not in {"working-copy", ""}:
                    return _error(
                        tool,
                        "cross_workbook_copy_unsupported",
                        "copy_range only supports a source range on the approved working copy; "
                        "source and foreign workbook identifiers are refused.",
                        action="Read the other workbook's range first, then write_range the values into this workbook.",
                    )
                source_rows, source_columns, source_cell_count = _bounded_shape(engine, source_target)
                dest_rows, dest_columns, _ = _bounded_shape(engine, normalized.target)
                if (source_rows, source_columns) != (dest_rows, dest_columns):
                    return _error(
                        tool,
                        "shape_mismatch",
                        f"Source range shape {(source_rows, source_columns)} does not match "
                        f"destination range shape {(dest_rows, dest_columns)}.",
                    )
                if source_cell_count > _MAX_CELLS_PER_RANGE_OPERATION:
                    return _error(
                        tool,
                        "range_too_large",
                        f"{source_cell_count} cells exceeds the current limit of {_MAX_CELLS_PER_RANGE_OPERATION}.",
                        details={"cell_count": source_cell_count, "limit": _MAX_CELLS_PER_RANGE_OPERATION},
                    )
                try:
                    source_values = engine.read_values(source_target)
                except Exception as exc:
                    return _error(tool, "source_read_failed", str(exc), details={"source": source_target.to_dict()})
                source_shape = _shape(source_values)
                try:
                    dest_before_values = engine.read_values(normalized.target)
                except Exception as exc:
                    return _error(tool, "before_fingerprint_failed", str(exc), details={"target": normalized.target.to_dict()})
                dest_shape = _shape(dest_before_values)
                expected_shape = (source_rows, source_columns)
                if source_shape != expected_shape or dest_shape != expected_shape:
                    return _error(
                        tool,
                        "shape_mismatch",
                        f"Source range shape {source_shape} does not match destination range shape {dest_shape}.",
                        details={"source_shape": source_shape, "destination_shape": dest_shape},
                    )
                cell_count = source_cell_count
                if normalized.dry_run:
                    return ToolResult.success(
                        tool,
                        affected_ranges=(_target_label(normalized.target),),
                        warnings=("Dry run: no range was copied.",),
                        before_evidence={"values": dest_before_values},
                        after_evidence={"would_copy_cell_count": cell_count, "mode": mode},
                    )
                engine.copy_range(source_target, normalized.target, mode=mode)
                return ToolResult.success(
                    tool,
                    changed=True,
                    affected_ranges=(_target_label(normalized.target),),
                    before_evidence={"values": dest_before_values},
                    after_evidence={"copied_cell_count": cell_count, "mode": mode},
                    metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000), cells_touched=cell_count),
                )

            argument_name = "values" if tool == "write_range" else "formulas"
            payload = normalized.arguments.get(argument_name)
            shape = _shape(payload)
            if shape is None or shape[0] < 1 or shape[1] < 1:
                return _error(tool, "invalid_matrix", f"arguments.{argument_name} must be a rectangular two-dimensional array.")
            target_rows, target_columns, cell_count = _bounded_shape(engine, normalized.target)
            if shape != (target_rows, target_columns):
                return _error(
                    tool,
                    "shape_mismatch",
                    f"Payload shape {shape} does not match target range shape "
                    f"{(target_rows, target_columns)}.",
                    details={"payload_shape": shape, "target_shape": (target_rows, target_columns)},
                )
            if cell_count > _MAX_CELLS_PER_RANGE_OPERATION:
                return _error(
                    tool,
                    "range_too_large",
                    f"{cell_count} cells exceeds the current chunking-free limit of {_MAX_CELLS_PER_RANGE_OPERATION}.",
                    details={"cell_count": cell_count, "limit": _MAX_CELLS_PER_RANGE_OPERATION},
                    action="Split the range into smaller bounded operations.",
                )
            try:
                before_values = engine.read_values(normalized.target) if tool == "write_range" else engine.read_formulas(normalized.target)
            except Exception as exc:
                return _error(tool, "before_fingerprint_failed", str(exc), details={"target": normalized.target.to_dict()})
            if normalized.dry_run:
                return ToolResult.success(
                    tool,
                    affected_ranges=(_target_label(normalized.target),),
                    warnings=("Dry run: workbook was not changed.",),
                    before_evidence={argument_name: before_values},
                    after_evidence={"shape": {"rows": shape[0], "columns": shape[1]}},
                )
            if tool == "write_range":
                engine.write_values(normalized.target, payload)
            else:
                engine.write_formulas(normalized.target, payload)
            return ToolResult.success(
                tool,
                changed=True,
                affected_ranges=(_target_label(normalized.target),),
                before_evidence={argument_name: before_values},
                after_evidence={"shape": {"rows": shape[0], "columns": shape[1]}},
                metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000), cells_touched=cell_count),
            )
        except Exception as exc:
            return _error(tool, "engine_operation_failed", str(exc), details={"target": normalized.target.to_dict() if normalized.target else {}})

    return _error(tool, "handler_missing", f"Tool is marked available but has no executor handler: {tool}", recoverable=False, action="Fix the catalogue/handler mismatch before continuing.")
