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

    target_error = _require_target(normalized)
    if target_error is not None:
        return target_error

    if tool in {"read_range", "write_range", "set_formula", "copy_range"}:
        if engine is None:
            return _error(tool, "engine_required", "No workbook engine was attached to this request.", action="Select and open a compatible engine before range operations.")
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
                if normalized.dry_run:
                    return ToolResult.success(tool, warnings=("Dry run: no range was copied.",), affected_ranges=(_target_label(normalized.target),))
                engine.copy_range(source_target, normalized.target, mode=mode)
                return ToolResult.success(tool, changed=True, affected_ranges=(_target_label(normalized.target),), metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000)))

            argument_name = "values" if tool == "write_range" else "formulas"
            payload = normalized.arguments.get(argument_name)
            shape = _shape(payload)
            if shape is None:
                return _error(tool, "invalid_matrix", f"arguments.{argument_name} must be a rectangular two-dimensional array.")
            if normalized.dry_run:
                return ToolResult.success(tool, affected_ranges=(_target_label(normalized.target),), warnings=("Dry run: workbook was not changed.",), after_evidence={"shape": {"rows": shape[0], "columns": shape[1]}})
            if tool == "write_range":
                engine.write_values(normalized.target, payload)
            else:
                engine.write_formulas(normalized.target, payload)
            return ToolResult.success(tool, changed=True, affected_ranges=(_target_label(normalized.target),), after_evidence={"shape": {"rows": shape[0], "columns": shape[1]}}, metrics=ToolMetrics(elapsed_ms=round((time.perf_counter() - started) * 1000), cells_touched=shape[0] * shape[1]))
        except Exception as exc:
            return _error(tool, "engine_operation_failed", str(exc), details={"target": normalized.target.to_dict() if normalized.target else {}})

    return _error(tool, "handler_missing", f"Tool is marked available but has no executor handler: {tool}", recoverable=False, action="Fix the catalogue/handler mismatch before continuing.")
