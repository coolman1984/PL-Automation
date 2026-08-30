"""P&L A08 automation command-line entry point.

Exit codes (execution plan section 8.16):
    0  success / dry-run ready
    2  CLI or configuration error
    3  preflight not ready
    4  safe execution failure
    5  validation failure
    6  publication failure
"""

from __future__ import annotations

import argparse
import json
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

from src.file_probe import probe_excel_file, render_probe_report, resolve_com_mode


def application_root() -> Path:
    """Return the user-facing application folder in source and frozen builds."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="app.py",
        description="Guarded Excel agent foundation and P&L August recipe.",
    )
    parser.add_argument("--file", required=False, help="Path to the source .xlsb workbook")
    parser.add_argument("--year", type=int, default=None, help="Plan year (default from config)")
    parser.add_argument("--month", type=int, default=None, help="Plan month number 1-12")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--self-check", action="store_true", help="Check the local runtime and Excel installation")
    mode.add_argument("--agent-start", action="store_true", help="Print the compact first-read guide for a coding agent")
    mode.add_argument("--project-status", action="store_true", help="Print the current machine-readable project status")
    mode.add_argument("--describe-tool", metavar="TOOL_NAME", help="Describe one declared agent tool")
    mode.add_argument("--run-tool", metavar="REQUEST_JSON", help="Execute one declared JSON tool request file")
    mode.add_argument("--list-tools", action="store_true", help="Print the machine-readable Excel agent tool catalogue")
    mode.add_argument("--backup-only", action="store_true", help="Create and verify a byte-for-byte backup without opening Excel")
    mode.add_argument("--snapshot", action="store_true", help="Create a read-only workbook JSON snapshot through Excel")
    mode.add_argument("--prepare", action="store_true", help="Create a verified backup and then a read-only JSON snapshot")
    mode.add_argument("--probe-only", action="store_true", help="Quick dependency-free container/protection check; does not open Excel")
    mode.add_argument("--dry-run", action="store_true", help="Read-only discovery proof; edits nothing")
    mode.add_argument("--execute", action="store_true", help="Run the full guarded transaction")
    parser.add_argument(
        "--mode",
        choices=("auto", "attach", "open"),
        default="auto",
        help="How to reach the source: attach to running Excel, open isolated, or auto",
    )
    parser.add_argument("--config", default="config.yaml", help="Configuration YAML path")
    parser.add_argument("--artifact-root", default="agent_artifacts", help="Backup and snapshot output folder")
    parser.add_argument("--snapshot-mode", choices=("auto", "inventory", "full"), default="auto")
    parser.add_argument("--max-snapshot-cells", type=int, default=250000)
    parser.add_argument("--format", choices=("text", "json"), default="text", help="Output format for agent discovery commands")
    parser.add_argument("--verbose", action="store_true", help="Verbose console logging")
    return parser


def main(argv: list[str] | None = None) -> int:
    # Windows legacy consoles default to CP1252 and cannot print characters
    # such as the star in the real workbook filename. Reconfigure to UTF-8 so
    # CLI output never aborts a run. File access always uses the true path.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.agent_start or args.project_status or args.describe_tool or args.list_tools:
        from src.agent_entry import render_agent_start, render_project_status, render_tool_description
        from src.tool_registry import tool_catalog

        json_output = args.format == "json"
        if args.agent_start:
            print(render_agent_start(application_root(), json_output=json_output), end="")
        elif args.project_status:
            print(render_project_status(application_root(), json_output=json_output), end="")
        elif args.describe_tool:
            print(render_tool_description(args.describe_tool, json_output=json_output), end="")
        else:
            if json_output:
                print(json.dumps(tool_catalog(), ensure_ascii=False, indent=2))
            else:
                catalog = tool_catalog()
                print("EXCEL AGENT TOOLS")
                print(f"Available: {catalog['available_count']} / Known: {catalog['tool_count']}")
                for item in catalog["tools"]:
                    print(f"[{item['status'].upper()}] {item['name']} — {item['description']}")
        return 0

    if args.run_tool:
        from src.tool_executor import execute_tool

        request_path = Path(args.run_tool).expanduser().resolve()
        try:
            request = json.loads(request_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            print(f"REQUEST ERROR: Could not read JSON tool request: {exc}")
            return 2
        result = execute_tool(request, project_root=application_root())
        print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))
        return 0 if result.ok else 4

    if args.self_check:
        from src.runtime_check import render_runtime_check, run_runtime_check

        report = run_runtime_check(application_root())
        print(render_runtime_check(report))
        return 0 if report["ready"] else 3

    if not args.file:
        parser.error("--file is required for this operation")

    source_path = Path(args.file).expanduser().resolve()
    if not source_path.exists() or not source_path.is_file():
        print(f"INPUT ERROR: Workbook was not found: {source_path}")
        return 2

    try:
        probe = probe_excel_file(source_path)
    except OSError as exc:
        print(f"INPUT ERROR: File could not be inspected: {exc}")
        return 2
    print(render_probe_report(probe))
    if args.probe_only:
        return 0 if probe.recognized else 3

    artifact_root = Path(args.artifact_root).expanduser()
    if not artifact_root.is_absolute():
        artifact_root = application_root() / artifact_root
    artifact_root = artifact_root.resolve()

    if args.backup_only:
        from src.backup_bundle import create_backup_bundle

        try:
            bundle = create_backup_bundle(source_path, artifact_root / "backups")
        except (OSError, ValueError) as exc:
            print(f"BACKUP ERROR: {exc}")
            return 4
        print(json.dumps(bundle.to_dict(), ensure_ascii=False, indent=2))
        return 0

    if args.snapshot or args.prepare:
        if not probe.recognized:
            print("SAFETY ERROR: Workbook format is not recognized safely.")
            return 3
        effective_mode = resolve_com_mode(args.mode, probe)
        if effective_mode == "auto":
            effective_mode = "open"
        if probe.protection in {"nasca_drm", "office_encrypted"} and effective_mode != "attach":
            print("SAFETY ERROR: Open the authorized workbook in Excel and use attach mode.")
            return 3
        if effective_mode not in {"attach", "open"}:
            print("SAFETY ERROR: No safe Excel connection route was selected.")
            return 3
        try:
            from src.backup_bundle import create_backup_bundle
            from src.excel_session import ExcelSession
            from src.workbook_snapshot import SnapshotOptions, build_workbook_snapshot, write_snapshot_json

            bundle = None
            if args.prepare:
                bundle = create_backup_bundle(source_path, artifact_root / "backups")
                output_dir = bundle.directory
            else:
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
                options = SnapshotOptions(mode=args.snapshot_mode, max_cells=args.max_snapshot_cells)
                snapshot = build_workbook_snapshot(workbook, source_path, options)
                snapshot_path = write_snapshot_json(snapshot, output_dir / "workbook_snapshot.json")

            result = {
                "status": "prepared" if args.prepare else "snapshotted",
                "backup": bundle.to_dict() if bundle else None,
                "snapshot_file": str(snapshot_path),
                "snapshot_mode": snapshot["snapshot_mode"],
                "cell_count": snapshot["cell_count"],
                "style_count": len(snapshot["styles"]),
            }
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0
        except ImportError as exc:
            print(f"SETUP ERROR: Excel execution dependencies are unavailable: {exc}")
            return 2
        except (OSError, ValueError) as exc:
            print(f"PREPARATION ERROR: {exc}")
            return 4

    # This P&L updater currently writes XLSB only. The probe still reports
    # fast-path eligibility for other workbook types so a future generic engine
    # can reuse the same deterministic decision layer.
    if source_path.suffix.casefold() != ".xlsb":
        print(
            "INPUT ERROR: This P&L A08 updater currently executes XLSB only. "
            "The quick check completed, but no write engine was selected."
        )
        return 2

    effective_mode = resolve_com_mode(args.mode, probe)
    if probe.protection in {"nasca_drm", "office_encrypted"} and effective_mode != "attach":
        print(
            "SAFETY ERROR: Protected workbooks require --mode attach after the "
            "exact file is opened manually in authorized desktop Excel."
        )
        return 3

    try:
        from src.config import load_config
        from src.errors import ConfigurationError, PLAutomationError, WorkbookFormatError, WorkbookNotFoundError
        from src.file_transaction import assert_source_candidate
        from src.workflow import run_dry_run, run_execute
    except ImportError as exc:
        print(
            "SETUP ERROR: Excel execution dependencies are unavailable. The quick "
            f"check succeeded, but execution cannot start: {exc}"
        )
        return 2

    try:
        assert_source_candidate(source_path)
    except (WorkbookNotFoundError, WorkbookFormatError) as exc:
        print(f"INPUT ERROR: {exc}")
        return 2

    config_path = Path(args.config).expanduser()
    if not config_path.is_absolute():
        config_path = application_root() / config_path
    config_path = config_path.resolve()
    try:
        config = load_config(config_path, year=args.year, month=args.month, execution=bool(args.execute))
    except ConfigurationError as exc:
        print(f"CONFIGURATION ERROR: {exc.message}")
        return 2

    project_root = application_root()

    try:
        if args.dry_run:
            return int(run_dry_run(source_path, config, effective_mode))
        return int(run_execute(source_path, config, effective_mode, project_root))
    except PLAutomationError as exc:
        # Safety net only; workflow normally maps errors itself.
        print(f"[{exc.code}] {exc.message}")
        if isinstance(exc, ConfigurationError):
            return 2
        return 4


if __name__ == "__main__":
    sys.exit(main())
