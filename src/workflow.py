"""Dry-run and execute orchestration. All safety sequencing lives here.

Implements the transaction required by the execution plan:

    SOURCE.xlsb -> SaveCopyAs WORKING.xlsb -> COM edits -> validate ->
    save/close -> reopen -> revalidate -> byte-publish FINAL.xlsb

Fail-closed rule: nothing is published unless every discovery, update,
calculation, and validation step succeeds; the source is never saved.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Sequence

from .block_locator import (
    candidate_has_business_lineage,
    detect_existing_actual,
    find_month_block_candidates,
    select_unique_month_block,
)
from .business_sheet_updater import get_last_used_row, update_business_sheet
from .calculation import calculate_workbook_once
from .constants import XL_EXCEL12, target_version_for_month
from .errors import (
    AmbiguousMonthBlockError,
    MissingMonthBlockError,
    MissingSheetError,
    NoRunningExcelError,
    PLAutomationError,
    PublicationError,
    ReopenValidationError,
    UnsupportedScopeError,
    UnsavedWorkbookError,
    ValidationError,
    WorkbookFormatError,
    WorkbookNotFoundError,
)
from .excel_session import ExcelSession
from .file_transaction import (
    assert_source_candidate,
    assert_source_unchanged,
    create_run_paths,
    publish_validated_workbook,
    retain_failed_workbook,
    save_working_copy,
)
from .header_discovery import effective_merged_value, normalize_header_value, read_header_snapshot
from .models import (
    AppConfig,
    HeaderSnapshot,
    MonthBlock,
    RunManifest,
    SheetUpdateResult,
    ValidationCheck,
    ValidationReport,
)
from .reporting import (
    configure_logging,
    render_dry_run_report,
    render_run_report,
    sanitize_for_logging,
    write_manifest_atomic,
    write_run_report,
)
from .total_pl_updater import update_total_pl
from .validation import (
    reconcile_total_pl_at_column,
    validate_business_formulas,
    validate_percent_formulas,
    validate_sheet_structure,
)
from .workbook_audit import collect_fingerprint, compare_preservation


@dataclass
class PreflightResult:
    """Everything dry-run and execute need from read-only discovery."""

    source_path: Path
    mode_requested: str
    mode_used: str
    blocks: dict[str, MonthBlock]
    problems: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ready(self) -> bool:
        return not self.problems


_EXIT_OK = 0
_EXIT_SAFE_EXECUTION_FAILURE = 4
_EXIT_VALIDATION_FAILURE = 5
_EXIT_PUBLICATION_FAILURE = 6


def _log_to_stderr(message: str) -> None:
    import sys

    print(f"[cleanup] {message}", file=sys.stderr)


def _assert_source_saved(workbook: object, source_path: Path) -> None:
    """Refuse to copy unsaved edits from an attached user workbook."""
    try:
        saved = bool(workbook.Saved)
    except Exception as exc:  # pragma: no cover - requires Excel
        raise UnsavedWorkbookError(
            f"Could not prove that the source workbook is saved: {source_path}"
        ) from exc
    if not saved:
        raise UnsavedWorkbookError(
            "The source workbook has unsaved Excel changes. Save it, close/reopen "
            "or discard those changes, then retry; unsaved state will not be copied."
        )


def _safe_close(session: ExcelSession | None) -> None:
    """Close automation-owned resources; never touch user-owned Excel."""
    if session is None:
        return
    try:
        session.restore_state()
    except Exception as exc:  # pragma: no cover - requires Excel
        _log_to_stderr(f"Excel state restore failed during cleanup: {exc}")
    try:
        session.close()
    except Exception as exc:  # pragma: no cover - requires Excel
        _log_to_stderr(f"Excel cleanup failed: {exc}")



# ---------------------------------------------------------------------------
# Source session management (mode auto/attach/open)
# ---------------------------------------------------------------------------


def _connect_source(source_path: Path, mode: str) -> ExcelSession:
    if mode == "attach":
        return ExcelSession.attach(source_path)
    if mode == "open":
        session = ExcelSession.create()
        try:
            session.source_workbook = session.open_workbook(
                source_path, read_only=True, update_links=False
            )
        except Exception:
            _safe_close(session)
            raise
        return session
    if mode == "auto":
        # Attach to an exact-path match first. Fall back to an isolated open
        # only when no matching workbook is open or Excel is not running at
        # all. An ambiguous open-workbook identity must propagate and stop.
        try:
            return ExcelSession.attach(source_path)
        except (WorkbookNotFoundError, NoRunningExcelError):
            pass
        session = ExcelSession.create()
        try:
            session.source_workbook = session.open_workbook(
                source_path, read_only=True, update_links=False
            )
        except Exception:
            _safe_close(session)
            raise
        return session
    raise UnsupportedScopeError(f"Unknown Excel access mode: {mode!r}")


# ---------------------------------------------------------------------------
# Discovery (rediscovered fresh every time; never cached between phases)
# ---------------------------------------------------------------------------


def _sheet_names(workbook: object) -> set[str]:
    return {
        str(workbook.Sheets(index).Name)
        for index in range(1, int(workbook.Sheets.Count) + 1)
    }


def _discover_blocks(
    workbook: object, config: AppConfig
) -> tuple[dict[str, MonthBlock], list[str]]:
    """Locate the unique August block in all four required sheets.

    Discovery problems are collected instead of raised so a report can show
    every sheet's finding. Structural failures (missing sheets) still raise.
    """
    ordered = (*config.target_sheets, config.total_sheet)
    present = _sheet_names(workbook)
    for name in ordered:
        if name not in present:
            raise MissingSheetError(f"Required sheet is missing: {name}")
    blocks: dict[str, MonthBlock] = {}
    problems: list[str] = []
    require_unique = config.validation.require_unique_header_match

    # Discover the three business sheets first.  Total PL uses their resolved
    # T08 columns as a lineage fingerprint, so it must not be selected from a
    # historical T08/S08 block merely because its headers look similar.
    for name in config.target_sheets:
        worksheet = workbook.Worksheets(name)
        snapshot = read_header_snapshot(worksheet)
        last_used_row = get_last_used_row(worksheet)
        candidates = find_month_block_candidates(snapshot, config.codes, last_used_row)
        try:
            block = select_unique_month_block(candidates, require_unique=require_unique)
        except (AmbiguousMonthBlockError, MissingMonthBlockError) as exc:
            evidence = getattr(exc, "evidence", None)
            message = f"{name}: {exc}"
            if evidence:
                message = f"{message} Evidence: {evidence!r}"
            problems.append(message)
            continue
        matches = detect_existing_actual(snapshot, block, config.codes.actual_version)
        if matches:
            problems.append(
                f"{name}: {config.codes.actual_version} already exists at rows/columns "
                f"{matches}; refusing to overwrite"
            )
            continue
        blocks[name] = block

    total_name = config.total_sheet
    if set(blocks) != set(config.target_sheets):
        problems.append(
            f"{total_name}: skipped lineage discovery because one or more business "
            "sheet August blocks were not proven"
        )
        return blocks, problems

    worksheet = workbook.Worksheets(total_name)
    snapshot = read_header_snapshot(worksheet)
    last_used_row = get_last_used_row(worksheet)
    source_columns = {
        name: blocks[name].target_col for name in config.target_sheets
    }
    candidates = find_month_block_candidates(
        snapshot,
        config.codes,
        last_used_row,
        require_period=False,
        source_columns=source_columns,
    )
    lineage_candidates = [
        candidate
        for candidate in candidates
        if candidate_has_business_lineage(
            worksheet,
            candidate,
            {name: blocks[name] for name in config.target_sheets},
        )
    ]
    if not lineage_candidates:
        problems.append(
            f"{total_name}: no August T08/S08 candidate has provable lineage "
            f"to {source_columns!r}"
        )
        return blocks, problems
    try:
        block = select_unique_month_block(
            lineage_candidates, require_unique=require_unique
        )
    except (AmbiguousMonthBlockError, MissingMonthBlockError) as exc:
        evidence = getattr(exc, "evidence", None)
        message = f"{total_name}: {exc}"
        if evidence:
            message = f"{message} Evidence: {evidence!r}"
        problems.append(message)
        return blocks, problems
    matches = detect_existing_actual(snapshot, block, config.codes.actual_version)
    if matches:
        problems.append(
            f"{total_name}: {config.codes.actual_version} already exists at rows/columns "
            f"{matches}; refusing to overwrite"
        )
    else:
        blocks[total_name] = block
    return blocks, problems


def discover_blocks_for_run(workbook: object, config: AppConfig):
    """Fail-closed discovery used inside execute; problems become errors."""
    blocks, problems = _discover_blocks(workbook, config)
    expected = (*config.target_sheets, config.total_sheet)
    if problems or set(blocks) != set(expected):
        raise ValidationError(
            "Rediscovery did not reproduce a clean editable state",
            evidence={"problems": problems, "discovered": sorted(blocks)},
        )
    return blocks


# ---------------------------------------------------------------------------
# Post-update structural checks specific to this workflow stage
# ---------------------------------------------------------------------------


def _merge_covers_column(worksheet: object, row: int, column: int) -> bool | None:
    cell = worksheet.Cells(row, column)
    try:
        if not bool(cell.MergeCells):
            return False
        area = cell.MergeArea
        first = int(area.Column)
        last = int(area.Column) + int(area.Columns.Count) - 1
    except Exception:  # pragma: no cover - requires Excel
        return None
    return first <= column <= last


def _august_merge_checks(
    worksheet: object, block: MonthBlock, actual_pct_col: int
) -> list[ValidationCheck]:
    if block.month_merge is None or block.month_header_row is None:
        return [
            ValidationCheck(
                "august_merge_extended",
                True,
                False,
                "No August month merge was discovered; nothing needed extending",
                {},
            )
        ]
    row = block.month_header_row
    anchor_col = block.month_merge.first_column
    covered = _merge_covers_column(worksheet, row, actual_pct_col)
    merge_address = "?"
    try:
        merge_address = str(worksheet.Cells(row, anchor_col).MergeArea.Address)
    except Exception:  # pragma: no cover - requires Excel
        pass
    return [
        ValidationCheck(
            "august_merge_extended",
            bool(covered),
            True,
            f"August merge covers A08 '%' column ({merge_address})",
            {"anchor_row": row, "anchor_col": anchor_col, "expected_last": actual_pct_col},
        )
    ]


def _september_still_present_check(
    snapshot: HeaderSnapshot, block: MonthBlock, codes, shift_after_insert: int = 2
) -> ValidationCheck:
    """After the two-column insert September sits two columns further right."""
    version_row = block.version_header_row
    shifted = block.september_start_col + shift_after_insert
    value = normalize_header_value(
        effective_merged_value(snapshot, version_row, shifted)
    )
    expected_period = ""
    period_ok: bool | None = None
    month_ok: bool | None = None
    if codes.month < 12:
        year_text, current_text = codes.period.split(".")
        expected_period = f"{year_text}.{int(current_text) + 1:03d}"
    if block.period_header_row is not None and expected_period:
        found_period = normalize_header_value(
            effective_merged_value(snapshot, block.period_header_row, shifted)
        )
        period_ok = found_period == normalize_header_value(expected_period)
    if block.month_header_row is not None:
        merged_value = normalize_header_value(
            effective_merged_value(snapshot, block.month_header_row, shifted)
        )
        month_ok = "SEPTEMBER" in merged_value or merged_value.startswith("SEP")
    expected_next_target = (
        target_version_for_month(codes.month + 1) if codes.month < 9 else ""
    )
    version_ok = (
        value == normalize_header_value(expected_next_target)
        or "A09" in value
        or value.startswith("SEP")
        or (
        bool(expected_period) and value == ""
        )
    )
    signals = [version_ok] + [item for item in (period_ok, month_ok) if item is not None]
    return ValidationCheck(
        "september_block_still_present",
        all(signals),
        True,
        f"September header still follows A08% at column {shifted}",
        {
            "value": value,
            "expected_period": expected_period,
            "period_match": period_ok,
            "month_match": month_ok,
            "column": shifted,
        },
    )


def _post_update_sheet_checks(
    worksheet: object,
    block: MonthBlock,
    result: SheetUpdateResult,
    codes,
) -> list[ValidationCheck]:
    snapshot = read_header_snapshot(worksheet)
    version_row = block.version_header_row
    amount_col = result.actual_amount_col
    pair_columns = range(block.start_col, block.september_start_col + 2)
    wanted = normalize_header_value(codes.actual_version)
    a08_cells = [
        column
        for column in pair_columns
        if normalize_header_value(
            effective_merged_value(snapshot, version_row, column)
        )
        == wanted
    ]
    t08_value = normalize_header_value(
        effective_merged_value(snapshot, version_row, block.target_col)
    )
    s08_value = normalize_header_value(
        effective_merged_value(snapshot, version_row, block.forecast_col)
    )
    checks = [
        ValidationCheck(
            "unique_a08_pair",
            a08_cells == [amount_col],
            True,
            f"Exactly one {codes.actual_version} header inside the extended block",
            {"a08_columns": a08_cells, "expected": [amount_col]},
        ),
        ValidationCheck(
            "prior_versions_preserved",
            t08_value == normalize_header_value(codes.target_version)
            and s08_value == normalize_header_value(codes.forecast_version),
            True,
            "T08 and S08 remain unchanged in the August block",
            {"t08": t08_value, "s08": s08_value},
        ),
        _september_still_present_check(snapshot, block, codes),
    ]
    checks.extend(_august_merge_checks(worksheet, block, result.actual_pct_col))
    return checks


def _validate_updated_sheet(
    workbook: object,
    name: str,
    result: SheetUpdateResult,
    config: AppConfig,
    report: ValidationReport,
) -> None:
    worksheet = workbook.Worksheets(name)
    report.checks.extend(
        validate_sheet_structure(worksheet, config.codes, result.actual_amount_col)
    )
    report.checks.extend(validate_business_formulas(worksheet, result, config.codes))
    report.checks.extend(validate_percent_formulas(worksheet, result))
    report.checks.extend(
        _post_update_sheet_checks(worksheet, result.before_block, result, config.codes)
    )


# ---------------------------------------------------------------------------
# Preflight / dry run
# ---------------------------------------------------------------------------


def preflight(source_path: Path, config: AppConfig, mode: str) -> PreflightResult:
    """Read-only proof pass: connect, fingerprint, verify format, discover."""
    assert_source_candidate(source_path)
    session = _connect_source(source_path, mode)
    try:
        workbook = session.source_workbook
        _assert_source_saved(workbook, source_path)
        fingerprint = collect_fingerprint(workbook, source_path)
        if fingerprint.file_format != XL_EXCEL12:
            raise WorkbookFormatError(
                "Source workbook is not Excel binary format "
                f"(FileFormat={fingerprint.file_format})"
            )
        blocks, problems = _discover_blocks(workbook, config)
        warnings = [
            f"{len(fingerprint.external_links)} external link source(s) recorded",
            f"Pivot tables per sheet: {fingerprint.pivot_counts}",
            f"Connections: {fingerprint.connection_count}; defined names: "
            f"{fingerprint.defined_name_count}; VBA project present: "
            f"{fingerprint.has_vba_project}",
        ]
        return PreflightResult(
            source_path=source_path,
            mode_requested=mode,
            mode_used=session.mode,
            blocks=blocks,
            problems=problems,
            warnings=warnings,
        )
    finally:
        _safe_close(session)


def run_dry_run(source_path: Path, config: AppConfig, mode: str) -> int:
    """Read-only proof run. Never saves, edits, or publishes anything."""
    assert_source_candidate(source_path)
    try:
        pre = preflight(source_path, config, mode)
    except PLAutomationError as exc:
        header_lines = [
            f"SOURCE: {source_path}",
            "MODE: DRY RUN",
            f"EXCEL ACCESS MODE: {mode}",
            f"YEAR: {config.codes.year}  MONTH: {config.codes.month_name}",
            f"VERSIONS: {config.codes.target_version}/"
            f"{config.codes.forecast_version}/{config.codes.actual_version}  "
            f"PERIOD: {config.codes.period}",
            "",
            "PREFLIGHT FAILED BEFORE DISCOVERY COULD COMPLETE:",
            f"  [{exc.code}] {exc.message}",
            "",
            "READY TO EXECUTE: NO",
        ]
        print("\n".join(header_lines))
        return 3
    text = render_dry_run_report(
        source_path,
        config,
        pre.mode_used,
        pre.blocks,
        ready=pre.ready,
        warnings=[],
        problems=pre.problems,
    )
    print(text)
    print("DRY-RUN NOTE: no files were created, edited, or saved.")
    return _EXIT_OK if pre.ready else 3


# ---------------------------------------------------------------------------
# Execute
# ---------------------------------------------------------------------------


def _validation_summary(reports: Sequence[ValidationReport]) -> dict[str, Any]:
    summary: dict[str, Any] = {}
    for report in reports:
        failed = [check.name for check in report.checks if not check.passed]
        mismatches = [
            item.total_pl_row for item in report.reconciliations if not item.passed
        ]
        summary[report.stage] = {
            "passed": report.passed,
            "checks_total": len(report.checks),
            "failed_checks": failed,
            "reconciliation_rows": len(report.reconciliations),
            "mismatch_rows": mismatches,
        }
    return summary


def _flush_manifest(
    paths,
    manifest: RunManifest,
    reports: list[ValidationReport],
    updates: dict[str, Any],
    *,
    status: str,
    phase: str,
    error_payload: dict[str, Any] | None = None,
) -> None:
    manifest.status = status
    manifest.phase = phase
    manifest.updates = updates
    manifest.validations = _validation_summary(reports)
    if error_payload is not None:
        manifest.error = error_payload
    if paths is not None:
        manifest.ended_utc = datetime.now(timezone.utc).isoformat()
        write_manifest_atomic(paths.manifest_path, manifest)
        write_run_report(paths.report_path, render_run_report(manifest, reports))


def _fail_closed(
    exc: PLAutomationError,
    paths,
    manifest: RunManifest,
    reports: list[ValidationReport],
    updates: dict[str, Any],
    logger=None,
) -> None:
    if logger is not None:
        logger.error(
            "Failure [%s] at %s: %s",
            exc.code,
            exc.phase or manifest.phase,
            exc.message,
        )
    if paths is not None and exc.code != "publication_error":
        retained = retain_failed_workbook(paths)
        if retained is not None:
            updates["failed_workbook_retained"] = str(retained)
    error_payload = {
        "code": exc.code,
        "message": exc.message,
        "phase": exc.phase or manifest.phase,
        "evidence": sanitize_for_logging(getattr(exc, "evidence", None)),
    }
    _flush_manifest(
        paths,
        manifest,
        reports,
        updates,
        status="FAILED",
        phase=error_payload["phase"],
        error_payload=error_payload,
    )


def run_execute(source_path: Path, config: AppConfig, mode: str, project_root: Path) -> int:
    """Full guarded transaction. Returns a process exit code."""
    manifest = RunManifest(
        run_id="pending",
        status="RUNNING",
        phase="starting",
        source=str(source_path),
        output=None,
        started_utc=datetime.now(timezone.utc).isoformat(),
    )
    manifest.codes = {
        key: getattr(config.codes, key)
        for key in (
            "year", "month", "month_name", "period",
            "target_version", "forecast_version", "actual_version",
        )
    }
    pre_session: ExcelSession | None = None
    work_session: ExcelSession | None = None
    reopen_session: ExcelSession | None = None
    paths = None
    reports: list[ValidationReport] = []
    updates: dict[str, Any] = {}
    logger = None

    try:
        # ---- steps 1-2: run the complete preflight again; never trust older runs
        pre = preflight(source_path, config, mode)
        if not pre.ready:
            reasons = "; ".join(pre.problems)
            print(f"EXECUTE REFUSED (preflight not ready): {reasons}")
            return 3

        # ---- step 3: run paths, logging, manifest initialization
        paths = create_run_paths(project_root, source_path, config.codes.actual_version)
        logger = configure_logging(paths.run_dir)
        logger.info("Run %s starting against %s (mode=%s)", paths.run_id, source_path, mode)
        manifest.run_id = paths.run_id
        manifest.fingerprints["discovery"] = {
            name: list(block.evidence) for name, block in pre.blocks.items()
        }
        _flush_manifest(paths, manifest, reports, updates, status="RUNNING", phase="preflight_done")

        # ---- steps 2+4: fingerprint/hash the source, then SaveCopyAs via COM
        pre_session = _connect_source(source_path, mode)
        try:
            source_wb = pre_session.source_workbook
            _assert_source_saved(source_wb, source_path)
            before_fp = collect_fingerprint(source_wb, source_path)
            manifest.fingerprints["source_before"] = {
                "sha256": before_fp.sha256,
                "size_bytes": before_fp.size_bytes,
                "modified_utc": before_fp.modified_utc,
                "file_format": before_fp.file_format,
                "external_links": before_fp.external_links,
                "sheet_names": before_fp.sheet_names,
            }
            _flush_manifest(paths, manifest, reports, updates,
                            status="RUNNING", phase="working_copy_creating")
            save_working_copy(source_wb, paths.working_path)
        finally:
            # ---- step 5: detach from the source without ever saving it
            _safe_close(pre_session)
            pre_session = None

        # ---- steps 6-7: isolated instance on the working copy; re-prove identity
        work_session = ExcelSession.create()
        workbook = work_session.open_workbook(
            paths.working_path, read_only=False, update_links=False
        )
        if workbook.FileFormat != XL_EXCEL12:
            raise WorkbookFormatError("Working copy lost Excel binary format")
        blocks = discover_blocks_for_run(workbook, config)
        # Baseline of the untouched Excel-created working copy.  The
        # preservation validator compares this with the edited workbook.
        working_fp = collect_fingerprint(workbook, paths.working_path)
        updates["working_copy"] = str(paths.working_path.name)
        _flush_manifest(paths, manifest, reports, updates,
                        status="RUNNING", phase="working_copy_confirmed")

        # ---- step 8: capture state, then apply the safe editing posture
        work_session.capture_state()
        work_session.apply_editing_state()

        # ---- steps 9-11: three business-total sheets, each locally validated
        results_by_sheet: dict[str, SheetUpdateResult] = {}
        for name in config.target_sheets:
            result = update_business_sheet(
                workbook.Worksheets(name), blocks[name], config.codes
            )
            results_by_sheet[name] = result
            updates[name] = {
                "actual_amount_col": result.actual_amount_col,
                "actual_pct_col": result.actual_pct_col,
                "merge_repaired": result.merge_repaired,
                "warnings": result.warnings,
            }
            _flush_manifest(paths, manifest, reports, updates,
                            status="RUNNING", phase=f"updated_{name}")
            logger.info("%s updated at columns %s/%s",
                        name, result.actual_amount_col, result.actual_pct_col)

        # ---- step 12: Total PL only through proven row mappings
        total_name = config.total_sheet
        total_result, mappings = update_total_pl(
            workbook.Worksheets(total_name),
            blocks[total_name],
            results_by_sheet,
            config.codes,
        )
        updates[total_name] = {
            "actual_amount_col": total_result.actual_amount_col,
            "actual_pct_col": total_result.actual_pct_col,
            "merge_repaired": total_result.merge_repaired,
            "warnings": total_result.warnings,
        }
        _flush_manifest(paths, manifest, reports, updates,
                        status="RUNNING", phase=f"updated_{total_name}")

        # ---- step 13: one full calculation rebuild, waited to completion
        elapsed = calculate_workbook_once(
            work_session.app,
            workbook,
            config.validation.calculation_timeout_seconds,
        )
        updates["calculation_seconds"] = round(elapsed, 2)
        _flush_manifest(paths, manifest, reports, updates,
                        status="RUNNING", phase="calculated")

        # ---- step 14: complete pre-save validation battery
        pre_save = ValidationReport(stage="pre_save")
        after_fp = collect_fingerprint(workbook, paths.working_path)
        pre_save.checks.extend(compare_preservation(working_fp, after_fp))
        for name in config.target_sheets:
            _validate_updated_sheet(workbook, name, results_by_sheet[name], config, pre_save)
        _validate_updated_sheet(workbook, total_name, total_result, config, pre_save)
        pre_save.reconciliations.extend(
            reconcile_total_pl_at_column(
                workbook.Worksheets(total_name),
                mappings,
                total_result.actual_amount_col,
                config.validation.numeric_tolerance,
            )
        )
        reports.append(pre_save)
        _flush_manifest(paths, manifest, reports, updates,
                        status="RUNNING", phase="pre_save_validated")
        # ---- step 15: any required failure stops publication here
        if not pre_save.passed:
            raise ValidationError(
                "Pre-save validation failed; refusing to save or publish",
                evidence=_validation_summary(reports),
            )

        # ---- steps 16-17: save and close cleanly through Excel COM
        workbook.Save()
        updates["saved_through_com"] = True
        work_session.close_workbook(workbook, save_changes=False)
        _safe_close(work_session)
        work_session = None
        _flush_manifest(paths, manifest, reports, updates,
                        status="RUNNING", phase="saved_and_closed")

        # ---- steps 18-20: reopen with links disabled; validate again
        reopen_session = ExcelSession.create()
        reopened = None
        recalculated = False
        try:
            reopened = reopen_session.open_workbook(
                paths.working_path, read_only=False, update_links=False
            )
            post_save = ValidationReport(stage="post_reopen")
            post_fp = collect_fingerprint(reopened, paths.working_path)
            post_save.checks.extend(compare_preservation(after_fp, post_fp))
            for name in config.target_sheets:
                _validate_updated_sheet(reopened, name, results_by_sheet[name], config, post_save)
            _validate_updated_sheet(reopened, total_name, total_result, config, post_save)
            # Recalculate only if Excel still owes us stable values.
            try:
                pending_state = int(reopen_session.app.CalculationState)
            except Exception:  # pragma: no cover - requires Excel
                pending_state = 0
            if pending_state != 0:
                elapsed_reopen = calculate_workbook_once(
                    reopen_session.app,
                    reopened,
                    config.validation.calculation_timeout_seconds,
                    full_rebuild=False,
                )
                recalculated = True
                updates["reopen_calculation_seconds"] = round(elapsed_reopen, 2)
            post_save.reconciliations.extend(
                reconcile_total_pl_at_column(
                    reopened.Worksheets(total_name),
                    mappings,
                    total_result.actual_amount_col,
                    config.validation.numeric_tolerance,
                )
            )
            reports.append(post_save)
            updates["post_reopen_recalculated"] = recalculated
            _flush_manifest(paths, manifest, reports, updates,
                            status="RUNNING", phase="post_reopen_validated")
            if not post_save.passed:
                raise ReopenValidationError(
                    "Post-reopen validation failed; refusing to publish",
                    evidence=_validation_summary(reports),
                )
        finally:
            if reopened is not None:
                reopen_session.close_workbook(reopened, save_changes=False)
            _safe_close(reopen_session)
            reopen_session = None
        _flush_manifest(paths, manifest, reports, updates,
                        status="RUNNING", phase="closed_cleanly")

        # ---- step 22: reconfirm the source never changed during the run
        assert_source_unchanged(before_fp, source_path)

        # ---- steps 23-24: publish closed file; byte-identity is verified inside
        final_output = publish_validated_workbook(paths.working_path, paths.final_path)
        manifest.output = final_output

        # ---- step 25: final reports with SUCCESS
        updates["final_sha256_matches_working"] = True
        _flush_manifest(paths, manifest, reports, updates,
                        status="SUCCESS", phase="published")
        logger.info("Published %s", final_output)
        print(f"SUCCESS: published {final_output}")
        print(f"Evidence: {paths.manifest_path}")
        print(f"Evidence: {paths.report_path}")
        return _EXIT_OK

    except PublicationError as exc:
        _fail_closed(exc, paths, manifest, reports, updates, logger)
        print("PUBLICATION FAILED - no final workbook was produced.")
        print(f"[{exc.code}] {exc.message}")
        return _EXIT_PUBLICATION_FAILURE
    except (ValidationError, ReopenValidationError) as exc:
        _fail_closed(exc, paths, manifest, reports, updates, logger)
        print("VALIDATION FAILED - nothing was published; diagnostics retained.")
        print(f"[{exc.code}] {exc.message}")
        return _EXIT_VALIDATION_FAILURE
    except PLAutomationError as exc:
        _fail_closed(exc, paths, manifest, reports, updates, logger)
        print("EXECUTION STOPPED SAFELY - see the run report for details.")
        print(f"[{exc.code}] {exc.message}")
        return _EXIT_SAFE_EXECUTION_FAILURE
    except Exception as exc:  # unexpected failures must also fail closed
        wrapped = PLAutomationError(f"Unexpected failure: {exc}", phase=manifest.phase)
        _fail_closed(wrapped, paths, manifest, reports, updates, logger)
        print("UNEXPECTED FAILURE - stopped closed; details are in the run log.")
        return _EXIT_SAFE_EXECUTION_FAILURE
    finally:
        # Every exception path restores state and closes only owned resources.
        _safe_close(work_session)
        _safe_close(reopen_session)
