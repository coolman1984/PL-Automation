"""Workflow-agnostic Excel session adapters for the mutating coordinator half.

These wrap the existing generic primitives (:class:`ExcelSession`,
``file_transaction.save_working_copy``, ``workbook_audit.collect_fingerprint``)
with no assumption about sheet layout, business content, or workbook format.
The P&L-specific workflow in ``src/workflow.py`` is untouched; this module is
the SAP-agnostic equivalent the generic coordinator composes.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from ..engines.excel_com import ExcelComEngine
from ..errors import NoRunningExcelError, WorkbookNotFoundError
from ..excel_session import ExcelSession
from ..file_transaction import save_working_copy
from ..models import ValidationCheck, WorkbookFingerprint
from ..workbook_audit import collect_fingerprint


def connect_source_readonly(source_path: Path, mode: str) -> ExcelSession:
    """Attach to an already-open exact-path workbook, or open it read-only."""
    if mode == "attach":
        return ExcelSession.attach(source_path)
    if mode == "open":
        session = ExcelSession.create()
        try:
            session.source_workbook = session.open_workbook(
                source_path, read_only=True, update_links=False
            )
        except Exception:
            session.close()
            raise
        return session
    if mode == "auto":
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
            session.close()
            raise
        return session
    raise ValueError(f"Unknown Excel access mode: {mode!r}")


def create_working_copy(source_path: Path, working_path: Path, mode: str) -> None:
    """SaveCopyAs the source into a brand-new working copy path via Excel COM.

    The source session is always read-only and is always closed without
    saving, so the source workbook can never be edited by this step.
    """
    session = connect_source_readonly(source_path, mode)
    try:
        save_working_copy(session.source_workbook, working_path)
    finally:
        session.close()


@dataclass
class WorkingCopyHandle:
    """An open, editable working copy plus the engine plan steps run against."""

    engine: Any
    _session: Any
    _workbook: Any
    _path: Path

    def fingerprint(self) -> WorkbookFingerprint:
        return collect_fingerprint(self._workbook, self._path)

    def save_and_close(self) -> None:
        self._workbook.Save()
        self._session.close_workbook(self._workbook, save_changes=False)
        self._session.close()

    def discard(self) -> None:
        """Close without saving; used on every failure path."""
        try:
            self._session.close_workbook(self._workbook, save_changes=False)
        finally:
            self._session.close()


def open_working_copy_for_edit(working_path: Path) -> WorkingCopyHandle:
    session = ExcelSession.create()
    try:
        workbook = session.open_workbook(working_path, read_only=False, update_links=False)
        session.capture_state()
        session.apply_editing_state()
        engine = ExcelComEngine(workbook, session=None, workbook_id="working-copy", read_only=False)
        return WorkingCopyHandle(engine=engine, _session=session, _workbook=workbook, _path=working_path)
    except Exception:
        session.close()
        raise


@dataclass
class ReopenedHandle:
    """A closed-then-reopened working copy, read for post-save validation."""

    _session: Any
    _workbook: Any
    _path: Path

    def fingerprint(self) -> WorkbookFingerprint:
        return collect_fingerprint(self._workbook, self._path)

    def close(self) -> None:
        try:
            self._session.close_workbook(self._workbook, save_changes=False)
        finally:
            self._session.close()


def reopen_working_copy(working_path: Path) -> ReopenedHandle:
    session = ExcelSession.create()
    try:
        workbook = session.open_workbook(working_path, read_only=False, update_links=False)
        return ReopenedHandle(_session=session, _workbook=workbook, _path=working_path)
    except Exception:
        session.close()
        raise


def compare_generic_preservation(
    before: WorkbookFingerprint, after: WorkbookFingerprint
) -> list[ValidationCheck]:
    """Structural preservation checks with no assumption about file format.

    Unlike ``workbook_audit.compare_preservation`` (which hard-requires the
    P&L workflow's XLSB format), this compares the *before* and *after*
    format for equality so the generic coordinator works for any Excel
    format the probe recognizes.
    """
    checks = [
        ValidationCheck(
            "sheet_names_preserved",
            before.sheet_names == after.sheet_names,
            True,
            "Sheet names are unchanged" if before.sheet_names == after.sheet_names else "Sheet names changed",
            {"before": before.sheet_names, "after": after.sheet_names},
        ),
        ValidationCheck(
            "sheet_count_preserved",
            before.sheet_count == after.sheet_count,
            True,
            "Sheet count is unchanged" if before.sheet_count == after.sheet_count else "Sheet count changed",
            {"before": before.sheet_count, "after": after.sheet_count},
        ),
        ValidationCheck(
            "file_format_preserved",
            before.file_format == after.file_format,
            True,
            "Workbook file format is unchanged" if before.file_format == after.file_format else "Workbook file format changed",
            {"before": before.file_format, "after": after.file_format},
        ),
        ValidationCheck(
            "external_links_preserved",
            before.external_links == after.external_links,
            True,
            "External link sources are unchanged" if before.external_links == after.external_links else "External link sources changed",
            {"before": before.external_links, "after": after.external_links},
        ),
    ]
    for name, before_value, after_value in (
        ("defined_name_count_preserved", before.defined_name_count, after.defined_name_count),
        ("connection_count_preserved", before.connection_count, after.connection_count),
        ("pivot_counts_preserved", before.pivot_counts, after.pivot_counts),
        ("vba_presence_preserved", before.has_vba_project, after.has_vba_project),
    ):
        if before_value is None or after_value is None:
            checks.append(
                ValidationCheck(name, True, False, "Optional preservation fact was unavailable", {})
            )
        else:
            checks.append(
                ValidationCheck(
                    name,
                    before_value == after_value,
                    True,
                    "Optional preservation fact is unchanged" if before_value == after_value else "Optional preservation fact changed",
                    {"before": before_value, "after": after_value},
                )
            )
    return checks
