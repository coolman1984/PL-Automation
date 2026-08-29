"""Excel COM lifecycle, ownership, and application-state management."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterator

import pythoncom
import win32com.client

from .constants import MSO_AUTOMATION_SECURITY_FORCE_DISABLE, XL_CALCULATION_MANUAL
from .errors import (
    DRMOpenBlockedError,
    ExcelConnectionError,
    ExcelNotInstalledError,
    NoRunningExcelError,
    WorkbookIdentityError,
    WorkbookNotFoundError,
    WorkbookReadOnlyError,
)
from .models import ExcelApplicationState


def normalize_windows_path(path: Path | str) -> str:
    return os.path.normcase(os.path.abspath(os.path.normpath(os.fspath(path))))


def iter_open_workbooks(app: object) -> Iterator[object]:
    try:
        count = int(app.Workbooks.Count)
        for index in range(1, count + 1):
            yield app.Workbooks(index)
    except Exception as exc:  # pragma: no cover - requires Excel
        raise ExcelConnectionError(f"Could not enumerate open Excel workbooks: {exc}") from exc


def find_open_workbook_by_full_path(app: object, path: Path) -> object | None:
    wanted = normalize_windows_path(path)
    matches = []
    for workbook in iter_open_workbooks(app):
        try:
            if normalize_windows_path(str(workbook.FullName)) == wanted:
                matches.append(workbook)
        except Exception:
            continue
    if len(matches) > 1:
        raise WorkbookIdentityError(
            f"More than one open workbook matched the exact path {path}"
        )
    return matches[0] if matches else None


def capture_application_state(app: object) -> ExcelApplicationState:
    try:
        return ExcelApplicationState(
            calculation=int(app.Calculation),
            screen_updating=bool(app.ScreenUpdating),
            enable_events=bool(app.EnableEvents),
            display_alerts=bool(app.DisplayAlerts),
            ask_to_update_links=bool(app.AskToUpdateLinks),
            visible=bool(app.Visible),
        )
    except Exception as exc:  # pragma: no cover - requires Excel
        raise ExcelConnectionError(f"Could not capture Excel application state: {exc}") from exc


def apply_editing_state(app: object) -> None:
    try:
        app.Calculation = XL_CALCULATION_MANUAL
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
    except Exception as exc:  # pragma: no cover - requires Excel
        raise ExcelConnectionError(f"Could not configure Excel for editing: {exc}") from exc


def restore_application_state(app: object, state: ExcelApplicationState) -> None:
    failures: list[str] = []
    for name, value in (
        ("Calculation", state.calculation),
        ("ScreenUpdating", state.screen_updating),
        ("EnableEvents", state.enable_events),
        ("DisplayAlerts", state.display_alerts),
        ("AskToUpdateLinks", state.ask_to_update_links),
        ("Visible", state.visible),
    ):
        try:
            setattr(app, name, value)
        except Exception as exc:  # pragma: no cover - requires Excel
            failures.append(f"{name}: {exc}")
    if failures:
        raise ExcelConnectionError(
            "Failed to restore one or more Excel settings: " + "; ".join(failures)
        )


class ExcelSession:
    """Context manager that never quits an Excel instance it did not create."""

    def __init__(self, app: object, *, owned_app: bool, mode: str):
        self.app = app
        self.owned_app = owned_app
        self.mode = mode
        self.application_state: ExcelApplicationState | None = None
        self._owned_workbooks: list[object] = []
        self._closed = False
        self._com_initialized = False

    @classmethod
    def attach(cls, source_path: Path) -> "ExcelSession":
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
        except Exception as exc:
            pythoncom.CoUninitialize()
            raise NoRunningExcelError(
                "No running Excel instance was available for attach mode"
            ) from exc
        try:
            workbook = find_open_workbook_by_full_path(app, source_path)
            if workbook is None:
                raise WorkbookNotFoundError(
                    f"The exact source workbook is not open in Excel: {source_path}"
                )
            session = cls(app, owned_app=False, mode="attach")
            session.source_workbook = workbook
            session._com_initialized = True
            return session
        except Exception:
            pythoncom.CoUninitialize()
            raise

    @classmethod
    def create(cls, *, visible: bool = False) -> "ExcelSession":
        pythoncom.CoInitialize()
        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            pythoncom.CoUninitialize()
            message = "Microsoft Excel desktop could not be started"
            if "class not registered" in str(exc).lower():
                raise ExcelNotInstalledError(message) from exc
            raise ExcelConnectionError(f"{message}: {exc}") from exc
        try:
            app.Visible = bool(visible)
            app.DisplayAlerts = False
            app.AskToUpdateLinks = False
            # Prevent Workbook_Open/auto-run code while an automation-owned
            # workbook is opened.  Attach mode never changes the user's app.
            app.EnableEvents = False
            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception as exc:  # pragma: no cover - requires Excel
            try:
                app.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
            raise ExcelConnectionError(f"Could not initialize isolated Excel: {exc}") from exc
        session = cls(app, owned_app=True, mode="open")
        session._com_initialized = True
        return session

    def __enter__(self) -> "ExcelSession":
        if not self._com_initialized:
            pythoncom.CoInitialize()
            self._com_initialized = True
        return self

    def open_workbook(
        self,
        path: Path,
        *,
        read_only: bool,
        update_links: bool = False,
    ) -> object:
        existing = find_open_workbook_by_full_path(self.app, path)
        if existing is not None:
            if not read_only:
                try:
                    if bool(existing.ReadOnly):
                        raise WorkbookReadOnlyError(f"Workbook is read-only: {path}")
                except WorkbookReadOnlyError:
                    raise
                except Exception:
                    pass
            return existing
        if not path.exists():
            raise WorkbookNotFoundError(f"Workbook was not found: {path}")
        try:
            workbook = self.app.Workbooks.Open(
                str(path),
                UpdateLinks=1 if update_links else 0,
                ReadOnly=bool(read_only),
                AddToMru=False,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
            )
        except Exception as exc:  # pragma: no cover - requires Excel
            text = str(exc)
            if any(token in text.lower() for token in ("protected", "drm", "authorization", "access denied")):
                raise DRMOpenBlockedError(
                    f"Excel/NASCA blocked opening the workbook. Open it manually and use attach mode: {path}"
                ) from exc
            raise ExcelConnectionError(f"Excel could not open workbook {path}: {exc}") from exc
        if not read_only:
            try:
                if bool(workbook.ReadOnly):
                    workbook.Close(SaveChanges=False)
                    raise WorkbookReadOnlyError(f"Workbook opened read-only: {path}")
            except WorkbookReadOnlyError:
                raise
            except Exception:
                pass
        self._owned_workbooks.append(workbook)
        return workbook

    def close_workbook(self, workbook: object, *, save_changes: bool) -> None:
        try:
            workbook.Close(SaveChanges=bool(save_changes))
        finally:
            self._owned_workbooks = [item for item in self._owned_workbooks if item is not workbook]

    def capture_state(self) -> ExcelApplicationState:
        self.application_state = capture_application_state(self.app)
        return self.application_state

    def apply_editing_state(self) -> None:
        apply_editing_state(self.app)

    def restore_state(self) -> None:
        if self.application_state is not None:
            restore_application_state(self.app, self.application_state)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        cleanup_errors: list[str] = []
        for workbook in list(self._owned_workbooks):
            try:
                workbook.Close(SaveChanges=False)
            except Exception as exc:  # pragma: no cover - requires Excel
                cleanup_errors.append(f"workbook close: {exc}")
        self._owned_workbooks.clear()
        if self.owned_app:
            try:
                self.app.Quit()
            except Exception as exc:  # pragma: no cover - requires Excel
                cleanup_errors.append(f"Excel quit: {exc}")
        if self._com_initialized:
            pythoncom.CoUninitialize()
            self._com_initialized = False
        if cleanup_errors:
            raise ExcelConnectionError("Excel cleanup failed: " + "; ".join(cleanup_errors))

    def __exit__(self, exc_type, exc, tb):
        try:
            self.restore_state()
        finally:
            self.close()
        return False
