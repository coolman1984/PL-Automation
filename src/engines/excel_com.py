"""Read-oriented Excel COM adapter behind the engine-independent contract.

This module does not import pywin32 at module load time.  That keeps probes,
documentation commands, and unit tests usable on non-Windows machines.  The
existing :class:`src.excel_session.ExcelSession` remains responsible for COM
initialization and lifecycle ownership.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from ..agent_contracts import TargetRef
from ..engine_contract import EngineCapabilities


def _as_matrix(value: Any) -> list[list[Any]]:
    """Normalize Excel's scalar/tuple variants to a JSON-safe matrix."""
    if value is None:
        return [[None]]
    if isinstance(value, tuple):
        if not value:
            return [[]]
        if all(isinstance(row, tuple) for row in value):
            return [list(row) for row in value]
        return [list(value)]
    if isinstance(value, list):
        if not value:
            return [[]]
        if all(isinstance(row, (list, tuple)) for row in value):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def _safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError, OverflowError):
        return default


class ExcelComEngine:
    """Explicit, selection-free adapter for a workbook already opened in Excel."""

    def __init__(
        self,
        workbook: object,
        *,
        session: object | None = None,
        workbook_id: str | None = None,
        read_only: bool = True,
    ) -> None:
        self.workbook = workbook
        self.session = session
        self.workbook_id = workbook_id or self._workbook_name()
        self.read_only = read_only

    def _workbook_name(self) -> str:
        try:
            return str(self.workbook.FullName)
        except Exception:
            try:
                return str(self.workbook.Name)
            except Exception:
                return "excel-workbook"

    @property
    def capabilities(self) -> EngineCapabilities:
        return EngineCapabilities(
            "excel_com",
            can_read=True,
            can_write=not self.read_only,
            supports_xlsb=True,
            supports_macros=True,
            supports_charts=True,
            supports_pivots=True,
            supports_external_links=True,
            notes=("Uses the user's authorized desktop Excel session.",),
        )

    def _resolve_target(self, target: TargetRef) -> object:
        if target.workbook_id not in {self.workbook_id, "working-copy", "source", ""}:
            raise ValueError(
                f"Target workbook does not match the opened workbook: {target.workbook_id}"
            )
        if not target.sheet:
            raise ValueError("An explicit worksheet name is required")
        try:
            worksheet = self.workbook.Worksheets(target.sheet)
        except Exception as exc:
            raise ValueError(f"Worksheet was not found: {target.sheet}") from exc
        if not target.address:
            raise ValueError("An explicit A1 range address is required")
        try:
            return worksheet.Range(target.address)
        except Exception as exc:
            raise ValueError(
                f"Range could not be resolved: {target.sheet}!{target.address}"
            ) from exc

    def _worksheet(self, name: str) -> object:
        try:
            return self.workbook.Worksheets(name)
        except Exception as exc:
            raise ValueError(f"Worksheet was not found: {name}") from exc

    def inspect(self) -> dict[str, Any]:
        sheets: list[dict[str, Any]] = []
        worksheets = self.workbook.Worksheets
        count = _safe_int(worksheets.Count)
        for index in range(1, count + 1):
            worksheet = worksheets(index)
            item: dict[str, Any] = {"name": str(worksheet.Name)}
            try:
                used = worksheet.UsedRange
                item["used_range"] = {
                    "address": str(used.Address),
                    "first_row": _safe_int(used.Row),
                    "first_column": _safe_int(used.Column),
                    "row_count": _safe_int(used.Rows.Count),
                    "column_count": _safe_int(used.Columns.Count),
                    "cell_count": _safe_int(used.Rows.Count) * _safe_int(used.Columns.Count),
                }
            except Exception as exc:
                item["used_range_error"] = str(exc)
            for property_name in ("ProtectContents", "Visible", "AutoFilterMode", "FilterMode"):
                try:
                    item[property_name.casefold()] = getattr(worksheet, property_name)
                except Exception:
                    pass
            sheets.append(item)
        result: dict[str, Any] = {
            "engine": "excel_com",
            "workbook_id": self.workbook_id,
            "read_only": self.read_only,
            "sheets": sheets,
        }
        for property_name in ("Name", "FullName", "FileFormat", "ReadOnly", "Saved"):
            try:
                result[property_name.casefold()] = getattr(self.workbook, property_name)
            except Exception:
                pass
        return result

    def read_values(self, target: TargetRef) -> Sequence[Sequence[Any]]:
        cell_range = self._resolve_target(target)
        return _as_matrix(cell_range.Value2)

    def read_formulas(self, target: TargetRef) -> Sequence[Sequence[Any]]:
        cell_range = self._resolve_target(target)
        try:
            value = cell_range.Formula2
        except Exception:
            value = cell_range.Formula
        return _as_matrix(value)

    def write_values(self, target: TargetRef, values: Sequence[Sequence[Any]]) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        self._resolve_target(target).Value2 = values

    def write_formulas(self, target: TargetRef, formulas: Sequence[Sequence[Any]]) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        cell_range = self._resolve_target(target)
        try:
            cell_range.Formula2 = formulas
        except Exception:
            cell_range.Formula = formulas

    def copy_range(self, source: TargetRef, destination: TargetRef, *, mode: str) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        source_range = self._resolve_target(source)
        destination_range = self._resolve_target(destination)
        if mode not in {"all", "values", "formulas", "formats"}:
            raise ValueError(f"Unsupported copy mode: {mode}")
        # Use Excel's own copy semantics so formulas, styles, merged areas, and
        # advanced workbook behavior are not reconstructed in Python.
        if mode == "all":
            source_range.Copy(Destination=destination_range)
            return
        if mode == "values":
            destination_range.Value2 = source_range.Value2
            return
        if mode == "formulas":
            try:
                destination_range.Formula2 = source_range.Formula2
            except Exception:
                destination_range.Formula = source_range.Formula
            return
        source_range.Copy()
        destination_range.PasteSpecial(Paste="formats")

    def close(self, *, save: bool = False) -> None:
        if self.session is not None:
            close = getattr(self.session, "close", None)
            if callable(close):
                close()
            return
        close = getattr(self.workbook, "Close", None)
        if callable(close):
            close(SaveChanges=bool(save))
