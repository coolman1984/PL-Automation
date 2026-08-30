"""Read-oriented Excel COM adapter behind the engine-independent contract.

This module does not import pywin32 at module load time.  That keeps probes,
documentation commands, and unit tests usable on non-Windows machines.  The
existing :class:`src.excel_session.ExcelSession` remains responsible for COM
initialization and lifecycle ownership.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from typing import Any

from ..agent_contracts import TargetRef
from ..constants import (
    XL_CELL_TYPE_FORMULAS,
    XL_DATABASE_SOURCE_TYPE,
    XL_ERRORS,
    XL_FORMAT_FROM_LEFT_OR_ABOVE,
    XL_PASTE_FORMATS,
    XL_TO_RIGHT,
)
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


_R1C1_SOURCE_RE = re.compile(
    r"^(?:'(?P<quoted_sheet>[^']+)'|(?P<sheet>[^!']+))!"
    r"R(?P<r1>\d+)C(?P<c1>\d+)(?::R(?P<r2>\d+)C(?P<c2>\d+))?$"
)


def _parse_r1c1_source(source_data: str | None) -> tuple[str, int, int, int, int] | None:
    """Parse Excel's canonical ``PivotCache.SourceData`` string.

    Excel always reports this in ``Sheet!R#C#[:R#C#]`` form regardless of the
    notation used when the source was set. Returns ``None`` if it cannot be
    parsed (e.g. a table name or an external/non-range source).
    """
    if not source_data:
        return None
    match = _R1C1_SOURCE_RE.match(source_data.strip())
    if not match:
        return None
    sheet = match.group("quoted_sheet") or match.group("sheet")
    r1, c1 = int(match.group("r1")), int(match.group("c1"))
    r2 = int(match.group("r2")) if match.group("r2") else r1
    c2 = int(match.group("c2")) if match.group("c2") else c1
    return (sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


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
        try:
            destination_range.PasteSpecial(Paste=XL_PASTE_FORMATS)
        finally:
            try:
                self.workbook.Application.CutCopyMode = False
            except Exception:
                pass

    def clear_range(self, target: TargetRef) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        # ClearContents only: never Clear() (which would also drop formatting)
        # and never a row/column delete.
        self._resolve_target(target).ClearContents()

    def resolve_bounds(self, target: TargetRef) -> tuple[int, int, int, int]:
        cell_range = self._resolve_target(target)
        first_row = _safe_int(cell_range.Row)
        first_col = _safe_int(cell_range.Column)
        last_row = first_row + _safe_int(cell_range.Rows.Count) - 1
        last_col = first_col + _safe_int(cell_range.Columns.Count) - 1
        return first_row, first_col, last_row, last_col

    def validate_bounded_range(self, target: TargetRef) -> tuple[int, int, int, int]:
        """Resolve one finite rectangle and reject union/entire-axis ranges."""
        cell_range = self._resolve_target(target)
        areas = getattr(cell_range, "Areas", None)
        if _safe_int(getattr(areas, "Count", 1), 1) != 1:
            raise ValueError("Multi-area ranges are not supported for mutation")
        worksheet = self._worksheet(str(target.sheet))
        row_count = _safe_int(cell_range.Rows.Count)
        column_count = _safe_int(cell_range.Columns.Count)
        if row_count >= _safe_int(worksheet.Rows.Count) or column_count >= _safe_int(worksheet.Columns.Count):
            raise ValueError("Whole-row and whole-column ranges are not supported for mutation")
        first_row = _safe_int(cell_range.Row)
        first_col = _safe_int(cell_range.Column)
        return (
            first_row,
            first_col,
            first_row + row_count - 1,
            first_col + column_count - 1,
        )

    def calculate_sheet(self, sheet: str) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        self._worksheet(sheet).Calculate()

    def fill_formula_down(self, template: TargetRef, target: TargetRef) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        if template.sheet != target.sheet:
            raise ValueError("The template and target ranges must be on the same sheet")
        template_range = self._resolve_target(template)
        target_range = self._resolve_target(target)
        worksheet = self._worksheet(str(template.sheet))
        # Excel's own FillDown propagates relative/absolute/structured
        # references with true Excel semantics; the combined range's first
        # row(s) are the source and are never themselves overwritten.
        combined = worksheet.Range(template_range, target_range)
        combined.FillDown()

    def insert_columns(self, target: TargetRef, count: int) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        anchor_range = self._resolve_target(target)
        worksheet = self._worksheet(str(target.sheet))
        insert_at_col = _safe_int(anchor_range.Column)
        # pywin32's late-bound dynamic dispatch does not accept named keyword
        # arguments such as Resize(ColumnSize=count), so span the exact
        # entire-column range via Range(col1, col2) instead of Resize.
        first_column = worksheet.Columns(insert_at_col)
        last_column = worksheet.Columns(insert_at_col + count - 1)
        insert_range = worksheet.Range(first_column, last_column)
        insert_range.Insert(Shift=XL_TO_RIGHT, CopyOrigin=XL_FORMAT_FROM_LEFT_OR_ABOVE)

    def count_formula_errors(self, sheet: str) -> int:
        worksheet = self._worksheet(sheet)
        try:
            errors = worksheet.UsedRange.SpecialCells(XL_CELL_TYPE_FORMULAS, XL_ERRORS)
            areas = getattr(errors, "Areas", None)
            area_count = _safe_int(getattr(areas, "Count", 1), 1)
            total = 0
            for index in range(1, area_count + 1):
                area = areas(index) if area_count > 1 else errors
                total += _safe_int(getattr(area, "CountLarge", getattr(area, "Count", 0)))
            return total
        except Exception as exc:
            text = f"{exc!s} {getattr(exc, 'args', ())!r}".casefold()
            if "no cells were found" in text:
                return 0
            raise RuntimeError(f"Could not count formula errors on {sheet}: {exc}") from exc

    def _resolve_pivot_table(self, sheet: str, name: str) -> object:
        worksheet = self._worksheet(sheet)
        try:
            return worksheet.PivotTables(name)
        except Exception as exc:
            raise ValueError(f"PivotTable was not found: {sheet}!{name}") from exc

    def _pivot_sharing(self, cache_index: int, sheet: str, name: str) -> list[str]:
        sharing: list[str] = []
        sheet_count = _safe_int(self.workbook.Worksheets.Count)
        for sheet_index in range(1, sheet_count + 1):
            other_sheet = self.workbook.Worksheets(sheet_index)
            try:
                pivot_count = _safe_int(other_sheet.PivotTables().Count)
            except Exception:
                pivot_count = 0
            for pivot_index in range(1, pivot_count + 1):
                other_pivot = other_sheet.PivotTables(pivot_index)
                other_sheet_name = str(other_sheet.Name)
                other_pivot_name = str(other_pivot.Name)
                if other_sheet_name == sheet and other_pivot_name == name:
                    continue
                try:
                    other_cache_index = _safe_int(other_pivot.CacheIndex)
                except Exception:
                    continue
                if other_cache_index == cache_index:
                    sharing.append(f"{other_sheet_name}!{other_pivot_name}")
        return sharing

    def inspect_pivot_table(self, sheet: str, name: str) -> dict[str, Any]:
        pivot = self._resolve_pivot_table(sheet, name)
        cache = pivot.PivotCache()
        source_type_raw = _safe_int(getattr(cache, "SourceType", None), default=-1)
        try:
            source_data = str(cache.SourceData)
        except Exception:
            source_data = None
        cache_index = _safe_int(getattr(pivot, "CacheIndex", None), default=-1)
        return {
            "name": str(pivot.Name),
            "sheet": sheet,
            "cache_index": cache_index,
            "source_type": "database" if source_type_raw == XL_DATABASE_SOURCE_TYPE else source_type_raw,
            "source_data": source_data,
            # Excel always reports SourceData in R1C1 notation regardless of
            # the notation used to set it, so callers must compare identity
            # through resolve_source_bounds rather than the raw string.
            "source_bounds": self.resolve_source_bounds(source_data) if source_data else None,
            "shared_with": self._pivot_sharing(cache_index, sheet, name),
        }

    def resolve_source_bounds(self, address: str) -> tuple[str, int, int, int, int] | None:
        """Resolve an address string to a canonical (sheet, first_row,
        first_col, last_row, last_col) tuple, so it can be compared against
        ``inspect_pivot_table``'s ``source_bounds`` regardless of notation.

        Accepts either Excel's own R1C1 ``Sheet!R#C#[:R#C#]`` form (so a
        caller can round-trip a value ``inspect_pivot_table`` just reported)
        or an ordinary A1-style, optionally cross-sheet, address.
        """
        parsed = _parse_r1c1_source(address)
        if parsed is not None:
            return parsed
        table_name = address.strip()
        for sheet_index in range(1, _safe_int(self.workbook.Worksheets.Count) + 1):
            worksheet = self.workbook.Worksheets(sheet_index)
            try:
                resolved = worksheet.ListObjects(table_name).Range
            except Exception:
                continue
            first_row = _safe_int(resolved.Row)
            first_col = _safe_int(resolved.Column)
            return (
                str(worksheet.Name),
                first_row,
                first_col,
                first_row + _safe_int(resolved.Rows.Count) - 1,
                first_col + _safe_int(resolved.Columns.Count) - 1,
            )
        try:
            resolved = self.workbook.Application.Range(address)
        except Exception:
            return None
        first_row = _safe_int(resolved.Row)
        first_col = _safe_int(resolved.Column)
        last_row = first_row + _safe_int(resolved.Rows.Count) - 1
        last_col = first_col + _safe_int(resolved.Columns.Count) - 1
        return (str(resolved.Worksheet.Name), first_row, first_col, last_row, last_col)

    def update_pivot_source(self, sheet: str, name: str, new_source_address: str) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")
        pivot = self._resolve_pivot_table(sheet, name)
        # Resolve to an actual Range first: PivotCaches().Create accepts a
        # Range object for any notation, avoiding ambiguity in the raw string.
        new_source_range = None
        for sheet_index in range(1, _safe_int(self.workbook.Worksheets.Count) + 1):
            worksheet = self.workbook.Worksheets(sheet_index)
            try:
                new_source_range = worksheet.ListObjects(new_source_address.strip()).Range
                break
            except Exception:
                continue
        if new_source_range is None:
            new_source_range = self.workbook.Application.Range(new_source_address)
        new_cache = self.workbook.PivotCaches().Create(XL_DATABASE_SOURCE_TYPE, new_source_range)
        pivot.ChangePivotCache(new_cache)
        # A targeted refresh of only this PivotTable; never Application.RefreshAll.
        pivot.RefreshTable()

    def close(self, *, save: bool = False) -> None:
        if self.session is not None:
            close = getattr(self.session, "close", None)
            if callable(close):
                close()
            return
        close = getattr(self.workbook, "Close", None)
        if callable(close):
            close(SaveChanges=bool(save))
