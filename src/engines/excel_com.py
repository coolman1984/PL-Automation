"""Read-oriented Excel COM adapter behind the engine-independent contract.

This module does not import pywin32 at module load time.  That keeps probes,
documentation commands, and unit tests usable on non-Windows machines.  The
existing :class:`src.excel_session.ExcelSession` remains responsible for COM
initialization and lifecycle ownership.
"""

from __future__ import annotations

import re
import time
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

    @staticmethod
    def _optional(collection: object, name: str) -> object | None:
        try:
            return collection(name)  # type: ignore[operator]
        except Exception:
            return None

    @staticmethod
    def _range_address(cell_range: object) -> str:
        return str(cell_range.Address).replace("$", "")

    def _require_write(self) -> None:
        if self.read_only:
            raise PermissionError("The Excel engine is read-only")

    def _wait_for_calculation(self, timeout_seconds: int) -> None:
        deadline = time.monotonic() + timeout_seconds
        application = self.workbook.Application
        while _safe_int(getattr(application, "CalculationState", 0)) != 0:
            if time.monotonic() >= deadline:
                raise TimeoutError("Excel calculation did not finish before the timeout")
            time.sleep(0.05)

    def _wait_for_connection(self, connection: object, timeout_seconds: int) -> None:
        deadline = time.monotonic() + timeout_seconds
        while True:
            refreshing = False
            for property_name in ("OLEDBConnection", "ODBCConnection"):
                try:
                    refreshing = refreshing or bool(getattr(getattr(connection, property_name), "Refreshing"))
                except Exception:
                    pass
            if not refreshing:
                return
            if time.monotonic() >= deadline:
                raise TimeoutError("Connection refresh did not finish before the timeout")
            time.sleep(0.05)

    def inspect_advanced(
        self, tool: str, target: TargetRef | None, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        if tool == "format_range":
            assert target is not None
            cell_range = self._resolve_target(target)
            return {
                "address": self._range_address(cell_range),
                "format": {
                    "font_name": cell_range.Font.Name,
                    "font_size": cell_range.Font.Size,
                    "bold": cell_range.Font.Bold,
                    "italic": cell_range.Font.Italic,
                    "font_color": cell_range.Font.Color,
                    "fill_color": cell_range.Interior.Color,
                    "number_format": cell_range.NumberFormat,
                    "horizontal_alignment": cell_range.HorizontalAlignment,
                    "vertical_alignment": cell_range.VerticalAlignment,
                    "wrap_text": cell_range.WrapText,
                    "row_height": cell_range.RowHeight,
                    "column_width": cell_range.ColumnWidth,
                },
            }
        if tool == "insert_rows":
            assert target is not None
            return {"address": target.address, "formula_error_count": self.count_formula_errors(str(target.sheet))}
        if tool == "manage_sheet":
            assert target is not None
            worksheet = self._optional(self.workbook.Worksheets, str(target.sheet))
            value = None
            empty = None
            if worksheet is not None:
                used = worksheet.UsedRange
                value = {"visibility": _safe_int(worksheet.Visible), "used_address": self._range_address(used)}
                empty = _safe_int(used.Cells.CountLarge, _safe_int(used.Cells.Count)) == 1 and used.Value2 is None
            return {
                "sheet_names": [str(self.workbook.Worksheets(i).Name) for i in range(1, _safe_int(self.workbook.Worksheets.Count) + 1)],
                "sheet": value,
                "empty": empty,
            }
        if tool == "manage_table":
            assert target is not None
            table = self._optional(self._worksheet(str(target.sheet)).ListObjects, str(target.object_name))
            return {"current": None if table is None else {"name": str(table.Name), "address": self._range_address(table.Range)}}
        if tool == "manage_filter":
            assert target is not None
            worksheet = self._worksheet(str(target.sheet))
            return {"current": {"auto_filter_mode": bool(worksheet.AutoFilterMode), "filter_mode": bool(worksheet.FilterMode)}}
        if tool == "manage_validation":
            assert target is not None
            try:
                validation = self._resolve_target(target).Validation
                current = {"type": _safe_int(validation.Type), "formula1": str(validation.Formula1), "formula2": str(validation.Formula2)}
            except Exception:
                current = None
            return {"current": current}
        if tool == "manage_comment":
            assert target is not None
            comment = getattr(self._resolve_target(target), "Comment", None)
            return {"current": None if comment is None else {"text": str(comment.Text())}}
        if tool == "manage_hyperlink":
            assert target is not None
            links = self._resolve_target(target).Hyperlinks
            link = links(1) if _safe_int(links.Count) else None
            return {"current": None if link is None else {"address": str(link.Address or ""), "sub_address": str(link.SubAddress or "")}}
        if tool == "manage_chart":
            assert target is not None
            chart = self._optional(self._worksheet(str(target.sheet)).ChartObjects(), str(arguments["name"]))
            return {"current": None if chart is None else {"name": str(chart.Name), "chart_type": _safe_int(chart.Chart.ChartType)}}
        if tool == "manage_name":
            name = self._optional(self.workbook.Names, str(arguments["name"]))
            return {"current": None if name is None else {"name": str(name.Name), "refers_to": str(name.RefersTo)}}
        if tool == "manage_connection":
            connection = self._optional(self.workbook.Connections, str(arguments["name"]))
            return {"current": None if connection is None else {"name": str(connection.Name), "type": _safe_int(connection.Type)}}
        if tool == "refresh_workbook":
            return {"connections": list(arguments.get("connection_names", [])), "pivot_tables": list(arguments.get("pivot_tables", []))}
        if tool == "calculate_workbook":
            return {"calculation_state": _safe_int(self.workbook.Application.CalculationState)}
        if tool == "validate_workbook":
            sheet_names = [str(self.workbook.Worksheets(i).Name) for i in range(1, _safe_int(self.workbook.Worksheets.Count) + 1)]
            checks: dict[str, bool] = {}
            expected = arguments.get("expected_sheet_names")
            if expected is not None:
                checks["sheet_names"] = sorted(sheet_names) == sorted(expected)
            checks["tables"] = all(
                self._optional(self._worksheet(str(item["sheet"])).ListObjects, str(item["name"])) is not None
                for item in arguments.get("required_tables", [])
            )
            checks["charts"] = all(
                self._optional(self._worksheet(str(item["sheet"])).ChartObjects(), str(item["name"])) is not None
                for item in arguments.get("required_charts", [])
            )
            checks["names"] = all(self._optional(self.workbook.Names, str(name)) is not None for name in arguments.get("required_names", []))
            formula_error_count = 0
            if arguments.get("require_no_formula_errors", False):
                formula_error_count = sum(self.count_formula_errors(name) for name in sheet_names)
                checks["formula_errors"] = formula_error_count == 0
            return {"passed": all(checks.values()), "checks": checks, "formula_error_count": formula_error_count, "sheet_names": sheet_names}
        raise ValueError(f"Unsupported advanced tool: {tool}")

    def execute_advanced(
        self, tool: str, target: TargetRef | None, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        self._require_write()
        operation = arguments.get("operation")
        if tool == "format_range":
            assert target is not None
            cell_range = self._resolve_target(target)
            patch = arguments["format"]
            direct = {"font_name": (cell_range.Font, "Name"), "font_size": (cell_range.Font, "Size"), "bold": (cell_range.Font, "Bold"), "italic": (cell_range.Font, "Italic"), "font_color": (cell_range.Font, "Color"), "fill_color": (cell_range.Interior, "Color"), "number_format": (cell_range, "NumberFormat"), "wrap_text": (cell_range, "WrapText"), "row_height": (cell_range.EntireRow, "RowHeight"), "column_width": (cell_range.EntireColumn, "ColumnWidth")}
            for key, (owner, property_name) in direct.items():
                if key in patch:
                    setattr(owner, property_name, patch[key])
            if "horizontal_alignment" in patch:
                cell_range.HorizontalAlignment = {"general": 1, "left": -4131, "center": -4108, "right": -4152}[patch["horizontal_alignment"]]
            if "vertical_alignment" in patch:
                cell_range.VerticalAlignment = {"top": -4160, "center": -4108, "bottom": -4107}[patch["vertical_alignment"]]
        elif tool == "insert_rows":
            assert target is not None
            worksheet = self._worksheet(str(target.sheet))
            first_row, _, _, _ = self.resolve_bounds(target)
            insert_range = worksheet.Range(worksheet.Rows(first_row), worksheet.Rows(first_row + int(arguments["count"]) - 1))
            insert_range.Insert(Shift=-4121, CopyOrigin=0)
            worksheet.Calculate()
        elif tool == "manage_sheet":
            assert target is not None
            name = str(target.sheet)
            worksheet = self._optional(self.workbook.Worksheets, name)
            if operation == "create":
                if worksheet is not None:
                    raise ValueError(f"Sheet already exists: {name}")
                worksheet = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
                worksheet.Name = name
            elif worksheet is None:
                raise ValueError(f"Worksheet was not found: {name}")
            elif operation == "rename":
                worksheet.Name = arguments["new_name"]
            elif operation == "set_visibility":
                visibility = arguments["visibility"]
                visible_count = sum(_safe_int(self.workbook.Worksheets(i).Visible) == -1 for i in range(1, _safe_int(self.workbook.Worksheets.Count) + 1))
                if visibility != "visible" and _safe_int(worksheet.Visible) == -1 and visible_count <= 1:
                    raise ValueError("Cannot hide the last visible worksheet")
                worksheet.Visible = {"visible": -1, "hidden": 0, "very_hidden": 2}[visibility]
            elif operation == "delete_empty":
                used = worksheet.UsedRange
                empty = _safe_int(used.Cells.CountLarge, _safe_int(used.Cells.Count)) == 1 and used.Value2 is None
                if not empty:
                    raise ValueError("Worksheet is not empty")
                alerts = bool(self.workbook.Application.DisplayAlerts)
                try:
                    self.workbook.Application.DisplayAlerts = False
                    worksheet.Delete()
                finally:
                    self.workbook.Application.DisplayAlerts = alerts
        elif tool == "manage_table":
            assert target is not None
            worksheet = self._worksheet(str(target.sheet))
            table = self._optional(worksheet.ListObjects, str(target.object_name))
            if operation == "create":
                if table is not None:
                    raise ValueError(f"Table already exists: {target.object_name}")
                table = worksheet.ListObjects.Add(1, self._resolve_target(target), None, 1 if arguments.get("has_headers", True) else 0)
                table.Name = target.object_name
            elif table is None:
                raise ValueError(f"Table was not found: {target.object_name}")
            elif operation == "resize":
                table.Resize(self._resolve_target(target))
            elif operation == "unlist":
                table.Unlist()
        elif tool == "manage_filter":
            assert target is not None
            worksheet = self._worksheet(str(target.sheet))
            if operation == "apply":
                kwargs: dict[str, Any] = {"Field": arguments["field"], "Criteria1": arguments["criteria1"]}
                if "operator" in arguments:
                    kwargs["Operator"] = arguments["operator"]
                if "criteria2" in arguments:
                    kwargs["Criteria2"] = arguments["criteria2"]
                self._resolve_target(target).AutoFilter(**kwargs)
            elif operation == "clear" and bool(worksheet.FilterMode):
                worksheet.ShowAllData()
            elif operation == "remove":
                worksheet.AutoFilterMode = False
        elif tool == "manage_validation":
            assert target is not None
            validation = self._resolve_target(target).Validation
            if operation == "delete":
                validation.Delete()
            else:
                if self.inspect_advanced(tool, target, arguments)["current"] is not None and not arguments.get("replace", False):
                    raise ValueError("Validation exists; set replace=true to replace it")
                try:
                    validation.Delete()
                except Exception:
                    pass
                validation.Add(Type={"whole": 1, "decimal": 2, "list": 3, "date": 4, "custom": 7}[arguments["validation_type"]], AlertStyle=1, Operator=int(arguments.get("operator", 1)), Formula1=arguments["formula1"], Formula2=arguments.get("formula2", ""))
        elif tool in {"manage_comment", "manage_hyperlink"}:
            assert target is not None
            cell = self._resolve_target(target)
            if tool == "manage_comment":
                if getattr(cell, "Comment", None) is not None:
                    cell.Comment.Delete()
                if operation == "set":
                    cell.AddComment(arguments["text"])
            else:
                cell.Hyperlinks.Delete()
                if operation == "set":
                    self._worksheet(str(target.sheet)).Hyperlinks.Add(Anchor=cell, Address=arguments.get("address", ""), SubAddress=arguments.get("sub_address", ""), TextToDisplay=arguments.get("display_text", ""))
        elif tool == "manage_chart":
            assert target is not None
            worksheet = self._worksheet(str(target.sheet))
            chart_object = self._optional(worksheet.ChartObjects(), str(arguments["name"]))
            if operation == "delete":
                if chart_object is None:
                    raise ValueError(f"Chart was not found: {arguments['name']}")
                chart_object.Delete()
            else:
                source = worksheet.Range(arguments["source_address"])
                if chart_object is None:
                    anchor = worksheet.Range(arguments.get("anchor_address") or target.address or "A1")
                    chart_object = worksheet.ChartObjects().Add(anchor.Left, anchor.Top, float(arguments.get("width", 480)), float(arguments.get("height", 280)))
                    chart_object.Name = arguments["name"]
                chart_object.Chart.ChartType = {"column": 51, "bar": 57, "line": 4, "pie": 5, "area": 1, "scatter": -4169}[arguments["chart_type"]]
                chart_object.Chart.SetSourceData(Source=source)
                if "title" in arguments:
                    chart_object.Chart.HasTitle = True
                    chart_object.Chart.ChartTitle.Text = arguments["title"]
        elif tool == "manage_name":
            existing = self._optional(self.workbook.Names, str(arguments["name"]))
            if existing is not None:
                existing.Delete()
            if operation == "set":
                self.workbook.Names.Add(Name=arguments["name"], RefersTo=arguments["refers_to"])
        elif tool == "manage_connection":
            connection = self._optional(self.workbook.Connections, str(arguments["name"]))
            if connection is None:
                raise ValueError(f"Connection was not found: {arguments['name']}")
            connection.Refresh()
            self._wait_for_connection(connection, int(arguments.get("timeout_seconds", 120)))
        elif tool == "refresh_workbook":
            timeout = int(arguments.get("timeout_seconds", 120))
            for name in arguments.get("connection_names", []):
                connection = self._optional(self.workbook.Connections, str(name))
                if connection is None:
                    raise ValueError(f"Connection was not found: {name}")
                connection.Refresh()
                self._wait_for_connection(connection, timeout)
            for item in arguments.get("pivot_tables", []):
                self._resolve_pivot_table(str(item["sheet"]), str(item["name"])).RefreshTable()
        elif tool == "calculate_workbook":
            if arguments.get("full_rebuild", False):
                self.workbook.Application.CalculateFullRebuild()
            else:
                self.workbook.Calculate()
            self._wait_for_calculation(int(arguments.get("timeout_seconds", 120)))
        else:
            raise ValueError(f"Unsupported advanced mutation: {tool}")
        return self.inspect_advanced(tool, target, arguments)

    def close(self, *, save: bool = False) -> None:
        if self.session is not None:
            close = getattr(self.session, "close", None)
            if callable(close):
                close()
            return
        close = getattr(self.workbook, "Close", None)
        if callable(close):
            close(SaveChanges=bool(save))
