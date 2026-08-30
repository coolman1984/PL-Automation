"""Small deterministic workbook engine used by contract tests.

It intentionally supports only rectangular values/formulas.  Its purpose is
to prove tool contracts and safety behavior without requiring Windows Excel.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import Any, Sequence

from .agent_contracts import TargetRef
from .engine_contract import EngineCapabilities


_CELL_RE = re.compile(r"^([A-Z]{1,3})([1-9][0-9]*)$")


def _column_number(value: str) -> int:
    number = 0
    for char in value:
        number = number * 26 + ord(char) - 64
    return number


def _address(value: str) -> tuple[int, int, int, int]:
    parts = value.upper().replace("$", "").split(":")
    if len(parts) == 1:
        parts *= 2
    match_one = _CELL_RE.match(parts[0])
    match_two = _CELL_RE.match(parts[1])
    if not match_one or not match_two:
        raise ValueError(f"Unsupported test range address: {value}")
    first_col = _column_number(match_one.group(1))
    last_col = _column_number(match_two.group(1))
    first_row = int(match_one.group(2))
    last_row = int(match_two.group(2))
    if last_col < first_col or last_row < first_row:
        raise ValueError(f"Range is reversed: {value}")
    return first_row, first_col, last_row, last_col


@dataclass
class FakeEngine:
    """In-memory workbook with explicit call history for assertions."""

    values: dict[str, dict[str, list[list[Any]]]]
    formulas: dict[str, dict[str, list[list[Any]]]] | None = None
    pivot_tables: dict[str, dict[str, Any]] | None = None
    features: dict[str, Any] | None = None
    """Keyed by ``"Sheet!PivotName"``; each value has at least
    ``source_type`` ("database" or something else), ``source_data``, and
    ``shared_with`` (a list of other ``"Sheet!Name"`` strings)."""

    def __post_init__(self) -> None:
        self.values = copy.deepcopy(self.values)
        self.formulas = copy.deepcopy(self.formulas or {})
        self.pivot_tables = copy.deepcopy(self.pivot_tables or {})
        self.features = copy.deepcopy(self.features or {})
        for name in (
            "formats", "tables", "filters", "validations", "comments",
            "hyperlinks", "charts", "names", "connections",
        ):
            self.features.setdefault(name, {})
        self.features.setdefault(
            "sheets", {name: {"visibility": "visible"} for name in self.values}
        )
        self.calls: list[tuple[str, str]] = []

    @property
    def capabilities(self) -> EngineCapabilities:
        return EngineCapabilities("fake", can_read=True, can_write=True)

    def _sheet(self, target: TargetRef) -> str:
        if not target.sheet:
            raise ValueError("A sheet is required for the fake engine")
        if target.sheet not in self.values:
            raise KeyError(f"Sheet not found: {target.sheet}")
        return target.sheet

    def _read(self, store: dict[str, dict[str, list[list[Any]]]], target: TargetRef) -> list[list[Any]]:
        sheet = self._sheet(target)
        if not target.address:
            raise ValueError("An address is required for the fake engine")
        first_row, first_col, last_row, last_col = _address(target.address)
        source = store.setdefault(sheet, {})
        result: list[list[Any]] = []
        for row in range(first_row, last_row + 1):
            row_values: list[Any] = []
            for column in range(first_col, last_col + 1):
                row_values.append(source.get(f"{row},{column}", None))
            result.append(row_values)
        return result

    def _write(self, store: dict[str, dict[str, list[list[Any]]]], target: TargetRef, values: Sequence[Sequence[Any]]) -> None:
        sheet = self._sheet(target)
        if not target.address:
            raise ValueError("An address is required for the fake engine")
        first_row, first_col, last_row, last_col = _address(target.address)
        expected_rows = last_row - first_row + 1
        expected_columns = last_col - first_col + 1
        materialized = [list(row) for row in values]
        if len(materialized) != expected_rows or any(len(row) != expected_columns for row in materialized):
            raise ValueError("Values do not match target range shape")
        destination = store.setdefault(sheet, {})
        for row_offset, row_values in enumerate(materialized):
            for column_offset, value in enumerate(row_values):
                destination[f"{first_row + row_offset},{first_col + column_offset}"] = copy.deepcopy(value)

    def inspect(self) -> dict[str, Any]:
        return {"engine": "fake", "sheets": sorted(self.values), "closed": False}

    def read_values(self, target: TargetRef) -> Sequence[Sequence[Any]]:
        self.calls.append(("read_values", target.address or ""))
        return self._read(self.values, target)

    def read_formulas(self, target: TargetRef) -> Sequence[Sequence[Any]]:
        self.calls.append(("read_formulas", target.address or ""))
        return self._read(self.formulas, target)

    def write_values(self, target: TargetRef, values: Sequence[Sequence[Any]]) -> None:
        self.calls.append(("write_values", target.address or ""))
        self._write(self.values, target, values)

    def write_formulas(self, target: TargetRef, formulas: Sequence[Sequence[Any]]) -> None:
        self.calls.append(("write_formulas", target.address or ""))
        self._write(self.formulas, target, formulas)

    def copy_range(self, source: TargetRef, destination: TargetRef, *, mode: str) -> None:
        self.calls.append((f"copy_range:{mode}", destination.address or ""))
        if mode not in {"values", "formulas"}:
            raise ValueError("Fake engine supports values or formulas copy only")
        store = self.values if mode == "values" else self.formulas
        self._write(store, destination, self._read(store, source))

    def clear_range(self, target: TargetRef) -> None:
        self.calls.append(("clear_range", target.address or ""))
        sheet = self._sheet(target)
        if not target.address:
            raise ValueError("An address is required for the fake engine")
        first_row, first_col, last_row, last_col = _address(target.address)
        destination = self.values.setdefault(sheet, {})
        for row in range(first_row, last_row + 1):
            for column in range(first_col, last_col + 1):
                destination.pop(f"{row},{column}", None)

    def resolve_bounds(self, target: TargetRef) -> tuple[int, int, int, int]:
        self._sheet(target)
        if not target.address:
            raise ValueError("An address is required for the fake engine")
        return _address(target.address)

    def validate_bounded_range(self, target: TargetRef) -> tuple[int, int, int, int]:
        return self.resolve_bounds(target)

    def calculate_sheet(self, sheet: str) -> None:
        if sheet not in self.values:
            raise KeyError(f"Sheet not found: {sheet}")
        self.calls.append(("calculate_sheet", sheet))

    def fill_formula_down(self, template: TargetRef, target: TargetRef) -> None:
        """Copy the template row's formula strings verbatim into every target
        row. This fake does not simulate Excel's relative-reference row
        shifting; real fidelity is proven only against real Excel COM."""
        self.calls.append(("fill_formula_down", target.address or ""))
        template_formulas = self._read(self.formulas, template)[0]
        target_first_row, target_first_col, target_last_row, target_last_col = self.resolve_bounds(target)
        if (target_last_col - target_first_col + 1) != len(template_formulas):
            raise ValueError("Template/target column count mismatch")
        sheet = self._sheet(target)
        destination = self.formulas.setdefault(sheet, {})
        for row in range(target_first_row, target_last_row + 1):
            for offset, formula in enumerate(template_formulas):
                destination[f"{row},{target_first_col + offset}"] = formula

    def insert_columns(self, target: TargetRef, count: int) -> None:
        """Shift this sheet's own cell keys right of the anchor by ``count``.

        Does not simulate Excel's cross-sheet or cross-workbook reference
        rewriting; real fidelity is proven only against real Excel COM.
        """
        self.calls.append(("insert_columns", target.address or ""))
        sheet = self._sheet(target)
        _, insert_at_col, _, _ = self.resolve_bounds(target)
        for store in (self.values, self.formulas):
            sheet_store = store.get(sheet)
            if not sheet_store:
                continue
            shifted: dict[str, Any] = {}
            for key, value in sheet_store.items():
                row_str, col_str = key.split(",")
                row, col = int(row_str), int(col_str)
                new_col = col + count if col >= insert_at_col else col
                shifted[f"{row},{new_col}"] = value
            store[sheet] = shifted

    def count_formula_errors(self, sheet: str) -> int:
        """Always zero: this fake never recomputes or introduces real errors."""
        if sheet not in self.values:
            raise KeyError(f"Sheet not found: {sheet}")
        return 0

    def inspect_pivot_table(self, sheet: str, name: str) -> dict[str, Any]:
        key = f"{sheet}!{name}"
        if key not in self.pivot_tables:
            raise KeyError(f"PivotTable not found: {key}")
        info = copy.deepcopy(self.pivot_tables[key])
        info.setdefault("source_bounds", info.get("source_data"))
        return info

    def resolve_source_bounds(self, address: str) -> Any:
        """This fake does not simulate Excel's R1C1 address normalization;
        it treats the address string itself as the comparable identity."""
        return address

    def update_pivot_source(self, sheet: str, name: str, new_source_address: str) -> None:
        self.calls.append(("update_pivot_source", f"{sheet}!{name}"))
        key = f"{sheet}!{name}"
        if key not in self.pivot_tables:
            raise KeyError(f"PivotTable not found: {key}")
        self.pivot_tables[key]["source_data"] = new_source_address
        self.pivot_tables[key]["source_bounds"] = new_source_address
        self.pivot_tables[key]["refreshed"] = True

    @staticmethod
    def _advanced_key(target: TargetRef | None) -> str:
        if target is None:
            return ""
        return "!".join(item for item in (target.sheet, target.address, target.object_name) if item)

    def inspect_advanced(
        self, tool: str, target: TargetRef | None, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        key = self._advanced_key(target)
        assert self.features is not None
        bucket_name = {
            "format_range": "formats", "manage_table": "tables",
            "manage_filter": "filters", "manage_validation": "validations",
            "manage_comment": "comments", "manage_hyperlink": "hyperlinks",
            "manage_chart": "charts", "manage_name": "names",
            "manage_connection": "connections",
        }.get(tool)
        if tool == "manage_sheet":
            return {
                "sheet_names": sorted(self.values),
                "sheet": copy.deepcopy(self.features["sheets"].get(str(target.sheet))) if target else None,
                "empty": not bool(self.values.get(str(target.sheet), {})),
            }
        if tool == "insert_rows":
            return {"anchor": key, "formula_error_count": 0}
        if bucket_name:
            object_key = str(arguments.get("name")) if tool in {"manage_chart", "manage_name", "manage_connection"} else key
            if tool == "manage_table" and target is not None:
                object_key = f"{target.sheet}!{target.object_name}"
            return {"key": object_key, "current": copy.deepcopy(self.features[bucket_name].get(object_key))}
        if tool == "refresh_workbook":
            return {"connections": list(arguments.get("connection_names", [])), "pivot_tables": list(arguments.get("pivot_tables", []))}
        if tool == "calculate_workbook":
            return {"calculated": bool(self.features.get("calculated", False))}
        if tool == "validate_workbook":
            expected = arguments.get("expected_sheet_names")
            checks: dict[str, bool] = {}
            if expected is not None:
                checks["sheet_names"] = sorted(expected) == sorted(self.values)
            for kind, feature in (("tables", "tables"), ("charts", "charts")):
                required = arguments.get(f"required_{kind}", [])
                checks[kind] = all(
                    "!".join((str(item.get("sheet", "")), str(item.get("name", "")))) in self.features[feature]
                    for item in required
                )
            checks["names"] = all(name in self.features["names"] for name in arguments.get("required_names", []))
            if arguments.get("require_no_formula_errors", False):
                checks["formula_errors"] = True
            return {"passed": all(checks.values()), "checks": checks, "formula_error_count": 0}
        raise ValueError(f"Unsupported advanced tool: {tool}")

    def execute_advanced(
        self, tool: str, target: TargetRef | None, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        assert self.features is not None
        self.calls.append((tool, self._advanced_key(target)))
        if tool == "format_range":
            self.features["formats"][self._advanced_key(target)] = copy.deepcopy(arguments["format"])
        elif tool == "insert_rows":
            self.features.setdefault("inserted_rows", []).append({"target": self._advanced_key(target), "count": arguments["count"]})
        elif tool == "manage_sheet":
            assert target is not None
            operation = arguments["operation"]
            name = str(target.sheet)
            if operation == "create":
                if name in self.values:
                    raise ValueError(f"Sheet already exists: {name}")
                self.values[name] = {}
                self.formulas[name] = {}
                self.features["sheets"][name] = {"visibility": "visible"}
            elif operation == "rename":
                new_name = arguments["new_name"]
                self.values[new_name] = self.values.pop(name)
                self.formulas[new_name] = self.formulas.pop(name, {})
                self.features["sheets"][new_name] = self.features["sheets"].pop(name)
            elif operation == "set_visibility":
                self.features["sheets"][name]["visibility"] = arguments["visibility"]
            elif operation == "delete_empty":
                if self.values.get(name):
                    raise ValueError("Sheet is not empty")
                self.values.pop(name)
                self.formulas.pop(name, None)
                self.features["sheets"].pop(name, None)
        elif tool in {"manage_table", "manage_filter", "manage_validation", "manage_comment", "manage_hyperlink", "manage_chart", "manage_name"}:
            bucket = {
                "manage_table": "tables", "manage_filter": "filters",
                "manage_validation": "validations", "manage_comment": "comments",
                "manage_hyperlink": "hyperlinks", "manage_chart": "charts", "manage_name": "names",
            }[tool]
            key = str(arguments.get("name")) if tool in {"manage_chart", "manage_name"} else self._advanced_key(target)
            if tool == "manage_table" and target is not None:
                key = f"{target.sheet}!{target.object_name}"
            operation = arguments["operation"]
            if operation in {"delete", "remove", "unlist", "clear"}:
                self.features[bucket].pop(key, None)
            else:
                value = copy.deepcopy(arguments)
                if tool == "manage_table" and target is not None:
                    value["address"] = target.address
                if tool == "manage_comment":
                    value["text"] = arguments.get("text")
                if tool == "manage_hyperlink":
                    value["address"] = arguments.get("address")
                if tool == "manage_name":
                    value["refers_to"] = arguments.get("refers_to")
                self.features[bucket][key] = value
        elif tool == "manage_connection":
            name = arguments["name"]
            current = self.features["connections"].setdefault(name, {})
            current["refreshed"] = int(current.get("refreshed", 0)) + 1
        elif tool == "refresh_workbook":
            for name in arguments.get("connection_names", []):
                current = self.features["connections"].setdefault(name, {})
                current["refreshed"] = int(current.get("refreshed", 0)) + 1
            for item in arguments.get("pivot_tables", []):
                key = f"{item.get('sheet')}!{item.get('name')}"
                if key not in self.pivot_tables:
                    raise KeyError(f"PivotTable not found: {key}")
                self.pivot_tables[key]["refreshed"] = True
        elif tool == "calculate_workbook":
            self.features["calculated"] = True
        else:
            raise ValueError(f"Unsupported advanced mutation: {tool}")
        return self.inspect_advanced(tool, target, arguments)

    def close(self, *, save: bool = False) -> None:
        self.calls.append(("close", "save" if save else "discard"))
