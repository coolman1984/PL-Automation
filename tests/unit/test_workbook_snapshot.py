from __future__ import annotations

from datetime import datetime
from types import SimpleNamespace

import pytest

import src.workbook_snapshot as snapshot_module
from src.workbook_snapshot import SnapshotOptions, _style_id, json_safe, resolve_snapshot_mode


def test_style_ids_are_stable_and_order_independent():
    left = {"font": {"bold": True}, "number_format": "0.00"}
    right = {"number_format": "0.00", "font": {"bold": True}}
    assert _style_id(left) == _style_id(right)


def test_json_safe_preserves_nested_values_and_dates():
    value = {"date": datetime(2026, 8, 29, 10, 30), "tuple": (1, "A")}
    assert json_safe(value) == {"date": "2026-08-29T10:30:00", "tuple": [1, "A"]}


def test_auto_snapshot_uses_full_below_limit_and_inventory_above_it():
    assert resolve_snapshot_mode("auto", 100, 100) == ("full", [])
    mode, warnings = resolve_snapshot_mode("auto", 101, 100)
    assert mode == "inventory"
    assert warnings


def test_full_snapshot_stops_above_explicit_cell_limit(monkeypatch, tmp_path):
    sheet = SimpleNamespace(Name="Sheet1")

    class Worksheets:
        Count = 1

        def __call__(self, _index):
            return sheet

    workbook = SimpleNamespace(Worksheets=Worksheets())
    monkeypatch.setattr(
        snapshot_module,
        "_sheet_inventory",
        lambda _sheet: {
            "name": "Sheet1",
            "used_range": {
                "first_row": 1,
                "first_column": 1,
                "row_count": 100,
                "column_count": 100,
                "cell_count": 10_000,
            },
        },
    )
    with pytest.raises(ValueError, match="above the safe limit"):
        snapshot_module.build_workbook_snapshot(
            workbook,
            tmp_path / "book.xlsx",
            SnapshotOptions(mode="full", max_cells=100),
        )


def test_small_auto_snapshot_records_cell_style_and_dimensions(monkeypatch, tmp_path):
    count = SimpleNamespace(Count=0)
    used = SimpleNamespace(
        Row=1,
        Column=1,
        Rows=SimpleNamespace(Count=1),
        Columns=SimpleNamespace(Count=1),
        MergeAreas=count,
        FormatConditions=count,
    )
    row = SimpleNamespace(RowHeight=20, Hidden=False, OutlineLevel=1)
    column = SimpleNamespace(ColumnWidth=12, Hidden=False, OutlineLevel=1)
    border = SimpleNamespace(LineStyle=1, Weight=2, Color=0, ColorIndex=1)
    cell = SimpleNamespace(
        Font=SimpleNamespace(Name="Arial", Size=10, Bold=True, Italic=False, Underline=0, Strikethrough=False, Color=255, ColorIndex=3),
        Interior=SimpleNamespace(Color=65535, ColorIndex=6, Pattern=1, PatternColor=0),
        Protection=SimpleNamespace(Locked=True, FormulaHidden=False),
        Borders=lambda _index: border,
        Style="Normal",
        NumberFormat="0.00",
        HorizontalAlignment=-4108,
        VerticalAlignment=-4108,
        WrapText=False,
        Orientation=0,
        IndentLevel=0,
        ShrinkToFit=False,
        Address="$A$1",
        Value2=10.5,
        Formula="=5+5.5",
        HasFormula=True,
        MergeCells=False,
        Validation=SimpleNamespace(Type=0, Operator=None, Formula1=None, Formula2=None, IgnoreBlank=True, InCellDropdown=True),
        Comment=None,
    )

    class Sheet:
        Name = "Sheet1"
        UsedRange = used
        Visible = -1
        ProtectContents = False
        Shapes = count
        ChartObjects = count
        ListObjects = count
        PivotTables = count
        Hyperlinks = count
        Comments = count
        AutoFilterMode = False
        FilterMode = False

        def Rows(self, _row):
            return row

        def Columns(self, _column):
            return column

        def Cells(self, _row, _column):
            return cell

    sheet = Sheet()

    class Worksheets:
        Count = 1

        def __call__(self, _key):
            return sheet

    workbook = SimpleNamespace(Worksheets=Worksheets(), Names=count, Connections=count)
    monkeypatch.setattr(
        snapshot_module,
        "collect_fingerprint",
        lambda _workbook, _path: SimpleNamespace(
            sha256="abc",
            size_bytes=10,
            file_format=50,
            sheet_count=1,
            sheet_names=["Sheet1"],
            defined_name_count=0,
            external_links=[],
            connection_count=0,
            pivot_counts={"Sheet1": 0},
            has_vba_project=False,
        ),
    )

    result = snapshot_module.build_workbook_snapshot(
        workbook,
        tmp_path / "book.xlsb",
        SnapshotOptions(mode="auto", max_cells=10),
    )
    assert result["snapshot_mode"] == "full"
    assert result["sheets"][0]["rows"][0]["height"] == 20
    assert result["sheets"][0]["columns"][0]["width"] == 12
    assert result["sheets"][0]["cells"][0]["formula"] == "=5+5.5"
    assert len(result["styles"]) == 1
