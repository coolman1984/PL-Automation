from types import SimpleNamespace

import pytest

from src.agent_contracts import TargetRef
from src.engines.excel_com import ExcelComEngine
from src.constants import XL_PASTE_FORMATS


class _Range:
    Address = "$A$1:$B$2"
    Value2 = ((1, "A"), (2, "B"))
    Formula = (("=1", "=\"A\""), ("=2", "=\"B\""))


class _Sheet:
    Name = "Data"
    ProtectContents = False
    Visible = -1
    AutoFilterMode = False
    FilterMode = False
    UsedRange = SimpleNamespace(
        Address="$A$1:$B$2",
        Row=1,
        Column=1,
        Rows=SimpleNamespace(Count=2),
        Columns=SimpleNamespace(Count=2),
    )

    def Range(self, address):
        assert address == "A1:B2"
        return _Range()


class _Sheets:
    Count = 1

    def __call__(self, key):
        assert key in (1, "Data")
        return _Sheet()


class _Workbook:
    Name = "book.xlsx"
    FullName = "C:\\work\\book.xlsx"
    FileFormat = 51
    ReadOnly = True
    Saved = True
    Worksheets = _Sheets()


def test_com_adapter_reads_explicit_target_without_importing_pywin32():
    engine = ExcelComEngine(_Workbook(), workbook_id="working-copy")
    target = TargetRef("working-copy", sheet="Data", address="A1:B2")

    assert engine.read_values(target) == [[1, "A"], [2, "B"]]
    assert engine.read_formulas(target)[0][0] == "=1"
    assert engine.inspect()["sheets"][0]["used_range"]["cell_count"] == 4


def test_com_adapter_rejects_unknown_sheet_and_read_only_write():
    engine = ExcelComEngine(_Workbook(), workbook_id="working-copy")
    target = TargetRef("working-copy", sheet="Data", address="A1")

    with pytest.raises(PermissionError):
        engine.write_values(target, [[1]])


def test_format_copy_uses_excel_enum_and_clears_copy_mode():
    calls = []

    class Range:
        Value2 = ((1,),)
        Formula = (("=1",),)

        def __init__(self, name):
            self.name = name

        def Copy(self, Destination=None):
            calls.append(("copy", self.name, Destination))

        def PasteSpecial(self, Paste=None):
            calls.append(("paste", Paste))

    source, destination = Range("source"), Range("destination")

    class Sheet(_Sheet):
        def Range(self, address):
            return source if address == "A1" else destination

    application = SimpleNamespace(CutCopyMode=True)
    workbook = SimpleNamespace(
        Worksheets=lambda _key: Sheet(),
        Application=application,
        Name="book.xlsx",
    )
    engine = ExcelComEngine(workbook, workbook_id="working-copy", read_only=False)

    engine.copy_range(
        TargetRef("working-copy", sheet="Data", address="A1"),
        TargetRef("working-copy", sheet="Data", address="B1"),
        mode="formats",
    )

    assert ("paste", XL_PASTE_FORMATS) in calls
    assert application.CutCopyMode is False


def test_mutation_range_validation_rejects_multi_area_range():
    cell_range = SimpleNamespace(
        Areas=SimpleNamespace(Count=2),
        Rows=SimpleNamespace(Count=1),
        Columns=SimpleNamespace(Count=1),
        Row=1,
        Column=1,
    )

    class Sheet:
        Rows = SimpleNamespace(Count=1_048_576)
        Columns = SimpleNamespace(Count=16_384)

        def Range(self, _address):
            return cell_range

    workbook = SimpleNamespace(Worksheets=lambda _key: Sheet(), Name="book.xlsx")
    engine = ExcelComEngine(workbook, workbook_id="working-copy", read_only=False)

    with pytest.raises(ValueError, match="Multi-area"):
        engine.validate_bounded_range(
            TargetRef("working-copy", sheet="Data", address="A1,B2")
        )


def test_table_name_resolves_to_real_bounds_for_pivot_guards():
    table_range = SimpleNamespace(
        Row=2,
        Column=3,
        Rows=SimpleNamespace(Count=10),
        Columns=SimpleNamespace(Count=4),
    )

    class Sheet:
        Name = "DB File"

        def ListObjects(self, name):
            if name != "SalesTable":
                raise KeyError(name)
            return SimpleNamespace(Range=table_range)

    class Sheets:
        Count = 1

        def __call__(self, _key):
            return Sheet()

    workbook = SimpleNamespace(Worksheets=Sheets(), Name="book.xlsx")
    engine = ExcelComEngine(workbook, workbook_id="working-copy", read_only=False)

    assert engine.resolve_source_bounds("SalesTable") == ("DB File", 2, 3, 11, 6)


def test_formula_error_count_sums_every_area_and_does_not_fail_open():
    class Areas:
        Count = 2

        def __call__(self, index):
            return SimpleNamespace(CountLarge=2 if index == 1 else 3)

    errors = SimpleNamespace(Areas=Areas())
    sheet = SimpleNamespace(
        UsedRange=SimpleNamespace(SpecialCells=lambda *_args: errors)
    )
    workbook = SimpleNamespace(Worksheets=lambda _key: sheet, Name="book.xlsx")
    engine = ExcelComEngine(workbook, workbook_id="working-copy", read_only=False)

    assert engine.count_formula_errors("Data") == 5

    sheet.UsedRange = SimpleNamespace(
        SpecialCells=lambda *_args: (_ for _ in ()).throw(OSError("permission denied"))
    )
    with pytest.raises(RuntimeError, match="Could not count"):
        engine.count_formula_errors("Data")
