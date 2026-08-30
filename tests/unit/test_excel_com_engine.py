from types import SimpleNamespace

import pytest

from src.agent_contracts import TargetRef
from src.engines.excel_com import ExcelComEngine


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

