from src.total_pl_updater import (
    candidate_has_business_lineage,
    extract_cross_sheet_a1_references,
    rewrite_business_source_columns,
)
from src.models import MonthBlock


class _Cell:
    def __init__(self, row, column):
        self.Row = row
        self.Column = column


class _Range:
    def __init__(self, formulas):
        self.Formula = tuple((value,) for value in formulas)


class _Worksheet:
    Name = "Total PL"

    def __init__(self, formulas):
        self._formulas = formulas

    def Cells(self, row, column):
        return _Cell(row, column)

    def Range(self, start, end):
        return _Range(self._formulas)


def _block(sheet, target_col):
    return MonthBlock(
        sheet=sheet,
        year=2026,
        month=8,
        period="2026.008",
        target_col=target_col,
        target_pct_col=target_col + 1,
        forecast_col=target_col + 2,
        forecast_pct_col=target_col + 3,
        insert_at_col=target_col + 4,
        version_header_row=15,
        period_header_row=12,
        month_header_row=14,
        month_merge=None,
        september_start_col=target_col + 4,
        last_used_row=20,
    )


def test_total_pl_candidate_requires_all_three_business_lineages():
    worksheet = _Worksheet([
        "='VD Total'!L10+'MX Total'!M10+'DA Total'!N10",
    ])
    assert candidate_has_business_lineage(
        worksheet,
        _block("Total PL", 12),
        {
            "VD Total": _block("VD Total", 12),
            "MX Total": _block("MX Total", 13),
            "DA Total": _block("DA Total", 14),
        },
    )


def test_total_pl_candidate_rejects_incomplete_business_lineage():
    worksheet = _Worksheet(["='VD Total'!L10+'MX Total'!M10"])
    assert not candidate_has_business_lineage(
        worksheet,
        _block("Total PL", 12),
        {
            "VD Total": _block("VD Total", 12),
            "MX Total": _block("MX Total", 13),
            "DA Total": _block("DA Total", 14),
        },
    )


def test_extracts_quoted_sheet_references_and_preserves_rows():
    formula = "='VD Total'!$XO$168+'MX Total'!WV202+'DA Total'!$WF$202"

    refs = extract_cross_sheet_a1_references(
        formula, ("VD Total", "MX Total", "DA Total")
    )

    assert [(ref.sheet, ref.column, ref.row) for ref in refs] == [
        ("VD Total", 639, 168),
        ("MX Total", 620, 202),
        ("DA Total", 604, 202),
    ]
    assert refs[0].absolute_column and refs[0].absolute_row


def test_rewrites_only_allowed_business_columns():
    formula = "='VD Total'!XO168+'MX Total'!WV202+'DA Total'!WF202"

    rewritten, refs = rewrite_business_source_columns(
        formula,
        {"VD Total": {639}, "MX Total": {620}, "DA Total": {604}},
        {"VD Total": 643, "MX Total": 624, "DA Total": 608},
    )

    assert rewritten == "='VD Total'!XS168+'MX Total'!WZ202+'DA Total'!WJ202"
    assert [(ref.sheet, ref.column, ref.row) for ref in refs] == [
        ("VD Total", 643, 168),
        ("MX Total", 624, 202),
        ("DA Total", 608, 202),
    ]
