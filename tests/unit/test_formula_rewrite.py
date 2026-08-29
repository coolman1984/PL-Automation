from src.formula_clone import (
    formula_contains_exact_quoted_version,
    rewrite_formula_exact_quoted_version,
)


def test_rewrites_exact_quoted_version_literals_only():
    formula = '=SUMIFS(A:A,B:B,"T08")+IF(C1="T08",1,0)'

    rewritten, count = rewrite_formula_exact_quoted_version(formula, "T08", "A08")

    assert rewritten == '=SUMIFS(A:A,B:B,"A08")+IF(C1="A08",1,0)'
    assert count == 2


def test_does_not_rewrite_partial_or_unquoted_tokens():
    formula = '=IF(C1="XT08",T08,IF(C2="T080",1,0))+\'T08 Sheet\'!A1'

    rewritten, count = rewrite_formula_exact_quoted_version(formula, "T08", "A08")

    assert rewritten == formula
    assert count == 0


def test_formula_contains_exact_quoted_version():
    assert formula_contains_exact_quoted_version('=IF(A1="T08",1,0)', "T08")
    assert not formula_contains_exact_quoted_version('=IF(A1="XT08",1,0)', "T08")

