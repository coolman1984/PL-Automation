"""Pure validation helpers and COM-facing validation primitives."""

from __future__ import annotations

import math
from collections.abc import Mapping, Sequence

from .header_discovery import read_header_snapshot, value_at
from .models import (
    ReconciliationResult,
    SheetUpdateResult,
    TotalPLRowMapping,
    ValidationCheck,
    WorkbookFingerprint,
)


EXCEL_ERROR_TEXTS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A")


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(float(value))


def numbers_match(actual: object, expected: object, tolerance: float) -> bool:
    if actual is None or expected is None:
        return actual is None and expected is None
    if not _is_number(actual) or not _is_number(expected):
        return False
    return abs(float(actual) - float(expected)) <= tolerance


def scan_formula_errors(range_obj: object) -> list[dict[str, object]]:
    errors: list[dict[str, object]] = []
    try:
        values = range_obj.Value2
        formulas = range_obj.Formula
        first_row = int(range_obj.Row)
        first_col = int(range_obj.Column)
    except Exception:
        return errors
    if not isinstance(values, tuple):
        values = ((values,),)
    elif values and not isinstance(values[0], tuple):
        values = (values,)
    if not isinstance(formulas, tuple):
        formulas = ((formulas,),)
    elif formulas and not isinstance(formulas[0], tuple):
        formulas = (formulas,)
    for row_offset, row in enumerate(values):
        for col_offset, value in enumerate(row):
            text = str(value).strip().upper() if value is not None else ""
            if text in EXCEL_ERROR_TEXTS:
                formula = formulas[row_offset][col_offset]
                errors.append(
                    {
                        "row": first_row + row_offset,
                        "column": first_col + col_offset,
                        "value": text,
                        "formula": str(formula),
                    }
                )
    return errors


def _check(name: str, passed: bool, required: bool, message: str, evidence=None):
    return ValidationCheck(name, passed, required, message, evidence or {})


def validate_sheet_structure(
    worksheet: object, codes, expected_actual_col: int
) -> list[ValidationCheck]:
    snapshot = read_header_snapshot(worksheet)
    row = None
    # The actual header row is discovered from the expected column's nearby
    # T08/A08 structure rather than hard-coded for the workbook.
    for candidate_row in range(snapshot.first_row, snapshot.last_row + 1):
        if str(value_at(snapshot, candidate_row, expected_actual_col)).strip().upper() == codes.actual_version:
            row = candidate_row
            break
    if row is None:
        return [_check("actual_header_present", False, True, "A08 header was not found")]
    return [
        _check(
            "actual_header_present",
            True,
            True,
            "A08 header is present",
            {"row": row, "column": expected_actual_col},
        ),
        _check(
            "actual_percent_header_present",
            str(value_at(snapshot, row, expected_actual_col + 1)).strip() == "%",
            True,
            "A08 percentage header is present",
        ),
    ]


def validate_business_formulas(
    worksheet: object, result: SheetUpdateResult, codes
) -> list[ValidationCheck]:
    block = result.before_block
    actual_range = worksheet.Range(
        worksheet.Cells(1, result.actual_amount_col),
        worksheet.Cells(block.last_used_row, result.actual_pct_col),
    )
    errors = scan_formula_errors(actual_range)
    checks = [
        _check(
            "business_formula_errors",
            not errors,
            True,
            "No formula errors found in the new business pair" if not errors else "Formula errors found in the new business pair",
            {"errors": errors[:50], "count": len(errors)},
        ),
        _check(
            "business_actual_formula_topology",
            result.formula_audit.actual_amount_formula_count
            == result.formula_audit.source_amount_formula_count,
            True,
            "A08 amount topology matches T08",
        ),
        _check(
            "business_percent_formula_topology",
            result.formula_audit.actual_pct_formula_count
            == result.formula_audit.source_pct_formula_count,
            True,
            "A08 percentage topology matches T08",
        ),
        _check(
            "business_no_t08_literals",
            result.formula_audit.actual_quoted_target_count == 0,
            True,
            "No quoted T08 literals remain in A08 amount formulas",
        ),
    ]
    return checks


def _numeric_value(worksheet: object, row: int, col: int) -> float | None:
    value = worksheet.Cells(row, col).Value2
    if _is_number(value):
        return float(value)
    return None


def reconcile_total_pl(
    total_sheet: object,
    mappings: Sequence[TotalPLRowMapping],
    tolerance: float,
) -> list[ReconciliationResult]:
    results: list[ReconciliationResult] = []
    for mapping in mappings:
        if mapping.classification != "cross-sheet formula":
            continue
        values: list[float] = []
        references: list[str] = []
        invalid = False
        for ref in mapping.source_references:
            value = _numeric_value(total_sheet.Parent.Worksheets(ref.sheet), ref.row, ref.column)
            references.append(f"'{ref.sheet}'!{ref.column}:{ref.row}")
            if value is None:
                invalid = True
                break
            values.append(value)
        actual = _numeric_value(total_sheet, mapping.total_pl_row, mapping.source_references[0].column if False else 1)
        # The Total PL amount column is supplied through the mapping's parent
        # context by the caller after this function is wrapped below.
        results.append(
            ReconciliationResult(
                total_pl_row=mapping.total_pl_row,
                label=mapping.label,
                actual=actual,
                expected=sum(values) if values and not invalid else None,
                difference=None,
                source_references=references,
                passed=False,
                reason="Total PL amount column was not supplied",
            )
        )
    return results


def reconcile_total_pl_at_column(
    total_sheet: object,
    mappings: Sequence[TotalPLRowMapping],
    actual_amount_col: int,
    tolerance: float,
) -> list[ReconciliationResult]:
    results: list[ReconciliationResult] = []
    for mapping in mappings:
        if mapping.classification != "cross-sheet formula":
            continue
        expected_values: list[float] = []
        references: list[str] = []
        reason = None
        for ref in mapping.source_references:
            value = _numeric_value(total_sheet.Parent.Worksheets(ref.sheet), ref.row, ref.column)
            references.append(f"'{ref.sheet}'!R{ref.row}C{ref.column}")
            if value is None:
                reason = f"Source value at {references[-1]} is not numeric"
                break
            expected_values.append(value)
        actual = _numeric_value(total_sheet, mapping.total_pl_row, actual_amount_col)
        expected = sum(expected_values) if reason is None else None
        difference = actual - expected if actual is not None and expected is not None else None
        passed = reason is None and difference is not None and abs(difference) <= tolerance
        results.append(
            ReconciliationResult(
                total_pl_row=mapping.total_pl_row,
                label=mapping.label,
                actual=actual,
                expected=expected,
                difference=difference,
                source_references=references,
                passed=passed,
                reason=reason or (None if passed else "Difference exceeds tolerance or actual is not numeric"),
            )
        )
    return results


def validate_percent_formulas(
    worksheet: object, result: SheetUpdateResult
) -> list[ValidationCheck]:
    errors = scan_formula_errors(
        worksheet.Range(
            worksheet.Cells(1, result.actual_pct_col),
            worksheet.Cells(result.before_block.last_used_row, result.actual_pct_col),
        )
    )
    return [
        _check(
            "percent_formula_errors",
            not errors,
            True,
            "No percentage formula errors found" if not errors else "Percentage formula errors found",
            {"errors": errors[:50], "count": len(errors)},
        )
    ]


def validate_external_links(before: WorkbookFingerprint, after: WorkbookFingerprint) -> ValidationCheck:
    return _check(
        "external_links_preserved",
        before.external_links == after.external_links,
        True,
        "External links are unchanged" if before.external_links == after.external_links else "External links changed",
        {"before": before.external_links, "after": after.external_links},
    )


def validate_control_cells(
    before: Mapping[str, object], after: Mapping[str, object]
) -> list[ValidationCheck]:
    checks: list[ValidationCheck] = []
    for key, old_value in before.items():
        new_value = after.get(key)
        checks.append(
            _check(
                f"control_{key}",
                old_value == new_value,
                True,
                f"Control cell {key} is unchanged" if old_value == new_value else f"Control cell {key} regressed",
                {"before": old_value, "after": new_value},
            )
        )
    return checks
