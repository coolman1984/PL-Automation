"""Insert and validate the A08 pair in the three business-total sheets."""

from __future__ import annotations

from .block_locator import ensure_actual_absent
from .constants import XL_FORMAT_FROM_LEFT_OR_ABOVE, XL_TO_RIGHT, target_version_for_month
from .errors import FormulaCloneError, ValidationError
from .formula_clone import audit_formula_pair, clone_range_with_excel, rewrite_exact_version_criteria
from .header_discovery import read_header_snapshot, value_at
from .merge_formatting import (
    capture_column_properties,
    ensure_august_merge_extended,
    restore_column_properties,
)
from .models import MonthBlock, RunCodes, SheetUpdateResult, ValidationCheck


def get_last_used_row(worksheet: object) -> int:
    used = worksheet.UsedRange
    return int(used.Row + used.Rows.Count - 1)


def insert_two_columns(worksheet: object, insert_at_col: int) -> None:
    try:
        columns = worksheet.Columns(int(insert_at_col)).Resize(ColumnSize=2)
        columns.Insert(Shift=XL_TO_RIGHT, CopyOrigin=XL_FORMAT_FROM_LEFT_OR_ABOVE)
    except Exception as exc:  # pragma: no cover - requires Excel
        raise FormulaCloneError(
            f"Could not insert two Excel columns at {insert_at_col} on {worksheet.Name}: {exc}"
        ) from exc


def _range(worksheet: object, first_row: int, last_row: int, first_col: int, last_col: int):
    return worksheet.Range(
        worksheet.Cells(first_row, first_col),
        worksheet.Cells(last_row, last_col),
    )


def set_actual_headers(
    worksheet: object,
    block: MonthBlock,
    actual_amount_col: int,
    actual_pct_col: int,
    codes: RunCodes,
) -> None:
    worksheet.Cells(block.version_header_row, actual_amount_col).Value2 = codes.actual_version
    worksheet.Cells(block.version_header_row, actual_pct_col).Value2 = "%"
    if block.period_header_row is not None:
        period = value_at(
            read_header_snapshot(worksheet), block.period_header_row, block.target_col
        )
        if str(period).strip() != codes.period:
            raise ValidationError(
                f"Period header changed unexpectedly while inserting {codes.actual_version} on {block.sheet}"
            )


def _check(name: str, passed: bool, message: str, evidence: dict[str, object] | None = None):
    return ValidationCheck(name, passed, True, message, evidence or {})


def validate_local_business_update(
    worksheet: object, result: SheetUpdateResult, codes: RunCodes
) -> list[ValidationCheck]:
    block = result.before_block
    actual_amount = worksheet.Cells(block.version_header_row, result.actual_amount_col).Value2
    actual_pct = worksheet.Cells(block.version_header_row, result.actual_pct_col).Value2
    september_col = block.september_start_col + 2
    september_header = worksheet.Cells(block.version_header_row, september_col).Value2
    expected_september = target_version_for_month(codes.month + 1)
    checks = [
        _check(
            "actual_amount_header",
            str(actual_amount).strip().upper() == codes.actual_version,
            f"A08 amount header is {actual_amount!r}",
        ),
        _check(
            "actual_percent_header",
            str(actual_pct).strip() == "%",
            f"A08 percentage header is {actual_pct!r}",
        ),
        _check(
            "september_shifted",
            str(september_header).strip().upper() == expected_september,
            f"September block still begins at column {september_col}",
            {
                "value": september_header,
                "expected": expected_september,
                "column": september_col,
            },
        ),
        _check(
            "formula_topology_amount",
            result.formula_audit.actual_amount_formula_count
            == result.formula_audit.source_amount_formula_count,
            "A08 amount formula count matches T08",
            {
                "source": result.formula_audit.source_amount_formula_count,
                "actual": result.formula_audit.actual_amount_formula_count,
            },
        ),
        _check(
            "formula_topology_percent",
            result.formula_audit.actual_pct_formula_count
            == result.formula_audit.source_pct_formula_count,
            "A08 percentage formula count matches T08",
            {
                "source": result.formula_audit.source_pct_formula_count,
                "actual": result.formula_audit.actual_pct_formula_count,
            },
        ),
        _check(
            "no_unexplained_t08_criteria",
            result.formula_audit.actual_quoted_target_count == 0,
            "No quoted T08 criterion remains in A08 amount formulas",
            {"count": result.formula_audit.actual_quoted_target_count},
        ),
    ]
    return checks


def update_business_sheet(
    worksheet: object, block: MonthBlock, codes: RunCodes
) -> SheetUpdateResult:
    snapshot = read_header_snapshot(worksheet)
    ensure_actual_absent(snapshot, block, codes.actual_version)
    last_used_row = block.last_used_row or get_last_used_row(worksheet)
    source_properties = capture_column_properties(
        worksheet, [block.target_col, block.target_pct_col]
    )
    source_range = _range(
        worksheet, 1, last_used_row, block.target_col, block.target_pct_col
    )
    insert_two_columns(worksheet, block.insert_at_col)
    actual_amount_col = block.insert_at_col
    actual_pct_col = block.insert_at_col + 1
    destination_range = _range(
        worksheet, 1, last_used_row, actual_amount_col, actual_pct_col
    )
    clone_range_with_excel(source_range, destination_range)
    restore_column_properties(
        worksheet, source_properties, [actual_amount_col, actual_pct_col]
    )
    set_actual_headers(worksheet, block, actual_amount_col, actual_pct_col, codes)
    merge_repaired = ensure_august_merge_extended(
        worksheet, block, actual_pct_col
    )
    rewrite_exact_version_criteria(
        destination_range, codes.target_version, codes.actual_version
    )
    try:
        worksheet.Application.CutCopyMode = False
    except Exception:
        pass

    source_amount = _range(worksheet, 1, last_used_row, block.target_col, block.target_col)
    source_pct = _range(worksheet, 1, last_used_row, block.target_pct_col, block.target_pct_col)
    actual_amount = _range(worksheet, 1, last_used_row, actual_amount_col, actual_amount_col)
    actual_pct = _range(worksheet, 1, last_used_row, actual_pct_col, actual_pct_col)
    formula_audit = audit_formula_pair(
        source_amount, source_pct, actual_amount, actual_pct, codes
    )
    provisional = SheetUpdateResult(
        sheet=str(worksheet.Name),
        before_block=block,
        actual_amount_col=actual_amount_col,
        actual_pct_col=actual_pct_col,
        formula_audit=formula_audit,
        merge_repaired=merge_repaired,
        locally_valid=False,
        warnings=list(formula_audit.warnings),
    )
    checks = validate_local_business_update(worksheet, provisional, codes)
    failed = [check for check in checks if check.required and not check.passed]
    provisional.locally_valid = not failed
    if failed:
        raise ValidationError(
            f"Local validation failed for {worksheet.Name}: "
            + "; ".join(check.message for check in failed),
            evidence=[check.evidence for check in failed],
        )
    return provisional
