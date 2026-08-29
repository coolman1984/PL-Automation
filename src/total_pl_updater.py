"""Conservative A1 cross-sheet reference parsing for Total PL lineage."""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence

from .errors import FormulaCloneError, TotalPLMappingError, UnsupportedFormulaError, ValidationError
from .formula_clone import audit_formula_pair, clone_range_with_excel
from .merge_formatting import (
    capture_column_properties,
    ensure_august_merge_extended,
    restore_column_properties,
)
from .models import (
    MonthBlock,
    RunCodes,
    SheetUpdateResult,
    SourceReference,
    TotalPLRowMapping,
)


_SHEET_REF_RE = re.compile(
    r"(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_ ]*)!"
    r"(?P<colabs>\$?)(?P<col>[A-Za-z]{1,3})(?P<rowabs>\$?)(?P<row>[0-9]+)"
)
_EXTERNAL_REF_RE = re.compile(r"\[[^\]]+\]")


def column_number_to_letters(column: int) -> str:
    if not isinstance(column, int) or column < 1:
        raise ValueError(f"Excel column number must be positive: {column!r}")
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_letters_to_number(letters: str) -> int:
    if not re.fullmatch(r"[A-Za-z]{1,3}", letters):
        raise ValueError(f"Invalid Excel column letters: {letters!r}")
    result = 0
    for char in letters.upper():
        result = result * 26 + ord(char) - 64
    return result


def _unquote_sheet(sheet: str) -> str:
    if sheet.startswith("'") and sheet.endswith("'"):
        return sheet[1:-1].replace("''", "'")
    return sheet


def extract_cross_sheet_a1_references(
    formula: str, allowed_sheets: Sequence[str]
) -> list[SourceReference]:
    if not isinstance(formula, str) or not formula.startswith("="):
        return []
    if _EXTERNAL_REF_RE.search(formula):
        raise UnsupportedFormulaError("External workbook reference cannot be mapped safely")
    upper = formula.upper()
    if "INDIRECT(" in upper or "OFFSET(" in upper:
        raise UnsupportedFormulaError("Dynamic address formula cannot be mapped safely")
    allowed = {name.casefold(): name for name in allowed_sheets}
    refs: list[SourceReference] = []
    for match in _SHEET_REF_RE.finditer(formula):
        sheet = _unquote_sheet(match.group("sheet"))
        canonical = allowed.get(sheet.casefold())
        if canonical is None:
            continue
        refs.append(
            SourceReference(
                sheet=canonical,
                column=column_letters_to_number(match.group("col")),
                row=int(match.group("row")),
                absolute_column=bool(match.group("colabs")),
                absolute_row=bool(match.group("rowabs")),
            )
        )
    return refs


def rewrite_business_source_columns(
    formula: str,
    old_source_columns: Mapping[str, set[int]],
    new_a08_columns: Mapping[str, int],
) -> tuple[str, list[SourceReference]]:
    """Rewrite only references to known old business T08 columns.

    Row numbers and absolute/relative markers are preserved. Any allowed-sheet
    reference to an unexpected source column is left unchanged so the caller can
    decide whether the lineage is provable.
    """
    refs = extract_cross_sheet_a1_references(formula, tuple(old_source_columns))
    replacements: dict[tuple[int, int], str] = {}
    for ref in refs:
        if ref.column in old_source_columns.get(ref.sheet, set()):
            replacements[(ref.column, ref.row)] = column_number_to_letters(
                new_a08_columns[ref.sheet]
            )

    def replace(match: re.Match[str]) -> str:
        sheet = _unquote_sheet(match.group("sheet"))
        canonical = next(
            (name for name in old_source_columns if name.casefold() == sheet.casefold()),
            None,
        )
        old_col = column_letters_to_number(match.group("col"))
        row = int(match.group("row"))
        if canonical is None or old_col not in old_source_columns.get(canonical, set()):
            return match.group(0)
        prefix_sheet = match.group("sheet")
        return (
            f"{prefix_sheet}!{match.group('colabs')}"
            f"{column_number_to_letters(new_a08_columns[canonical])}"
            f"{match.group('rowabs')}{row}"
        )

    rewritten = _SHEET_REF_RE.sub(replace, formula)
    new_refs = extract_cross_sheet_a1_references(
        rewritten, tuple(old_source_columns)
    )
    return rewritten, new_refs


def _column_values(range_obj: object) -> list[object]:
    raw = range_obj.Formula
    if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple):
        return [row[0] for row in raw]
    if isinstance(raw, tuple):
        return list(raw)
    return [raw]


def _row_label(worksheet: object, row: int) -> str:
    for column in (5, 6, 2, 1):
        try:
            value = worksheet.Cells(row, column).Value2
        except Exception:
            continue
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


def _lineage_references(
    formula: str,
    business_results: Mapping[str, SheetUpdateResult],
) -> list[SourceReference]:
    allowed = tuple(business_results)
    refs = extract_cross_sheet_a1_references(formula, allowed)
    unexpected = [
        ref
        for ref in refs
        if ref.column != business_results[ref.sheet].before_block.target_col
    ]
    if unexpected:
        raise TotalPLMappingError(
            "Total PL formula references an unexpected source column",
            evidence={"formula": formula, "unexpected": [ref.__dict__ for ref in unexpected]},
        )
    return refs


def choose_lineage_formula(
    t08_formula: str | None,
    s08_formula: str | None,
    source_blocks: Mapping[str, MonthBlock],
) -> tuple[str, str]:
    """Choose T08 first, then S08, only if it exposes current business columns."""
    old_cols = {sheet: {block.target_col} for sheet, block in source_blocks.items()}
    for source_name, formula in (("T08", t08_formula), ("S08", s08_formula)):
        if not isinstance(formula, str) or not formula.startswith("="):
            continue
        refs = extract_cross_sheet_a1_references(formula, tuple(source_blocks))
        if refs and all(ref.column in old_cols[ref.sheet] for ref in refs):
            return source_name, formula
    raise TotalPLMappingError(
        "No analogous Total PL T08/S08 formula exposed a provable business-sheet mapping",
        evidence={"t08_formula": t08_formula, "s08_formula": s08_formula},
    )


def candidate_has_business_lineage(
    worksheet: object,
    block: MonthBlock,
    source_blocks: Mapping[str, MonthBlock],
) -> bool:
    """Disambiguate Total PL's historical blocks using target-sheet columns."""
    expected = {sheet: block_result.target_col for sheet, block_result in source_blocks.items()}
    try:
        raw = worksheet.Range(
            worksheet.Cells(1, block.target_col),
            worksheet.Cells(block.last_used_row, block.target_col),
        ).Formula
    except Exception:
        return False
    values = raw if isinstance(raw, tuple) else (raw,)
    if values and isinstance(values[0], tuple):
        values = [row[0] for row in values]
    for formula in values:
        if not isinstance(formula, str) or not formula.startswith("="):
            continue
        try:
            refs = extract_cross_sheet_a1_references(formula, tuple(expected))
        except UnsupportedFormulaError:
            continue
        if refs and set(ref.sheet for ref in refs) == set(expected) and all(
            ref.column == expected[ref.sheet] for ref in refs
        ):
            return True
    return False


def classify_total_pl_row(
    value: object, formula: object, pct_formula: object
) -> str:
    if isinstance(formula, str) and formula.startswith("="):
        if "INDIRECT(" in formula.upper() or "OFFSET(" in formula.upper():
            return "unsupported/special formula"
        return "formula"
    if value is None or str(value).strip() == "":
        return "blank/static row"
    if isinstance(value, str) and value.startswith("="):
        return "formula"
    if pct_formula not in (None, "") and isinstance(pct_formula, str) and pct_formula.startswith("="):
        return "formatting-only/static amount"
    return "text/static label"


def prove_total_pl_row_mapping(
    total_sheet: object,
    row: int,
    total_block: MonthBlock,
    business_results: Mapping[str, SheetUpdateResult],
) -> TotalPLRowMapping:
    t08_formula = total_sheet.Cells(row, total_block.target_col).Formula
    s08_formula = total_sheet.Cells(row, total_block.forecast_col).Formula
    value = total_sheet.Cells(row, total_block.target_col).Value2
    pct_formula = total_sheet.Cells(row, total_block.target_pct_col).Formula
    classification = classify_total_pl_row(value, t08_formula, pct_formula)
    label = _row_label(total_sheet, row)
    if classification != "formula":
        return TotalPLRowMapping(
            total_pl_row=row,
            label=label,
            lineage_source="none",
            original_formula=str(t08_formula or ""),
            rewritten_formula=str(t08_formula or ""),
            source_references=[],
            classification=classification,
        )
    try:
        lineage_source, formula = choose_lineage_formula(
            str(t08_formula), str(s08_formula),
            {name: result.before_block for name, result in business_results.items()},
        )
    except TotalPLMappingError:
        # Same-sheet formulas (subtotals/ratios) can be safely retained from
        # native copy; only cross-sheet business totals need explicit mapping.
        refs = extract_cross_sheet_a1_references(
            str(t08_formula), tuple(business_results)
        )
        if not refs:
            return TotalPLRowMapping(
                total_pl_row=row,
                label=label,
                lineage_source="native-copy",
                original_formula=str(t08_formula),
                rewritten_formula=str(t08_formula),
                source_references=[],
                classification="formula-no-cross-sheet",
            )
        raise
    old_cols = {name: {result.before_block.target_col} for name, result in business_results.items()}
    new_cols = {name: result.actual_amount_col for name, result in business_results.items()}
    refs = _lineage_references(formula, business_results)
    rewritten, new_refs = rewrite_business_source_columns(formula, old_cols, new_cols)
    if not refs or not new_refs:
        raise TotalPLMappingError(
            "Total PL formula did not contain a rewriteable business-source reference",
            evidence={"row": row, "formula": formula},
        )
    return TotalPLRowMapping(
        total_pl_row=row,
        label=label,
        lineage_source=lineage_source,
        original_formula=formula,
        rewritten_formula=rewritten,
        source_references=new_refs,
        classification="cross-sheet formula",
    )


def update_total_pl(
    total_sheet: object,
    total_block: MonthBlock,
    business_results: Mapping[str, SheetUpdateResult],
    codes: RunCodes,
) -> tuple[SheetUpdateResult, list[TotalPLRowMapping]]:
    if not all(result.locally_valid for result in business_results.values()):
        raise ValidationError("Total PL cannot be updated before all business sheets validate")
    last_used_row = total_block.last_used_row
    mappings = [
        prove_total_pl_row_mapping(total_sheet, row, total_block, business_results)
        for row in range(1, last_used_row + 1)
    ]
    source_properties = capture_column_properties(
        total_sheet, [total_block.target_col, total_block.target_pct_col]
    )
    source_range = total_sheet.Range(
        total_sheet.Cells(1, total_block.target_col),
        total_sheet.Cells(last_used_row, total_block.target_pct_col),
    )
    insert_at = total_block.insert_at_col
    try:
        total_sheet.Columns(int(insert_at)).Resize(ColumnSize=2).Insert(
            Shift=-4161, CopyOrigin=0
        )
    except Exception as exc:  # pragma: no cover - requires Excel
        raise FormulaCloneError(f"Could not insert Total PL A08 columns: {exc}") from exc
    actual_amount_col = insert_at
    actual_pct_col = insert_at + 1
    destination_range = total_sheet.Range(
        total_sheet.Cells(1, actual_amount_col),
        total_sheet.Cells(last_used_row, actual_pct_col),
    )
    clone_range_with_excel(source_range, destination_range)
    restore_column_properties(
        total_sheet, source_properties, [actual_amount_col, actual_pct_col]
    )
    total_sheet.Cells(total_block.version_header_row, actual_amount_col).Value2 = codes.actual_version
    total_sheet.Cells(total_block.version_header_row, actual_pct_col).Value2 = "%"
    merge_repaired = ensure_august_merge_extended(
        total_sheet, total_block, actual_pct_col
    )

    for mapping in mappings:
        if mapping.classification == "cross-sheet formula":
            total_sheet.Cells(mapping.total_pl_row, actual_amount_col).Formula = mapping.rewritten_formula
    try:
        total_sheet.Application.CutCopyMode = False
    except Exception:
        pass

    source_amount = total_sheet.Range(
        total_sheet.Cells(1, total_block.target_col),
        total_sheet.Cells(last_used_row, total_block.target_col),
    )
    source_pct = total_sheet.Range(
        total_sheet.Cells(1, total_block.target_pct_col),
        total_sheet.Cells(last_used_row, total_block.target_pct_col),
    )
    actual_amount = total_sheet.Range(
        total_sheet.Cells(1, actual_amount_col),
        total_sheet.Cells(last_used_row, actual_amount_col),
    )
    actual_pct = total_sheet.Range(
        total_sheet.Cells(1, actual_pct_col),
        total_sheet.Cells(last_used_row, actual_pct_col),
    )
    formula_audit = audit_formula_pair(
        source_amount, source_pct, actual_amount, actual_pct, codes
    )
    result = SheetUpdateResult(
        sheet=str(total_sheet.Name),
        before_block=total_block,
        actual_amount_col=actual_amount_col,
        actual_pct_col=actual_pct_col,
        formula_audit=formula_audit,
        merge_repaired=merge_repaired,
        locally_valid=True,
        warnings=list(formula_audit.warnings),
    )
    return result, mappings
