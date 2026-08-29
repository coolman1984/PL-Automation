"""Native Excel cloning plus narrowly scoped formula-token rewrites."""

from __future__ import annotations

import re
from .errors import FormulaCloneError, FormulaRewriteError, UnsupportedFormulaError
from .models import FormulaAudit, RunCodes


def clone_range_with_excel(source_range: object, destination_range: object) -> None:
    try:
        source_range.Copy(Destination=destination_range)
    except Exception as exc:  # pragma: no cover - requires Excel
        raise FormulaCloneError(f"Excel native range copy failed: {exc}") from exc


def is_formula_value(value: object) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _version_literal_pattern(version: str) -> re.Pattern[str]:
    if not version or '"' in version:
        raise FormulaRewriteError("Version must be a non-empty token without quote characters")
    return re.compile(re.escape(f'"{version}"'))


def formula_contains_exact_quoted_version(formula: str, version: str) -> bool:
    return bool(_version_literal_pattern(version).search(formula))


def rewrite_formula_exact_quoted_version(
    formula: str, old_version: str, new_version: str
) -> tuple[str, int]:
    if not is_formula_value(formula):
        return formula, 0
    pattern = _version_literal_pattern(old_version)
    return pattern.subn(f'"{new_version}"', formula)


def _iter_formula_cells(range_obj: object):
    """Yield ordinary formula cells without reading the whole workbook cell-by-cell."""
    try:
        formula_cells = range_obj.SpecialCells(-4123)  # xlCellTypeFormulas
    except Exception:
        return
    for area_index in range(1, int(formula_cells.Areas.Count) + 1):
        area = formula_cells.Areas(area_index)
        for row_offset in range(int(area.Rows.Count)):
            for col_offset in range(int(area.Columns.Count)):
                yield area.Cells(row_offset + 1, col_offset + 1)


def _formula_text(cell: object) -> str:
    try:
        return str(cell.Formula)
    except Exception:
        return str(cell.Formula2)


def rewrite_exact_version_criteria(
    range_obj: object, old_version: str, new_version: str
) -> int:
    total = 0
    cells = _iter_formula_cells(range_obj)
    if cells is None:
        return 0
    for cell in cells:
        formula = _formula_text(cell)
        if not is_formula_value(formula):
            continue
        # Individual assignment is deliberately limited to the new two-column range.
        # Special formulas are detected before assignment and left to native Excel copy.
        try:
            if bool(getattr(cell, "HasArray", False)) or bool(getattr(cell, "HasSpill", False)):
                if formula_contains_exact_quoted_version(formula, old_version):
                    raise UnsupportedFormulaError(
                        f"Special formula at {cell.Address} contains a literal {old_version}"
                    )
                continue
        except UnsupportedFormulaError:
            raise
        except Exception:
            pass
        rewritten, count = rewrite_formula_exact_quoted_version(formula, old_version, new_version)
        if count:
            try:
                cell.Formula = rewritten
            except Exception as exc:  # pragma: no cover - requires Excel
                raise FormulaRewriteError(
                    f"Could not rewrite formula at {cell.Address}: {exc}"
                ) from exc
            total += count
    return total


def classify_formula(formula: str, context: dict[str, object] | None = None) -> str:
    if not is_formula_value(formula):
        return "no formula"
    upper = formula.upper()
    if "INDIRECT(" in upper or "OFFSET(" in upper:
        return "unsupported/special formula"
    if '"T08"' in formula or '"A08"' in formula:
        return "literal-version formula"
    if "!" in formula:
        return "cross-sheet formula"
    if "IFERROR(" in upper or "/" in formula:
        return "percentage formula"
    if any(fn in upper for fn in ("SUM(", "SUBTOTAL(", "AGGREGATE(")):
        return "subtotal/parent sum formula"
    if context and context.get("header_driven"):
        return "header-driven formula"
    return "formula"


def count_formula_cells(range_obj: object) -> int:
    return sum(1 for _ in (_iter_formula_cells(range_obj) or ()))


def _count_formula_literals(range_obj: object, version: str) -> int:
    count = 0
    for cell in (_iter_formula_cells(range_obj) or ()):
        if formula_contains_exact_quoted_version(_formula_text(cell), version):
            count += 1
    return count


def audit_formula_pair(
    source_amount: object,
    source_pct: object,
    actual_amount: object,
    actual_pct: object,
    codes: RunCodes,
) -> FormulaAudit:
    actual_target_count = _count_formula_literals(actual_amount, codes.target_version)
    actual_actual_count = _count_formula_literals(actual_amount, codes.actual_version)
    warnings: list[str] = []
    if actual_target_count:
        warnings.append(
            f"{actual_target_count} A08 amount formulas still contain quoted {codes.target_version}"
        )
    return FormulaAudit(
        source_amount_formula_count=count_formula_cells(source_amount),
        actual_amount_formula_count=count_formula_cells(actual_amount),
        source_pct_formula_count=count_formula_cells(source_pct),
        actual_pct_formula_count=count_formula_cells(actual_pct),
        actual_quoted_target_count=actual_target_count,
        actual_quoted_actual_count=actual_actual_count,
        warnings=warnings,
    )
