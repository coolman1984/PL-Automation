"""Plain data contracts shared by discovery, update, validation, and reporting."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class RunCodes:
    year: int
    month: int
    month_name: str
    period: str
    target_version: str
    forecast_version: str
    actual_version: str


@dataclass(frozen=True)
class SafetyConfig:
    update_external_links: bool = False
    refresh_pivots: bool = False
    overwrite_existing_actual: bool = False
    keep_failed_workbook: bool = True
    reopen_after_save: bool = True


@dataclass(frozen=True)
class ValidationConfig:
    numeric_tolerance: float = 0.01
    require_all_sheets: bool = True
    require_unique_header_match: bool = True
    calculation_timeout_seconds: int = 1800


@dataclass(frozen=True)
class AppConfig:
    target_sheets: tuple[str, ...]
    total_sheet: str
    codes: RunCodes
    safety: SafetyConfig = SafetyConfig()
    validation: ValidationConfig = ValidationConfig()


@dataclass
class ExcelApplicationState:
    calculation: int
    screen_updating: bool
    enable_events: bool
    display_alerts: bool
    ask_to_update_links: bool
    visible: bool


@dataclass
class MergedArea:
    first_row: int
    first_column: int
    row_count: int
    column_count: int
    top_left_value: Any = None

    @property
    def last_row(self) -> int:
        return self.first_row + self.row_count - 1

    @property
    def last_column(self) -> int:
        return self.first_column + self.column_count - 1


@dataclass(frozen=True)
class HeaderSnapshot:
    sheet: str
    first_column: int
    last_column: int
    first_row: int
    last_row: int
    values: tuple[tuple[Any, ...], ...]
    merged_areas: tuple[MergedArea, ...] = ()


@dataclass(frozen=True)
class MonthBlock:
    sheet: str
    year: int
    month: int
    period: str
    target_col: int
    target_pct_col: int
    forecast_col: int
    forecast_pct_col: int
    insert_at_col: int
    version_header_row: int
    period_header_row: int | None
    month_header_row: int | None
    month_merge: MergedArea | None
    september_start_col: int
    last_used_row: int
    evidence: tuple[str, ...] = ()

    @property
    def start_col(self) -> int:
        return self.month_merge.first_column if self.month_merge else self.target_col

    @property
    def end_col(self) -> int:
        return self.september_start_col - 1


@dataclass
class ColumnProperties:
    width: float
    hidden: bool
    outline_level: int | None = None


@dataclass
class FormulaAudit:
    source_amount_formula_count: int
    actual_amount_formula_count: int
    source_pct_formula_count: int
    actual_pct_formula_count: int
    actual_quoted_target_count: int
    actual_quoted_actual_count: int
    special_formula_count: int = 0
    warnings: list[str] = field(default_factory=list)


@dataclass
class SheetUpdateResult:
    sheet: str
    before_block: MonthBlock
    actual_amount_col: int
    actual_pct_col: int
    formula_audit: FormulaAudit
    merge_repaired: bool
    locally_valid: bool
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class SourceReference:
    sheet: str
    column: int
    row: int
    absolute_column: bool = False
    absolute_row: bool = False


@dataclass
class TotalPLRowMapping:
    total_pl_row: int
    label: str
    lineage_source: str
    original_formula: str
    rewritten_formula: str
    source_references: list[SourceReference]
    classification: str


@dataclass
class ReconciliationResult:
    total_pl_row: int
    label: str
    actual: float | None
    expected: float | None
    difference: float | None
    source_references: list[str]
    passed: bool
    reason: str | None = None


@dataclass
class ValidationCheck:
    name: str
    passed: bool
    required: bool
    message: str
    evidence: dict[str, object] = field(default_factory=dict)


@dataclass
class ValidationReport:
    stage: str
    checks: list[ValidationCheck] = field(default_factory=list)
    reconciliations: list[ReconciliationResult] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return all(check.passed for check in self.checks if check.required) and all(
            item.passed for item in self.reconciliations
        )


@dataclass
class WorkbookFingerprint:
    path: str
    file_name: str
    size_bytes: int
    modified_utc: str
    sha256: str | None
    file_format: int | float | None
    sheet_names: list[str]
    sheet_count: int
    external_links: list[str] = field(default_factory=list)
    defined_name_count: int | None = None
    pivot_counts: dict[str, int] = field(default_factory=dict)
    connection_count: int | None = None
    has_vba_project: bool | None = None


@dataclass
class RunPaths:
    run_id: str
    run_dir: Path
    working_path: Path
    report_path: Path
    manifest_path: Path
    final_path: Path


@dataclass
class RunManifest:
    run_id: str
    status: str
    phase: str
    source: str
    output: str | None
    started_utc: str
    ended_utc: str | None = None
    codes: dict[str, object] = field(default_factory=dict)
    fingerprints: dict[str, object] = field(default_factory=dict)
    discovery: dict[str, object] = field(default_factory=dict)
    updates: dict[str, object] = field(default_factory=dict)
    validations: dict[str, object] = field(default_factory=dict)
    error: dict[str, object] | None = None

    def to_dict(self) -> dict[str, object]:
        return asdict(self)

