# P&L A08 Automation — Detailed Agent Execution Plan V1

## 1. Purpose of this document

This document is the implementation handoff for the coding agent. It converts `PL_A08_AUTOMATION_CODING_PLAN_V1.md` into small, ordered, verifiable coding tasks with explicit module boundaries, function contracts, algorithms, tests, checkpoints, and stop conditions.

The original plan remains the business and safety source of truth. If this document and the original plan appear to conflict, follow the safer interpretation and stop for review. Never weaken a safety rule to make implementation easier.

The first production scope is only:

```text
Year:             2026
Month:            August (8)
Target version:   T08
Forecast version: S08
Actual version:   A08
Period:           2026.008
Sheets:           VD Total, MX Total, DA Total, Total PL
File type:        .xlsb
Writer:           Microsoft Excel desktop through COM only
```

## 2. Instructions to the coding agent

Follow these rules during every coding session:

1. Read this file and the original plan before changing code.
2. Work on one numbered task at a time.
3. Do not begin a task until its dependencies are complete.
4. Keep every change small enough to test immediately.
5. Run the listed verification before marking a task complete.
6. Record completed tasks and test results in a short development log.
7. Never test destructive behavior on the original workbook.
8. Never save the workbook through a non-Excel library.
9. Never guess workbook coordinates, formulas, row mappings, or version codes.
10. If a required workbook fact cannot be proven, raise the specified error and stop safely.
11. Do not claim production success until the final `.xlsb` closes, reopens in Excel, passes post-reopen validation, and the source hash is unchanged.

The agent may use pure-Python unit tests without Excel. Workbook integration tests require Windows, Microsoft Excel desktop, `pywin32`, and a disposable copy created through the approved transaction flow.

## 3. Absolute safety invariants

These invariants must be enforced in code, not left as documentation only.

### 3.1 Workbook writing

- Only Microsoft Excel COM may alter or save workbook content.
- `openpyxl`, `pandas`, `pyxlsb`, `xlsxwriter`, LibreOffice, ZIP/XML edits, and binary edits are forbidden for workbook writes.
- Python file-copy operations may publish an already closed and validated workbook byte-for-byte; they may not rewrite its content.

### 3.2 Source immutability

The only permitted transaction is:

```text
SOURCE.xlsb
  -> Excel Workbook.SaveCopyAs(WORKING.xlsb)
  -> edit WORKING.xlsb through Excel COM
  -> validate, save, close
  -> reopen WORKING.xlsb through Excel COM
  -> validate again, close
  -> byte-copy/move to FINAL.xlsb
```

Capture the source SHA-256 before the run and after the run. A changed source hash is a critical failure.

### 3.3 Fail-closed behavior

No final file may be published when any required discovery, calculation, mapping, or validation result is ambiguous, unavailable, timed out, or failed.

On failure:

- stop further edits;
- close the working workbook without additional saves unless preserving the already-saved failed state is explicitly required;
- never save the source;
- keep or move the working file under `failed_runs/<run-id>/` when configured;
- write the error and validation evidence to the report and manifest;
- return a non-zero process exit code.

### 3.4 Prohibited shortcuts

Do not:

- hard-code Excel column letters;
- globally replace `T08` with `A08`;
- infer Total PL row mapping from matching visible labels alone;
- refresh pivots, connections, queries, or external links;
- dismiss unexpected Excel dialogs automatically;
- implement replacement of an existing A08 pair in V1;
- invent Actual codes for October through December;
- rewrite formulas in Python when native Excel copying can safely preserve them.

## 4. Required project layout

Create this structure under `pl_actual_automation/`:

```text
pl_actual_automation/
├─ app.py
├─ requirements.txt
├─ README.md
├─ RUN_A08.bat
├─ config.yaml
├─ src/
│  ├─ __init__.py
│  ├─ constants.py
│  ├─ errors.py
│  ├─ models.py
│  ├─ config.py
│  ├─ excel_session.py
│  ├─ file_transaction.py
│  ├─ workbook_audit.py
│  ├─ header_discovery.py
│  ├─ block_locator.py
│  ├─ formula_clone.py
│  ├─ merge_formatting.py
│  ├─ business_sheet_updater.py
│  ├─ total_pl_updater.py
│  ├─ calculation.py
│  ├─ validation.py
│  ├─ reporting.py
│  └─ workflow.py
├─ tests/
│  ├─ unit/
│  │  ├─ test_constants.py
│  │  ├─ test_config.py
│  │  ├─ test_header_discovery.py
│  │  ├─ test_block_locator.py
│  │  ├─ test_formula_rewrite.py
│  │  ├─ test_total_pl_mapping.py
│  │  ├─ test_idempotency.py
│  │  └─ test_validation_logic.py
│  ├─ integration/
│  │  ├─ conftest.py
│  │  ├─ test_excel_session.py
│  │  ├─ test_save_copy_transaction.py
│  │  ├─ test_business_sheet_update.py
│  │  ├─ test_total_pl_update.py
│  │  └─ test_full_workflow.py
│  └─ fixtures/
├─ logs/
├─ output/
├─ work/
├─ backups/
└─ failed_runs/
```

Do not add a web application, database, server, cloud dependency, or GUI framework.

## 5. Dependency graph and implementation order

```text
constants + errors
        |
        v
models + configuration
        |
        +----------------------+
        v                      v
Excel session             pure discovery/formula helpers
        |                      |
        v                      v
file transaction          workbook audit
        |                      |
        +----------+-----------+
                   v
          business-sheet updater
                   |
                   v
            Total PL updater
                   |
                   v
        calculation + validation
                   |
                   v
          reporting + workflow
                   |
                   v
            CLI + batch launcher
                   |
                   v
       real-workbook staged tests
```

Shared data contracts must be completed before COM-heavy modules. Business sheets must be updated and locally validated in the order `VD Total`, `MX Total`, `DA Total`. `Total PL` must be updated only after all three have valid A08 blocks.

## 6. Data model contracts

Implement dataclasses or frozen dataclasses in `src/models.py`. JSON-facing objects must have explicit serialization helpers that convert `Path`, datetime, enum, and tuple values into JSON-safe primitives.

### 6.1 `RunCodes`

```python
@dataclass(frozen=True)
class RunCodes:
    year: int
    month: int
    month_name: str
    period: str
    target_version: str
    forecast_version: str
    actual_version: str
```

For V1, `resolve_run_codes(2026, 8)` must return `2026.008`, `T08`, `S08`, and `A08`. Execution for any month other than 8 must raise a clear scope error. Pure mapping helpers may contain January–September logic for testing, but the execute workflow remains A08-only.

### 6.2 `SafetyConfig` and `ValidationConfig`

```python
@dataclass(frozen=True)
class SafetyConfig:
    update_external_links: bool
    refresh_pivots: bool
    overwrite_existing_actual: bool
    keep_failed_workbook: bool
    reopen_after_save: bool

@dataclass(frozen=True)
class ValidationConfig:
    numeric_tolerance: float
    require_all_sheets: bool
    require_unique_header_match: bool
    calculation_timeout_seconds: int
```

Reject unsafe V1 configuration values:

- `update_external_links: true`
- `refresh_pivots: true`
- `overwrite_existing_actual: true`
- `reopen_after_save: false`
- negative numeric tolerance
- non-positive timeout

### 6.3 `AppConfig`

```python
@dataclass(frozen=True)
class AppConfig:
    target_sheets: tuple[str, ...]
    total_sheet: str
    codes: RunCodes
    safety: SafetyConfig
    validation: ValidationConfig
```

The exact target sheet tuple must be `("VD Total", "MX Total", "DA Total")` in V1 unless the business owner approves a new scope.

### 6.4 `ExcelApplicationState`

```python
@dataclass
class ExcelApplicationState:
    calculation: int
    screen_updating: bool
    enable_events: bool
    display_alerts: bool
    ask_to_update_links: bool
    visible: bool
```

### 6.5 `WorkbookFingerprint`

Include at least:

```python
@dataclass
class WorkbookFingerprint:
    path: str
    file_name: str
    size_bytes: int
    modified_utc: str
    sha256: str | None
    file_format: int
    sheet_names: list[str]
    sheet_count: int
    external_links: list[str]
    defined_name_count: int
    pivot_counts: dict[str, int]
    connection_count: int | None
    has_vba_project: bool | None
```

Unavailable optional audit properties must be recorded as `null` with a warning; they must not crash the audit. Required facts—path, extension, sheets, and file format—must be available.

### 6.6 Header and block models

```python
@dataclass(frozen=True)
class HeaderSnapshot:
    sheet: str
    first_column: int
    last_column: int
    first_row: int
    last_row: int
    values: tuple[tuple[object, ...], ...]
    merged_areas: tuple["MergedArea", ...]

@dataclass(frozen=True)
class MergedArea:
    first_row: int
    first_column: int
    row_count: int
    column_count: int
    top_left_value: object

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
    period_header_row: int
    month_header_row: int | None
    month_merge: MergedArea | None
    september_start_col: int
    last_used_row: int
    evidence: tuple[str, ...]
```

All Excel row and column indices in COM-facing models are one-based.

### 6.7 Formula and update results

```python
@dataclass
class FormulaAudit:
    source_amount_formula_count: int
    actual_amount_formula_count: int
    source_pct_formula_count: int
    actual_pct_formula_count: int
    actual_quoted_target_count: int
    actual_quoted_actual_count: int
    special_formula_count: int
    warnings: list[str]

@dataclass
class SheetUpdateResult:
    sheet: str
    before_block: MonthBlock
    actual_amount_col: int
    actual_pct_col: int
    formula_audit: FormulaAudit
    merge_repaired: bool
    locally_valid: bool
    warnings: list[str]
```

### 6.8 Total PL mapping and reconciliation

```python
@dataclass(frozen=True)
class SourceReference:
    sheet: str
    column: int
    row: int
    absolute_column: bool
    absolute_row: bool

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
    reason: str | None
```

### 6.9 Validation and run state

```python
@dataclass
class ValidationCheck:
    name: str
    passed: bool
    required: bool
    message: str
    evidence: dict[str, object]

@dataclass
class ValidationReport:
    stage: str
    checks: list[ValidationCheck]
    reconciliations: list[ReconciliationResult]
    warnings: list[str]

    @property
    def passed(self) -> bool: ...

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
    ended_utc: str | None
    codes: dict[str, object]
    fingerprints: dict[str, object]
    discovery: dict[str, object]
    updates: dict[str, object]
    validations: dict[str, object]
    error: dict[str, object] | None
```

Allowed statuses: `STARTED`, `PREFLIGHT_FAILED`, `EXECUTION_FAILED`, `VALIDATION_FAILED`, `SUCCESS`.

## 7. Error taxonomy

Create these exception classes in `src/errors.py`, all inheriting from `PLAutomationError`:

```text
ConfigurationError
UnsupportedScopeError
ExcelNotInstalledError
ExcelConnectionError
WorkbookNotFoundError
WorkbookIdentityError
WorkbookReadOnlyError
DRMOpenBlockedError
MissingSheetError
WorkbookFormatError
AmbiguousMonthBlockError
MissingMonthBlockError
ExistingActualColumnError
CopyCreationError
FormulaCloneError
FormulaRewriteError
UnsupportedFormulaError
MergeRepairError
TotalPLMappingError
CalculationTimeoutError
ValidationError
SaveError
ReopenValidationError
PublicationError
SourceChangedError
```

Each error must carry:

- a stable machine-readable code;
- a plain user-facing message;
- optional structured evidence;
- the phase in which it occurred.

Do not expose raw COM tracebacks as the only user message. Preserve the traceback in the diagnostic log.

## 8. Module and function specification

### 8.1 `src/constants.py`

Implement:

```python
TARGET_SHEETS = ("VD Total", "MX Total", "DA Total")
TOTAL_SHEET = "Total PL"
HEADER_FIRST_ROW = 1
HEADER_LAST_ROW = 30
XLSB_EXTENSION = ".xlsb"

def format_period(year: int, month: int) -> str: ...
def target_version_for_month(month: int) -> str: ...
def forecast_version_for_month(month: int) -> str: ...
def actual_version_for_month(month: int) -> str: ...
def resolve_run_codes(year: int, month: int, *, execution: bool) -> RunCodes: ...
```

Rules:

- `format_period(2026, 8) == "2026.008"`.
- January–September target/forecast/actual helpers return two decimal digits.
- Target October–December may map to `T0A`, `T0B`, `T0C` only in the pure target helper.
- Actual October–December must raise `UnsupportedScopeError`.
- `execution=True` rejects every month except August in V1.

Keep required Excel numeric constants in one clearly named section if generated COM constants are unavailable. Add a comment naming the Excel constant represented by each number.

### 8.2 `src/config.py`

Implement:

```python
def load_yaml(path: Path) -> dict[str, object]: ...
def build_config(raw: Mapping[str, object], *, year: int, month: int,
                 execution: bool) -> AppConfig: ...
def validate_config(config: AppConfig) -> None: ...
def load_config(path: Path, *, year: int, month: int,
                execution: bool) -> AppConfig: ...
```

Use `yaml.safe_load`. Reject unknown top-level keys to catch misspellings. CLI year/month override the YAML run section. Print the resolved codes before workbook edits.

### 8.3 `src/excel_session.py`

Responsibilities:

- initialize/uninitialize COM on the current thread;
- attach to the exact source workbook already open in Excel, or create an isolated Excel instance;
- distinguish ownership of the Excel application;
- capture and restore application state;
- never quit an Excel application owned by the user.

Implement an `ExcelSession` context manager and these helpers:

```python
def normalize_windows_path(path: Path | str) -> str: ...
def iter_open_workbooks(app: object) -> Iterator[object]: ...
def find_open_workbook_by_full_path(app: object, path: Path) -> object | None: ...
def capture_application_state(app: object) -> ExcelApplicationState: ...
def apply_editing_state(app: object) -> None: ...
def restore_application_state(app: object, state: ExcelApplicationState) -> None: ...

class ExcelSession:
    @classmethod
    def attach(cls, source_path: Path) -> "ExcelSession": ...

    @classmethod
    def create(cls, *, visible: bool) -> "ExcelSession": ...

    def open_workbook(self, path: Path, *, read_only: bool,
                      update_links: bool = False) -> object: ...
    def close_workbook(self, workbook: object, *, save_changes: bool) -> None: ...
    def close(self) -> None: ...
```

Path matching must use normalized absolute paths, case-insensitive comparison, and must not select a workbook by filename alone. If more than one Excel process exists and the exact workbook cannot be proven, stop with `WorkbookIdentityError`.

For `Workbooks.Open`, explicitly disable external-link updates. Do not suppress prompts before the source is successfully authorized/opened. In Mode B, make Excel visible when an authorization prompt may require the user.

State restoration belongs in `finally`. If state restoration partially fails, log each failed property; never quit the user's Excel instance.

### 8.4 `src/file_transaction.py`

Implement:

```python
def make_run_id(actual_version: str, now: datetime | None = None) -> str: ...
def create_run_paths(project_root: Path, source_path: Path,
                     actual_version: str) -> RunPaths: ...
def assert_source_candidate(path: Path) -> None: ...
def save_working_copy(source_workbook: object, working_path: Path) -> None: ...
def assert_source_unchanged(before: WorkbookFingerprint,
                            source_path: Path) -> None: ...
def publish_validated_workbook(closed_working_path: Path,
                               final_path: Path) -> str: ...
def retain_failed_workbook(paths: RunPaths) -> Path | None: ...
```

Requirements:

- Create a unique run directory before `SaveCopyAs`.
- Working and final paths must not equal the source after canonical resolution.
- Do not overwrite an existing final file; add a collision-safe suffix.
- Call the source workbook's native `SaveCopyAs` method.
- Verify that the working path exists, has non-zero size, and retains `.xlsb`.
- Publish only a closed workbook whose final validation report passed.
- After publication, compare SHA-256 of working and final files; they must match.

### 8.5 `src/workbook_audit.py`

Implement:

```python
def sha256_file(path: Path, chunk_size: int = 4 * 1024 * 1024) -> str: ...
def get_sheet_names(workbook: object) -> list[str]: ...
def get_external_links(workbook: object) -> list[str]: ...
def get_pivot_counts(workbook: object) -> dict[str, int]: ...
def get_connection_count(workbook: object) -> int | None: ...
def detect_vba_project(workbook: object) -> bool | None: ...
def collect_fingerprint(workbook: object, path: Path) -> WorkbookFingerprint: ...
def compare_preservation(before: WorkbookFingerprint,
                         after: WorkbookFingerprint) -> list[ValidationCheck]: ...
def capture_control_cells(workbook: object,
                          configured_controls: list[dict]) -> dict[str, object]: ...
```

External-link retrieval can legitimately return no links. Normalize link paths for comparison but retain original display strings in reports. Never call link update methods.

VBA/pivot/connection APIs may fail under corporate restrictions. Treat inability to read optional counts as a warning unless the preflight can read them and post-edit cannot; that regression requires review.

### 8.6 `src/header_discovery.py`

Implement pure and COM-facing parts separately:

```python
def normalize_header_value(value: object) -> str: ...
def read_header_snapshot(worksheet: object, *, first_row: int = 1,
                         last_row: int = 30) -> HeaderSnapshot: ...
def value_at(snapshot: HeaderSnapshot, row: int, column: int) -> str: ...
def effective_merged_value(snapshot: HeaderSnapshot,
                           row: int, column: int) -> str: ...
def columns_matching(snapshot: HeaderSnapshot, exact_text: str) -> list[tuple[int, int]]: ...
```

`read_header_snapshot` must:

- determine used first/last columns without relying only on `UsedRange` when formatting inflation is extreme;
- read rows 1–30 across the selected columns with one bulk `Value2` call;
- enumerate only merged areas intersecting the header range;
- return immutable plain data so discovery tests do not need Excel.

Normalization trims leading/trailing whitespace, converts non-`None` values safely to text, and uppercases for matching. It must not collapse meaningful punctuation in `2026.008` or `%`.

### 8.7 `src/block_locator.py`

Implement:

```python
def find_month_block_candidates(snapshot: HeaderSnapshot,
                                codes: RunCodes,
                                last_used_row: int) -> list[MonthBlock]: ...
def detect_existing_actual(snapshot: HeaderSnapshot,
                           block: MonthBlock,
                           actual_version: str) -> list[tuple[int, int]]: ...
def select_unique_month_block(candidates: Sequence[MonthBlock],
                              *, require_unique: bool = True) -> MonthBlock: ...
def locate_month_block(worksheet: object, codes: RunCodes,
                       *, require_unique: bool = True) -> MonthBlock: ...
```

Candidate algorithm:

1. Find every exact `T08` in the header snapshot.
2. For each candidate, identify the version header row and require `%` in the immediately following logical column.
3. Search to the right within the same version row for exact `S08`; require its `%` immediately after it.
4. Require the candidate columns to share period `2026.008`, allowing the period to be represented by a merged header spanning the block.
5. If present, require the effective merged month label to be August.
6. Require a discoverable September boundary to the right using period `2026.009`, September month label, or September version structure.
7. Set `insert_at_col = forecast_pct_col + 1` and require it to equal the September boundary start.
8. Capture evidence for every satisfied anchor.
9. Reject candidates missing any mandatory anchor; do not pick the highest score from incomplete candidates.
10. Require exactly one valid candidate.

Idempotency search must examine the August block between its start and September boundary. One or more exact `A08` matches cause `ExistingActualColumnError` in V1.

Dry-run discovery must locate all four sheets first and return no editable COM ranges. If any sheet fails, `READY TO EXECUTE` is `NO`.

### 8.8 `src/formula_clone.py`

Implement:

```python
def clone_range_with_excel(source_range: object,
                           destination_range: object) -> None: ...
def is_formula_value(value: object) -> bool: ...
def formula_contains_exact_quoted_version(formula: str,
                                          version: str) -> bool: ...
def rewrite_formula_exact_quoted_version(formula: str,
                                         old_version: str,
                                         new_version: str) -> tuple[str, int]: ...
def rewrite_exact_version_criteria(range_obj: object,
                                   old_version: str,
                                   new_version: str) -> int: ...
def classify_formula(formula: str, context: dict[str, object]) -> str: ...
def count_formula_cells(range_obj: object) -> int: ...
def audit_formula_pair(source_amount: object, source_pct: object,
                       actual_amount: object, actual_pct: object,
                       codes: RunCodes) -> FormulaAudit: ...
```

Rules for formula rewriting:

- Operate only inside the new A08 destination range.
- Operate only on cells containing formulas.
- Replace only the exact quoted string literal `"T08"` with `"A08"`.
- Do not replace unquoted tokens, partial strings, headers, values, sheet names, or other ranges.
- Preserve formulas that already derive the criterion from the destination header.
- Detect array formulas, dynamic/spill formulas, data-table formulas, and other special formula containers before assigning cell formulas individually.
- For special formulas, prefer native-copy preservation; if a literal T08 criterion remains and cannot be changed without reconstructing a special formula, stop with `UnsupportedFormulaError` and report the cell.

Do not bulk-write a complete column's formula array if that would overwrite constants, labels, or special formula structures. Iterate only formula cells/areas in the two-column destination range.

Formula count differences are not automatically safe. Every difference must be explainable by merged header behavior or other explicitly recorded source structure.

### 8.9 `src/merge_formatting.py`

Implement:

```python
@dataclass
class ColumnProperties:
    width: float
    hidden: bool
    outline_level: int | None

def capture_column_properties(worksheet: object,
                              columns: Sequence[int]) -> dict[int, ColumnProperties]: ...
def restore_column_properties(worksheet: object,
                              source_properties: dict[int, ColumnProperties],
                              destination_columns: Sequence[int]) -> None: ...
def capture_month_merge(worksheet: object, block: MonthBlock) -> MergedArea | None: ...
def ensure_august_merge_extended(worksheet: object, block: MonthBlock,
                                 actual_pct_col: int) -> bool: ...
```

Merge repair algorithm:

1. Capture only the August month-header merge before insertion.
2. After insertion, inspect whether Excel expanded it through the new A08 percentage column.
3. If already correct, do nothing.
4. Otherwise, ensure no unrelated merge overlaps the intended repaired range.
5. Preserve the top-left value, horizontal/vertical alignment, number format, fill, font, borders, and protection.
6. Unmerge only the captured August merge.
7. Merge from the original August start column through the new A08 percentage column.
8. Restore preserved properties and verify the resulting `MergeArea` exactly.

Any overlap with an unrelated merge is ambiguous and must stop the sheet update.

### 8.10 `src/business_sheet_updater.py`

Implement:

```python
def get_last_used_row(worksheet: object) -> int: ...
def insert_two_columns(worksheet: object, insert_at_col: int) -> None: ...
def set_actual_headers(worksheet: object, block: MonthBlock,
                       actual_amount_col: int, actual_pct_col: int,
                       codes: RunCodes) -> None: ...
def validate_local_business_update(worksheet: object,
                                   result: SheetUpdateResult,
                                   codes: RunCodes) -> list[ValidationCheck]: ...
def update_business_sheet(worksheet: object, block: MonthBlock,
                          codes: RunCodes) -> SheetUpdateResult: ...
```

Exact update sequence for one business sheet:

1. Reconfirm that the sheet name and discovered block match.
2. Re-read the header and reconfirm no A08 exists immediately before editing.
3. Capture T08 pair formula counts, widths, hidden state, outline level, and August merge.
4. Insert two full worksheet columns at `block.insert_at_col` using Excel's column insertion.
5. Set `actual_amount_col = insert_at_col`; set `actual_pct_col = insert_at_col + 1`.
6. Copy rows 1 through `last_used_row` from the August T08 amount and percentage pair into the new pair using native Excel copy with destination behavior.
7. Restore destination column widths/hidden/outline properties from the corresponding T08 columns.
8. Set only the new version header cells to `A08` and `%`.
9. Verify the period remains `2026.008` and month remains August.
10. Extend or repair the August month merge if necessary.
11. Rewrite exact quoted T08 criteria only inside the new pair.
12. Clear Excel copy mode.
13. Run the formula audit and local structural validation.
14. Return `SheetUpdateResult`; if any required local check fails, raise `ValidationError` before touching the next sheet.

Do not copy the S08 pair. The business rule explicitly requires T08 equation lineage.

### 8.11 `src/total_pl_updater.py`

This is the highest-risk module. Implement conservative formula parsing and reject unsupported mappings.

Implement:

```python
def column_number_to_letters(column: int) -> str: ...
def column_letters_to_number(letters: str) -> int: ...
def extract_cross_sheet_a1_references(formula: str,
                                      allowed_sheets: Sequence[str]) -> list[SourceReference]: ...
def classify_total_pl_row(value: object, formula: object,
                          pct_formula: object) -> str: ...
def choose_lineage_formula(t08_formula: str | None,
                           s08_formula: str | None,
                           source_blocks: Mapping[str, MonthBlock]) -> tuple[str, str]: ...
def rewrite_business_source_columns(formula: str,
                                    old_source_columns: Mapping[str, set[int]],
                                    new_a08_columns: Mapping[str, int]) -> tuple[str, list[SourceReference]]: ...
def prove_total_pl_row_mapping(total_sheet: object, row: int,
                               total_block: MonthBlock,
                               business_results: Mapping[str, SheetUpdateResult]) -> TotalPLRowMapping: ...
def update_total_pl(total_sheet: object, total_block: MonthBlock,
                    business_results: Mapping[str, SheetUpdateResult],
                    codes: RunCodes) -> tuple[SheetUpdateResult, list[TotalPLRowMapping]]: ...
```

Supported conservative A1 reference forms must include quoted sheet names such as `'VD Total'!$XQ$168` and unquoted names where valid. Preserve absolute/relative row markers and exact row numbers. Reject external workbook references, `INDIRECT`, string-built addresses, or mappings that depend on a formula form the parser cannot prove.

Total PL sequence:

1. Confirm all three business update results are locally valid.
2. Confirm their A08 columns are unique and still discoverable.
3. Locate and reconfirm the Total PL August block and absence of A08.
4. Capture formatting and merge state.
5. Insert two columns after S08 `%` and before September.
6. Native-copy the structurally appropriate August pair for styles and percentage logic. Prefer T08 for consistent business-rule lineage unless workbook inspection proves S08 is structurally safer for Total PL formatting.
7. Set the A08 headers and repair the August merge.
8. For every data row, classify the analogous amount cell as formula, blank, text/static label, formatting-only, or unsupported.
9. Preserve blank/static/label behavior.
10. For formula rows, inspect both T08 and S08 analogues and choose a formula only when it exposes provable references to one or more allowed business sheets with consistent row mapping.
11. Rewrite only the referenced source column for each allowed business sheet to that sheet's newly created A08 amount column. Preserve every referenced row exactly.
12. Reject formulas with relevant references outside the three allowed sheets unless those references are demonstrated to be intentional non-business operands and unchanged.
13. Require the resulting business-total formula to reference the expected business sheets indicated by the analogous lineage. If the business rule requires all three but the lineage omits one without proof, stop.
14. Keep the native-cloned percentage formula, allowing Excel's normal relative translation to reference the new A08 amount cell.
15. Audit all amount and percentage formulas before calculation.

Do not map rows by P&L label equality. Labels may be used only in error reports and human review.

When mapping cannot be proven, raise `TotalPLMappingError` with:

- Total PL row and label;
- T08 formula;
- S08 formula;
- references parsed from both;
- specific ambiguity or unsupported construct.

### 8.12 `src/calculation.py`

Implement:

```python
def wait_for_calculation(app: object, timeout_seconds: int,
                         poll_seconds: float = 0.25) -> float: ...
def calculate_workbook_once(app: object, workbook: object,
                            timeout_seconds: int,
                            *, full_rebuild: bool = True) -> float: ...
```

Sequence:

1. Complete all four structural updates while calculation is manual.
2. Invoke one controlled calculation after the edits.
3. Because columns were inserted, use `Application.CalculateFullRebuild()` for the production validation run.
4. Poll `Application.CalculationState` until complete.
5. Raise `CalculationTimeoutError` on timeout and record elapsed time.
6. Restore the user's original calculation setting in the outer session `finally` block.

Do not repeatedly calculate after individual cell changes.

### 8.13 `src/validation.py`

Implement pure comparison helpers plus COM-facing validators:

```python
EXCEL_ERROR_TEXTS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!")

def numbers_match(actual: object, expected: object,
                  tolerance: float) -> bool: ...
def scan_formula_errors(range_obj: object) -> list[dict[str, object]]: ...
def validate_sheet_structure(worksheet: object, codes: RunCodes,
                             expected_actual_col: int) -> list[ValidationCheck]: ...
def validate_business_formulas(worksheet: object,
                               result: SheetUpdateResult,
                               codes: RunCodes) -> list[ValidationCheck]: ...
def reconcile_total_pl(total_sheet: object,
                       mappings: Sequence[TotalPLRowMapping],
                       tolerance: float) -> list[ReconciliationResult]: ...
def validate_percent_formulas(worksheet: object,
                              result: SheetUpdateResult) -> list[ValidationCheck]: ...
def validate_external_links(before: WorkbookFingerprint,
                            after: WorkbookFingerprint) -> ValidationCheck: ...
def validate_control_cells(before: Mapping[str, object],
                           after: Mapping[str, object]) -> list[ValidationCheck]: ...
def validate_workbook(workbook: object, config: AppConfig,
                      before: WorkbookFingerprint,
                      business_results: Mapping[str, SheetUpdateResult],
                      total_result: SheetUpdateResult,
                      mappings: Sequence[TotalPLRowMapping],
                      *, stage: str) -> ValidationReport: ...
```

Required structural checks:

- original sheet list preserved exactly;
- file format remains Excel binary workbook;
- all target sheets exist;
- exactly one A08 amount/percentage pair occurs in each correct August block;
- T08 and S08 remain in the August block;
- A08 is immediately before September;
- August merge includes A08 `%`;
- no unintended sheet was added.

Required business-formula checks:

- A08 formula topology/count is consistent with T08;
- no source formula became a static value in the destination;
- no unexplained exact quoted `"T08"` remains in A08 amount formulas;
- expected `"A08"` criteria exist where the source used literal `"T08"`;
- no new formula errors were introduced relative to the source pattern;
- no new external link source or changed link path exists.

Required Total PL checks:

- every mapped row has a provable source-reference list;
- read calculated values after calculation, not before;
- compute expected from the exact mapped A08 source cells;
- `abs(actual - expected) <= 0.01` by default;
- report actual, expected, difference, label, and source references;
- zero unresolved mismatches are allowed.

If a source cell is blank, text, error, or nonnumeric, apply the same semantics used by the proven formula. Do not silently coerce unexpected text to zero. Report unsupported value types.

Post-reopen validation must repeat at least:

- workbook format and target sheet readability;
- unique A08 block discovery on all four sheets;
- formula/error checks for the new pairs;
- Total PL reconciliation;
- external-link preservation;
- source hash unchanged.

### 8.14 `src/reporting.py`

Implement:

```python
def configure_logging(run_paths: RunPaths, *, verbose: bool) -> logging.Logger: ...
def write_manifest_atomic(path: Path, manifest: RunManifest) -> None: ...
def render_dry_run_report(...) -> str: ...
def render_run_report(manifest: RunManifest,
                      validation_reports: Sequence[ValidationReport]) -> str: ...
def write_run_report(path: Path, text: str) -> None: ...
def sanitize_for_logging(value: object) -> object: ...
```

Use atomic replacement for JSON/text report files, not for workbook content. Update the manifest after each major phase so a crash still leaves useful evidence.

Do not log bulk confidential cell values. It is acceptable to log the minimum reconciliation values and formula addresses needed to diagnose a mismatch.

### 8.15 `src/workflow.py`

Implement the orchestration API:

```python
def preflight(source_path: Path, config: AppConfig,
              mode: str) -> "PreflightResult": ...
def run_dry_run(source_path: Path, config: AppConfig,
                mode: str) -> int: ...
def run_execute(source_path: Path, config: AppConfig,
                mode: str, project_root: Path) -> int: ...
```

`mode` values:

- `attach`: require the exact source workbook to be open in a running Excel instance;
- `open`: create an isolated Excel instance and open the source with link updates disabled;
- `auto`: attempt exact-path attach first, then isolated open only if no matching workbook is open. Do not fall back when an ambiguous open-workbook identity exists.

Dry-run sequence:

1. Validate CLI/configuration and source path.
2. Connect to Excel using the selected mode.
3. Open/identify source without editing.
4. Collect fingerprint and verify `.xlsb`.
5. Verify all required sheets.
6. Read header snapshots and locate unique blocks in all four sheets.
7. Verify A08 does not already exist.
8. Record external-link, pivot, connection, name, and VBA observations.
9. Print resolved codes and discovered coordinates.
10. Mark ready only if every required check passes.
11. Close only workbooks/applications owned by the automation; never save source.

Execute sequence:

1. Run the complete preflight again; do not trust an older dry-run report.
2. Collect the source fingerprint and source hash.
3. Create run paths and initialize manifest/logging.
4. Call source workbook `SaveCopyAs(working_path)`.
5. Detach from the source without saving it. If attached to user Excel, leave the user's source workbook and Excel application open and untouched.
6. Open the working copy in an automation-owned isolated Excel instance where feasible.
7. Confirm working copy identity, format, sheets, and preflight blocks again.
8. Capture application state, then enable manual calculation, disable events and screen updates, and use alerts suppression only for known non-destructive operations.
9. Update and locally validate `VD Total`.
10. Update and locally validate `MX Total`.
11. Update and locally validate `DA Total`.
12. Update `Total PL` using proven row mappings.
13. Run one full calculation rebuild and wait for completion.
14. Run complete pre-save validation.
15. If validation fails, do not publish.
16. Save the working workbook through Excel COM.
17. Close it cleanly.
18. Reopen the same working path through Excel COM with link updates disabled.
19. Recalculate only if necessary to obtain stable validation values; record whether this occurred.
20. Run post-reopen validation.
21. Close the workbook and Excel instance cleanly.
22. Reconfirm source hash unchanged.
23. Publish the closed validated working file to a unique output filename.
24. Verify working and final hashes match.
25. Write final report and manifest with `SUCCESS`.

Every exception path must update the manifest, restore Excel state, close automation-owned Excel resources, leave user-owned Excel running, retain diagnostics, and return non-zero.

### 8.16 `app.py`

Use `argparse`. Required interface:

```text
python app.py --file PATH --year 2026 --month 8 --dry-run [--mode auto|attach|open]
python app.py --file PATH --year 2026 --month 8 --execute [--mode auto|attach|open]
```

Add:

```text
--config PATH       default: config.yaml
--verbose
```

`--dry-run` and `--execute` must be mutually exclusive and one must be required. Resolve the input path to an absolute path. Refuse a non-`.xlsb` source.

Exit codes:

```text
0  success / dry-run ready
2  CLI or configuration error
3  preflight not ready
4  safe execution failure
5  validation failure
6  publication failure
```

Catch `PLAutomationError` at the boundary, print a short safe failure message, and keep detailed diagnostics in the run log. Unexpected exceptions must also fail closed.

### 8.17 `RUN_A08.bat`

Requirements:

1. Use the Python interpreter from a local `.venv` if present; otherwise explain how to create it.
2. Accept a workbook dropped onto the batch file as `%~1`.
3. If no argument is supplied, prompt for a path.
4. Strip surrounding quotes safely.
5. Run dry-run first.
6. Continue only when dry-run returns exit code 0.
7. Ask the user for explicit `Y` confirmation before execute.
8. Run execute with year 2026 and month 8.
9. Open the output folder only after exit code 0.
10. Keep the console open on failure so the user can read the message.

The batch file must not delete files, overwrite the source, or attempt to bypass NASCA prompts.

## 9. Configuration file

Use this safe initial `config.yaml`:

```yaml
workbook:
  target_sheets:
    - "VD Total"
    - "MX Total"
    - "DA Total"
  total_sheet: "Total PL"

run:
  year: 2026
  month: 8

safety:
  update_external_links: false
  refresh_pivots: false
  overwrite_existing_actual: false
  keep_failed_workbook: true
  reopen_after_save: true

validation:
  numeric_tolerance: 0.01
  require_all_sheets: true
  require_unique_header_match: true
  calculation_timeout_seconds: 1800
```

The resolved `T08`, `S08`, `A08`, and `2026.008` must be derived and printed. A conflicting manually supplied version code must be rejected rather than silently accepted.

## 10. Dry-run report contract

The dry-run output must contain:

```text
SOURCE
MODE: DRY RUN
EXCEL ACCESS MODE
YEAR / MONTH / PERIOD
TARGET / FORECAST / ACTUAL VERSION
SOURCE FILE FORMAT
SOURCE SHA-256 status

For each of VD Total, MX Total, DA Total, Total PL:
  T08 column number and Excel letters
  T08 % column
  S08 column
  S08 % column
  insertion column
  September boundary
  header row numbers
  last used row
  August merge address
  existing A08 status
  evidence used to prove the match

External-link update: DISABLED
Pivot refresh: DISABLED
Source modification: PROHIBITED
READY TO EXECUTE: YES or NO
```

Column letters are display-only. Logic must retain numeric coordinates and rediscover them at execution time.

## 11. Test strategy

### 11.1 Unit tests that do not require Excel

#### Constants/configuration

- period formatting for months 1, 8, 9;
- target October–December mapping;
- rejection of Actual October–December;
- execute-mode rejection of non-August scope;
- rejection of unsafe config flags and unknown keys.

#### Header discovery

Create in-memory header snapshots for:

- one valid August block;
- merged month and period headers;
- duplicate T08 candidates where only one has period `2026.008`;
- two fully valid candidates, which must raise ambiguity;
- missing `%` after T08;
- missing S08 pair;
- missing September boundary;
- existing A08;
- whitespace and case normalization.

#### Formula rewrite

Test:

```text
=SUMIFS(A:A,B:B,"T08")       -> exact replacement
=IF(C1="T08",1,0)            -> exact replacement
=SUMIFS(A:A,B:B,"XT08")      -> unchanged
=T08                          -> unchanged
='T08 Sheet'!A1              -> unchanged
=IF(C1="T080",1,0)           -> unchanged
header-driven formula         -> unchanged
multiple exact literals       -> every exact literal replaced and counted
```

#### Total PL reference parsing

Test quoted sheet names, `$` absolute markers, lowercase column letters, multi-term sums, parenthesized formulas, and formulas containing irrelevant same-sheet references. Reject external workbook references, `INDIRECT`, and unprovable dynamic addressing.

Verify that only source columns change and row numbers remain byte-for-byte equivalent in the rewritten references.

#### Validation helpers

- values equal within 0.01;
- values outside tolerance fail;
- `None`, text, booleans, NaN, and Excel error representations are handled explicitly;
- required failed checks make the report fail;
- warning-only checks do not override a required pass.

### 11.2 COM integration tests

Integration tests must be marked, skipped clearly when Excel is unavailable, and must never use the production source as an editable test target.

Run in this order:

1. Attach/open identity test.
2. Application state restoration test, including forced exception.
3. `SaveCopyAs` test and source-hash comparison.
4. Dry-run on the real workbook with zero edits.
5. VD-only update on a working copy.
6. MX-only update on a fresh working copy.
7. DA-only update on a fresh working copy.
8. All three business sheets on a fresh working copy.
9. Total PL mapping audit without assignment.
10. Total PL update on a fresh working copy.
11. Full calculation and reconciliation.
12. Save/close/reopen validation.
13. Repeated run against updated output; it must stop before insertion.
14. Forced ambiguous-header failure; no final output.
15. Forced calculation timeout; no final output and Excel state restored.

### 11.3 Test evidence

For every real-workbook stage, retain:

- source and working hashes;
- run ID;
- detected block coordinates;
- formula counts;
- reconciliation mismatch count;
- Excel version;
- save/reopen result;
- source-unchanged confirmation.

Do not retain confidential workbook values beyond the minimum reconciliation evidence approved by the plan.

## 12. Numbered implementation tasks

### Task 1 — Scaffold the local project

**Description:** Create the directory structure, dependency files, safe default configuration, and importable package.

**Dependencies:** None.

**Files:** project root files, `src/__init__.py`.

**Acceptance criteria:**

- [ ] The documented structure exists.
- [ ] Dependencies are limited to `pywin32`, `PyYAML`, and test tooling such as `pytest`.
- [ ] Importing `src` does not start Excel or touch a workbook.

**Verification:**

```powershell
python -m pytest --collect-only
python -c "import src"
```

### Task 2 — Implement constants, models, and errors

**Description:** Establish all stable contracts before writing COM logic.

**Dependencies:** Task 1.

**Files:** `constants.py`, `models.py`, `errors.py`, related unit tests.

**Acceptance criteria:**

- [ ] A08 codes resolve exactly.
- [ ] Unsafe/unsupported month cases raise explicit errors.
- [ ] Models serialize to JSON-safe dictionaries.

**Verification:**

```powershell
python -m pytest tests/unit/test_constants.py
```

### Task 3 — Implement strict configuration loading

**Description:** Load YAML, apply CLI overrides, and reject settings that violate V1 safety.

**Dependencies:** Task 2.

**Files:** `config.py`, `config.yaml`, `test_config.py`.

**Acceptance criteria:**

- [ ] Safe config loads.
- [ ] Unknown keys and unsafe true flags fail clearly.
- [ ] Execute mode rejects non-August runs.

**Verification:**

```powershell
python -m pytest tests/unit/test_config.py
```

### Checkpoint A — Foundation

- [ ] All unit tests pass.
- [ ] No module has opened Excel.
- [ ] No workbook file has been created or modified.

### Task 4 — Implement Excel session ownership and state restoration

**Description:** Safely attach/open Excel and guarantee cleanup behavior.

**Dependencies:** Tasks 2–3.

**Files:** `excel_session.py`, `test_excel_session.py`.

**Acceptance criteria:**

- [ ] Exact-path workbook attachment works.
- [ ] Automation-owned Excel quits; user-owned Excel does not.
- [ ] Application settings restore after success and forced failure.

**Verification:** Run the marked Excel integration tests and manually confirm the user's Excel instance remains open in attach mode.

### Task 5 — Implement fingerprints and workbook audit

**Description:** Capture sufficient evidence before and after edits.

**Dependencies:** Task 4.

**Files:** `workbook_audit.py`, unit/integration audit tests.

**Acceptance criteria:**

- [ ] Required metadata and source SHA-256 are recorded.
- [ ] Optional protected properties degrade to warnings.
- [ ] External links are read without being updated.

**Verification:** Audit the source read-only and inspect the serialized fingerprint.

### Task 6 — Implement safe file transaction

**Description:** Create working copies through native `SaveCopyAs` and collision-safe final paths.

**Dependencies:** Tasks 4–5.

**Files:** `file_transaction.py`, `test_save_copy_transaction.py`.

**Acceptance criteria:**

- [ ] Working copy is created by Excel and opens as `.xlsb`.
- [ ] Source hash and timestamp remain unchanged.
- [ ] Source, working, and final paths cannot collide.

**Verification:** Run the copy-only integration test, close/reopen the working copy, and compare hashes.

### Checkpoint B — Safe Excel/file layer

- [ ] A working `.xlsb` copy can be created without editing the source.
- [ ] Excel ownership and state restoration are proven.
- [ ] No production structural edit exists yet.

### Task 7 — Implement bulk header snapshots

**Description:** Read the top header area efficiently and preserve merged-header context.

**Dependencies:** Task 2.

**Files:** `header_discovery.py`, `test_header_discovery.py`.

**Acceptance criteria:**

- [ ] Values are read in one bulk call per sheet.
- [ ] Merged month/period values resolve correctly.
- [ ] Pure snapshots are testable without Excel.

**Verification:** Unit tests plus read-only snapshots from all four real sheets.

### Task 8 — Implement unique August block discovery and idempotency

**Description:** Locate T08/S08/August/September using multiple mandatory anchors.

**Dependencies:** Task 7.

**Files:** `block_locator.py`, `test_block_locator.py`, `test_idempotency.py`.

**Acceptance criteria:**

- [ ] Exactly one valid block is returned per target sheet on the source.
- [ ] Duplicate or incomplete matches fail closed.
- [ ] Existing A08 is detected before insertion.

**Verification:** Run all discovery fixtures and a full read-only dry-run against the source.

### Task 9 — Implement the dry-run workflow and report

**Description:** Deliver the first usable safety gate before any workbook edit code.

**Dependencies:** Tasks 3–8.

**Files:** `reporting.py`, partial `workflow.py`, `app.py`.

**Acceptance criteria:**

- [ ] Dry-run prints all resolved codes and four proven blocks.
- [ ] Dry-run performs no save, insertion, formula assignment, or publication.
- [ ] Any ambiguity produces `READY TO EXECUTE: NO` and non-zero status.

**Verification:** Compare source hash, size, and modified time before/after dry-run.

### Checkpoint C — Read-only proof

- [ ] The real workbook has a unique valid August block in every target sheet.
- [ ] The discovery coordinates and evidence have been reviewed by a human.
- [ ] A08 does not already exist.
- [ ] Do not continue if this checkpoint is not approved.

### Task 10 — Implement native formula cloning and conservative rewrite

**Description:** Copy formula/style pairs with Excel and rewrite only exact literal criteria.

**Dependencies:** Tasks 4 and 8.

**Files:** `formula_clone.py`, `test_formula_rewrite.py`.

**Acceptance criteria:**

- [ ] Exact quoted rewrite unit cases pass.
- [ ] Header-driven formulas remain unchanged.
- [ ] Unsupported special formulas stop with cell evidence.

**Verification:** Unit tests, then a disposable integration fixture with ordinary and special formulas.

### Task 11 — Implement column properties and August merge repair

**Description:** Preserve column-level formatting and surgically extend the August header.

**Dependencies:** Task 4.

**Files:** `merge_formatting.py`, focused integration tests.

**Acceptance criteria:**

- [ ] Width, hidden state, and outline level are restored.
- [ ] Only the August merge may be unmerged/remerged.
- [ ] Overlapping unrelated merges fail closed.

**Verification:** Test on a controlled workbook fixture and compare merge addresses before/after.

### Task 12 — Implement and prove VD Total update

**Description:** Complete the first vertical slice: discover, copy, insert, rewrite, and locally validate VD Total.

**Dependencies:** Tasks 8, 10, 11.

**Files:** `business_sheet_updater.py`, `validation.py`, integration test.

**Acceptance criteria:**

- [ ] One A08 pair appears before September.
- [ ] Formula counts/topology and formatting match the T08 pattern.
- [ ] The working copy saves, closes, and reopens without a repair warning.

**Verification:** Run VD-only integration test on a fresh working copy and preserve its report.

### Task 13 — Prove MX Total update

**Description:** Apply the same shared updater to MX Total without sheet-specific coordinates.

**Dependencies:** Task 12.

**Files:** shared updater only if a genuine generic fix is required; MX integration test.

**Acceptance criteria:**

- [ ] MX is handled by semantic discovery, not a new hard-coded branch.
- [ ] Local validation passes.
- [ ] VD behavior remains covered by tests.

**Verification:** MX-only test on a fresh copy, then rerun VD tests.

### Task 14 — Prove DA Total and combined business update

**Description:** Validate DA and then the ordered three-sheet workflow.

**Dependencies:** Task 13.

**Files:** shared updater/workflow tests.

**Acceptance criteria:**

- [ ] DA local validation passes.
- [ ] VD → MX → DA completes on one fresh working copy.
- [ ] September shifts correctly on all three sheets.

**Verification:** DA-only and combined-business integration tests.

### Checkpoint D — Business sheets complete

- [ ] Each business sheet has exactly one valid A08 pair.
- [ ] No residual unexplained quoted T08 criteria remain in new amount formulas.
- [ ] The combined workbook saves and reopens cleanly.
- [ ] Source remains unchanged.

### Task 15 — Implement conservative Total PL reference parsing

**Description:** Parse and rewrite only provable cross-sheet A1 references.

**Dependencies:** Tasks 2 and 8.

**Files:** `total_pl_updater.py`, `test_total_pl_mapping.py`.

**Acceptance criteria:**

- [ ] Allowed formulas parse with exact sheet, column, row, and absolute markers.
- [ ] Rewriting changes columns only and preserves rows.
- [ ] Dynamic/external/unprovable formulas are rejected.

**Verification:** Comprehensive unit formula corpus.

### Task 16 — Audit real Total PL row mappings without writing

**Description:** Prove the workbook's analogous T08/S08 formula lineage before enabling assignment.

**Dependencies:** Tasks 9, 14, 15.

**Files:** reporting/audit path; no production write required.

**Acceptance criteria:**

- [ ] Every intended Total PL formula row is classified.
- [ ] Every business-total row has a provable mapping or a precise blocker report.
- [ ] Blank/static rows are explicitly identified.

**Verification:** Human review of the mapping audit. Stop here if any required row cannot be proven.

### Task 17 — Implement Total PL A08 update

**Description:** Insert the pair and create A08 formulas from proven business-sheet lineage.

**Dependencies:** Tasks 14 and 16.

**Files:** `total_pl_updater.py`, integration test.

**Acceptance criteria:**

- [ ] Total PL has one A08 pair before September.
- [ ] Amount formulas reference the newly created business A08 columns using preserved row mapping.
- [ ] Percentage formulas preserve the analogous Excel logic.

**Verification:** Run Total PL update on a fresh working copy, then inspect the mapping report before calculation.

### Task 18 — Implement calculation and full validation

**Description:** Calculate once and enforce all structural, formula, reconciliation, preservation, and control checks.

**Dependencies:** Task 17.

**Files:** `calculation.py`, `validation.py`, unit/integration tests.

**Acceptance criteria:**

- [ ] Calculation reaches complete state or times out safely.
- [ ] Total PL mismatch count is zero within tolerance.
- [ ] Required validation failures prevent save/publication.

**Verification:** Full validation test plus forced mismatch and forced timeout tests.

### Task 19 — Implement save/reopen validation and publication

**Description:** Complete the transactional guarantee and publish only a proven closed workbook.

**Dependencies:** Task 18.

**Files:** `workflow.py`, `file_transaction.py`, `reporting.py`, full integration test.

**Acceptance criteria:**

- [ ] Working file saves and closes through Excel.
- [ ] Reopened file passes critical validations again.
- [ ] Final file hash equals the validated closed working file hash.

**Verification:** Full end-to-end test on a fresh source-derived working copy.

### Task 20 — Implement reports, launcher, and user documentation

**Description:** Make the tool operable by a non-technical user and diagnosable after failure.

**Dependencies:** Task 19.

**Files:** `README.md`, `RUN_A08.bat`, report fixtures/examples.

**Acceptance criteria:**

- [ ] Drag/drop or prompted path runs dry-run first.
- [ ] Execute requires confirmation.
- [ ] Success opens output; failure leaves an actionable report.

**Verification:** Manual launcher acceptance test with success, cancel, bad path, and preflight failure.

### Task 21 — Run idempotency and forced-failure acceptance suite

**Description:** Prove that repeated and unsafe runs stop without damage.

**Dependencies:** Task 20.

**Files:** test fixtures/reports only unless a defect is found.

**Acceptance criteria:**

- [ ] Second run detects A08 and performs no insertion.
- [ ] Ambiguous-header run publishes nothing.
- [ ] Source hash is unchanged after every case.

**Verification:** Integration Tests 13–15 with retained evidence.

### Checkpoint E — Production candidate

- [ ] All unit tests pass.
- [ ] All applicable COM integration tests pass.
- [ ] The real source dry-run is uniquely ready.
- [ ] The final candidate reopens without repair/corruption warnings.
- [ ] Total PL reconciliation has zero unresolved mismatches.
- [ ] External links and existing controls did not regress.
- [ ] Source SHA-256 is unchanged.
- [ ] Human owner reviews the final report and workbook before production use.

## 13. Required run manifest phases

Update `phase` as execution proceeds:

```text
INITIALIZING
PREFLIGHT
FINGERPRINTED
WORKING_COPY_CREATED
WORKING_COPY_OPENED
VD_UPDATED
MX_UPDATED
DA_UPDATED
TOTAL_PL_UPDATED
CALCULATED
PRE_SAVE_VALIDATED
SAVED_AND_CLOSED
REOPENED
POST_REOPEN_VALIDATED
SOURCE_REVERIFIED
PUBLISHED
COMPLETE
```

On failure, retain the last completed phase and include the failing operation separately.

## 14. Required human-readable failure messages

Use messages following this pattern:

```text
FAILED SAFELY.
The original workbook was not changed.
Phase: <phase>
Reason: <plain specific reason>
Working copy retained: <path or NO>
Report: <path>
```

Examples:

```text
Reason: August T08 block was found twice in MX Total and both candidates matched period 2026.008.
```

```text
Reason: Total PL row 168 uses INDIRECT, so its business-sheet row mapping could not be proven safely.
```

```text
Reason: Total PL A08 reconciliation failed on 2 rows; the final workbook was not published.
```

## 15. Code quality rules

- Add type hints to public functions.
- Keep COM calls behind COM-facing modules; pure algorithms receive plain Python data.
- Use `Path` for filesystem paths and canonicalize before comparisons.
- Use structured logging with run ID, phase, sheet, and operation.
- Never use broad `except Exception: pass`.
- Wrap COM exceptions with the domain error taxonomy while preserving traceback in logs.
- Keep functions focused; orchestration belongs in `workflow.py`.
- Do not cache workbook coordinates between dry-run and execute without rediscovery.
- Do not introduce concurrency around Excel COM; COM workbook operations remain sequential on their initialized thread.
- Do not optimize before measuring. Four header bulk reads and four two-column native copies are the intended scale.

## 16. Production acceptance matrix

| Area | Required evidence | Pass condition |
|---|---|---|
| Source safety | Before/after SHA-256 | Identical |
| File format | Extension and Excel FileFormat | Remains `.xlsb` / Excel binary |
| VD Total | Discovery + local validation | One correct A08 pair |
| MX Total | Discovery + local validation | One correct A08 pair |
| DA Total | Discovery + local validation | One correct A08 pair |
| Total PL | Mapping + reconciliation | One correct A08 pair; zero mismatches |
| Formula criteria | Formula audit | A08 used where T08 literal existed |
| Percentages | Formula topology/error scan | Correct A08 references; no new errors |
| August header | Merge comparison | Covers new A08 pair only |
| September | Structural discovery | Still present immediately after A08 |
| Links | Before/after normalized list | No additions/path changes/breaks |
| Pivots | Run configuration/log | No refresh requested |
| Controls | Before/after captured values | No unexplained regression |
| Calculation | State and elapsed time | Completed within timeout |
| Reopen | Critical validation report | Passed without repair warning |
| Idempotency | Second-run result | Stops before edit |
| Reporting | Report + manifest | Complete and actionable |

## 17. Final Definition of Done

The implementation is complete only when a non-technical user can select the original protected workbook, run the launcher, approve execution after a successful dry run, and receive a separate validated `.xlsb` containing:

```text
VD Total : August -> T08 | % | S08 | % | A08 | %
MX Total : August -> T08 | % | S08 | % | A08 | %
DA Total : August -> T08 | % | S08 | % | A08 | %
Total PL : August -> T08 | % | S08 | % | A08 | %
```

and every mapped row satisfies:

```text
Total PL A08 = VD Total A08 + MX Total A08 + DA Total A08
```

within the configured tolerance, with:

- the original workbook byte-for-byte unchanged;
- no corruption or Excel repair warning;
- native formulas and formatting preserved;
- no unintended refresh or link changes;
- successful close/reopen validation;
- a complete `run_report.txt` and `run_manifest.json` proving the outcome;
- safe refusal on a repeated run.

Anything less is a partial implementation, not production completion.
