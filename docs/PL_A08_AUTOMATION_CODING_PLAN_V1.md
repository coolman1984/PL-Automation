# P&L Actual-Month Column Automation — Coding Plan V1

## 0. Mission

Build a **safe Windows desktop automation** that updates the existing company P&L workbook **without rebuilding or converting it**.

The current first production use case is **Actual August 2026 = `A08`**.

The automation must:

1. Open or attach to the real Microsoft Excel desktop application through **COM automation** so NASCA/DRM-protected workbooks can be handled through the authorized Excel process.
2. Never edit the user's original workbook in place.
3. Create a safe working copy in the original `.xlsb` format.
4. In these three sheets:
   - `VD Total`
   - `MX Total`
   - `DA Total`

   insert **two new August columns**:
   - `A08` amount / quantity column
   - the matching `%` column

5. The new `A08` pair in those three sheets must use the **same calculation logic and formatting pattern as the August `T08` pair**, but must retrieve/calculate the **Actual `A08` version**, not Target `T08`.
6. In `Total PL`, insert the corresponding two August columns and make its `A08` values reconcile to the combined `A08` values of:
   - `VD Total`
   - `MX Total`
   - `DA Total`
7. Preserve all existing workbook logic, formulas, styles, merged cells, grouping, hidden rows/columns, pivots, names, external links, formulas, macros/VBA if present, and workbook format.
8. Validate the result, close it, reopen it in Excel, validate again, and only then publish the final output file.

This is a **surgical workbook edit**, not a workbook rewrite.

---

# 1. Non-negotiable safety rules

## 1.1 Only Excel COM may write the workbook

Use:

- Python
- `pywin32`
- Microsoft Excel desktop COM object model

Do **not** write/save this workbook using:

- `openpyxl`
- `pandas`
- `pyxlsb`
- `xlsxwriter`
- LibreOffice
- direct ZIP/XML/binary manipulation

Those tools may be useful for separate analytical tasks, but **must not be used to save this production `.xlsb` workbook**.

Reason: this workbook is a very large binary Excel model with shared formulas, pivots, external links, formatting, merged cells, calculation-chain behavior, and possible NASCA/DRM handling. Excel itself must remain the file-format authority.

## 1.2 Never modify the source file

The source file is immutable.

Required flow:

`SOURCE.xlsb -> Excel SaveCopyAs -> WORKING.xlsb -> edit -> validate -> FINAL.xlsb`

If any step fails, the source file remains untouched.

## 1.3 No guessing

If the expected August structure cannot be uniquely detected, **stop**.

Do not guess:

- a column letter
- a header row
- a P&L row
- a formula source
- a version code

The automation must report exactly what it could not identify.

## 1.4 No destructive global find/replace

Never replace `T08` with `A08` across the workbook.

Only modify formulas inside the newly created `A08` ranges, and only when the `T08` token is the actual version criterion that must become `A08`.

## 1.5 No pivot refresh by default

Do not refresh pivot tables or external links unless the future workflow explicitly requires it.

This phase only adds the new calculated reporting columns.

## 1.6 Fail closed

If validation is not successful, do not publish the final file.

Move the failed working copy into a `failed_runs` folder and keep the log for diagnosis.

---

# 2. Business rules that are source of truth

## 2.1 Version terminology

### Target

For January through September:

- January = `T01`
- February = `T02`
- March = `T03`
- April = `T04`
- May = `T05`
- June = `T06`
- July = `T07`
- August = `T08`
- September = `T09`

For October through December:

- October = `T0A`
- November = `T0B`
- December = `T0C`

### Actual

`A` means Actual.

Current implemented case:

- August Actual = `A08`

Do **not** invent October-December Actual codes until the business owner defines them.

### Forecast

`S` means Forecast.

Example:

- August Forecast = `S08`

### Yearly plan version

`P` is a **yearly plan revision**, not a month.

Examples:

- `P01` = first yearly plan version
- `P02` = second yearly plan version
- `P03` = third yearly plan version

Never interpret `P01` as January.

## 2.2 Current A08 insertion rule

In each of:

- `VD Total`
- `MX Total`
- `DA Total`

August currently contains a block conceptually similar to:

`T08 | % | S08 | %`

Insert the new Actual pair so August becomes:

`T08 | % | S08 | % | A08 | %`

The new pair must be inserted **before the September block**.

## 2.3 Formula rule for the three business sheets

For `VD Total`, `MX Total`, and `DA Total`:

- the new `A08` amount column follows the **same equation structure as August `T08`**;
- the new `%` column follows the **same percentage logic as the August `T08` percentage column**;
- where the copied formula contains the version criterion `T08`, that criterion must become `A08`;
- all relative references must be allowed to move exactly as Excel would move them during a normal native copy.

The result is not a hard-coded value column. It is a real calculated Actual column.

## 2.4 `Total PL` rule

`Total PL` gets the same two new August columns:

`A08 | %`

For Actual August, `Total PL` must reconcile to the Actual August totals of:

`VD Total + MX Total + DA Total`

Do not independently invent a second A08 business calculation engine in `Total PL`.

Where an existing analogous `Total PL` formula shows the correct cross-sheet row mapping, reuse that formula lineage and point it to the newly created `A08` columns in the three business sheets.

If no reliable analogous cross-sheet mapping can be discovered, stop and produce an audit report instead of guessing row numbers.

## 2.5 Original workbook model remains authoritative

Do not change:

- `DB`
- `Guide`
- `PV`
- `PV QTY`
- `rawdata`
- other business P&L sheets
- Summary or executive sheets

unless a later business instruction explicitly adds them to the scope.

---

# 3. Technology architecture

## 3.1 Runtime

Use a tested Windows x64 Python environment.

Recommended baseline:

- Python 3.12 or 3.13 x64
- `pywin32`

Keep dependencies intentionally small.

DuckDB is **not required in Phase 1** because this automation is not rebuilding the source-data calculation engine. It is performing a controlled Excel structural/formula update. Add DuckDB later only if the automation begins transforming large source datasets outside Excel.

## 3.2 Excel access modes

Implement two Excel access modes.

### Mode A — attach to already-open workbook

This should be the preferred NASCA-safe mode.

Workflow:

1. User opens the protected workbook normally in Microsoft Excel.
2. NASCA performs its normal authorization/decryption behavior.
3. Automation connects to the existing Excel application/workbook through COM.
4. Automation creates a copy and edits the copy.

This avoids trying to bypass DRM and uses the same authorized Excel session as the user.

### Mode B — COM opens the workbook

If NASCA allows it:

1. Start a new Excel instance using COM.
2. Call `Workbooks.Open(...)`.
3. Do not update external links.
4. If Excel/NASCA displays an interactive authorization prompt, allow the user to complete it.

Never attempt to bypass DRM, Protected View, authentication, or corporate security controls.

## 3.3 Main components

Implement these logical modules:

1. Excel session manager
2. Source/working/final file transaction manager
3. Workbook fingerprint/audit collector
4. Header/block discovery engine
5. August business-sheet insertion engine
6. Formula clone/version rewrite engine
7. `Total PL` builder
8. Formatting/merge repair engine
9. Calculation controller
10. Validation engine
11. Logging/run-manifest engine
12. User entry point

---

# 4. Project structure

Create approximately this structure:

```text
pl_actual_automation/
│
├─ app.py
├─ requirements.txt
├─ README.md
├─ RUN_A08.bat
├─ config.yaml
│
├─ src/
│  ├─ __init__.py
│  ├─ constants.py
│  ├─ models.py
│  ├─ excel_session.py
│  ├─ file_transaction.py
│  ├─ workbook_audit.py
│  ├─ header_discovery.py
│  ├─ block_locator.py
│  ├─ formula_clone.py
│  ├─ business_sheet_updater.py
│  ├─ total_pl_updater.py
│  ├─ merge_formatting.py
│  ├─ calculation.py
│  ├─ validation.py
│  ├─ reporting.py
│  └─ errors.py
│
├─ tests/
│  ├─ test_month_codes.py
│  ├─ test_header_discovery.py
│  ├─ test_formula_rewrite.py
│  ├─ test_idempotency.py
│  └─ test_validation_logic.py
│
├─ logs/
├─ output/
├─ work/
├─ backups/
└─ failed_runs/
```

Do not over-engineer this into a web app or database service.

The first version is a local Windows automation.

---

# 5. Configuration

Use a small `config.yaml` such as:

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
  actual_version: "A08"
  target_version: "T08"
  forecast_version: "S08"
  period: "2026.008"

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
```

The application may derive `T08`, `A08`, and `2026.008` from year/month, but the resolved values must be printed in the dry-run report before any edit.

---

# 6. Excel session implementation

## 6.1 COM creation

Use `win32com.client.DispatchEx("Excel.Application")` when starting an isolated Excel instance.

When attaching to an already-open protected workbook, use the running Excel object and identify the workbook by exact full path or user-selected workbook.

## 6.2 Preserve Excel state

Before changing settings, record:

- `Application.Calculation`
- `ScreenUpdating`
- `EnableEvents`
- `DisplayAlerts`
- `AskToUpdateLinks`
- application visibility state

Restore them in a `finally` block even if the run fails.

## 6.3 Performance settings during the edit

Temporarily set:

- calculation = manual
- screen updating = false
- events = false
- display alerts = false only for known non-destructive actions

Do not leave Excel in manual calculation mode after the run.

## 6.4 Close rules

The automation must know whether it created the Excel application or attached to the user's existing Excel application.

- If it created Excel, it may quit Excel after closing its workbook.
- If it attached to the user's Excel, it must **not quit the user's Excel application**.

---

# 7. Safe file transaction

## 7.1 Run folder

Each run gets a unique folder:

```text
work/2026-08-27_143500_A08_<short-id>/
```

Store:

- source metadata
- working workbook
- run manifest
- logs
- validation report

## 7.2 Source fingerprint

Before editing, record:

- full path
- filename
- file size
- modified timestamp
- SHA-256 of source file if Windows access permits it
- workbook sheet names
- sheet count
- workbook format/file format
- external link sources
- named item count
- pivot table count per sheet where practical
- connection count where practical
- whether VBA project is present if detectable

This is the pre-edit fingerprint.

## 7.3 Create the working copy

Preferred method:

1. Open/attach source workbook.
2. Call Excel's native `SaveCopyAs(working_path)`.
3. Close/detach source without saving.
4. Open the working copy through Excel COM.
5. Perform all edits on the working copy.

If `SaveCopyAs` is blocked by NASCA, stop and report that the corporate protection policy blocks automated copy creation. Do not work around it silently.

## 7.4 Final publication

After successful validation:

1. Save working workbook.
2. Close it.
3. Reopen it in Excel.
4. Perform post-reopen validation.
5. Close it cleanly.
6. Move/copy it to a final name such as:

```text
★Final PL Statement S08 T09 V4(1)__A08_UPDATED.xlsb
```

Never replace the original file by default.

---

# 8. Header and block discovery — absolutely no fixed column letters

The workbook is huge and its columns may shift between versions.

Do not hard-code positions like `XQ`, `XR`, etc.

## 8.1 Search scope

For each target sheet, read the top header area in one COM call, for example:

- rows 1 to 30
- columns from first used column to last used column

Use a 2D bulk value read.

Normalize header text:

- trim spaces
- convert to string safely
- uppercase only for matching

## 8.2 Identify the correct August block

There may be more than one `T08` in the sheet, for example August target and September's previous target comparison.

Therefore a version label alone is not enough.

A valid August candidate must satisfy multiple anchors:

1. version header exactly `T08`;
2. nearby period header = `2026.008`;
3. month group above it corresponds to August if a month label exists;
4. `T08` is followed by its `%` column;
5. nearby August block also contains `S08` followed by `%`;
6. the next logical block is September or period `2026.009`.

Expected conceptual structure:

```text
August
T08 | % | S08 | %
September
...
```

After insertion:

```text
August
T08 | % | S08 | % | A08 | %
September
...
```

If there is not exactly one high-confidence match, stop.

## 8.3 Discovery result object

Return a structured object such as:

```python
MonthBlock(
    sheet="VD Total",
    year=2026,
    month=8,
    period="2026.008",
    target_col=...,       # T08 amount
    target_pct_col=...,   # T08 %
    forecast_col=...,     # S08 amount
    forecast_pct_col=..., # S08 %
    insert_at_col=...,    # before September
    version_header_row=...,
    period_header_row=...,
    month_header_row=...,
    last_used_row=...
)
```

Print this in dry-run output.

---

# 9. Idempotency — prevent duplicate A08 columns

Before editing a sheet, search the same August block for `A08`.

If `A08` already exists:

- default behavior = stop and report `A08 already exists`;
- do not add another pair;
- only a future explicit `--replace-existing` option may update an existing pair.

The first production version should **not implement destructive replacement** unless specifically requested.

---

# 10. Insert A08 in `VD Total`, `MX Total`, `DA Total`

Process these three sheets **before** `Total PL`.

## 10.1 Snapshot source pair

For each business sheet, identify:

- August `T08` amount column
- August `T08` `%` column
- rows 1 through last used row
- relevant column widths
- hidden state
- outline/group state where accessible
- number formats
- merged areas intersecting the August month header

## 10.2 Insert two worksheet columns

Insert two full Excel worksheet columns at `insert_at_col`, immediately after `S08 | %` and before September.

Use the Excel object model, not array rewriting of the whole sheet.

Allow Excel to update existing formula references naturally.

## 10.3 Clone the August T08 pair

Copy the **August T08 amount + % pair** into the newly inserted pair using Excel's native copy behavior.

The native Excel copy is important because it preserves:

- formulas
- relative-reference translation
- number formats
- conditional formatting
- borders
- fills
- fonts
- alignment
- cell protection
- comments/notes where applicable

Then explicitly copy/restore column width, hidden state, and any other column-level property that native range copy does not preserve correctly.

## 10.4 Set the new headers

Change only the new version header:

- amount column header -> `A08`
- percent column header -> `%`

Period remains:

- `2026.008`

Month remains:

- August

Do not change `T08` or `S08` headers.

## 10.5 Expand/repair August merged header

If the month label `August` is a merged cell spanning the old August block, ensure the merged area spans the newly added two columns as well.

Algorithm:

1. detect the existing month-header merge area before insertion;
2. after insertion, check whether Excel automatically expanded it;
3. if it did not, preserve the top-left value and formatting;
4. unmerge only that specific month-header area;
5. re-merge from the original August start column through the new A08 `%` column;
6. restore value/alignment/format.

Do not globally unmerge anything.

## 10.6 Convert the copied T08 logic to A08 logic

After native copy, inspect formulas in the new pair.

Goal:

- structure stays the same as T08;
- version criterion becomes A08 where required.

Use this rule:

### Relative references

Leave Excel's copied relative references alone.

### Literal version criteria

Where a formula contains the exact literal version criterion `"T08"`, change **that exact quoted criterion** to `"A08"` inside the new destination range only.

Example concept:

```text
SUMIFS(..., VersionRange, "T08", ...)
```

becomes:

```text
SUMIFS(..., VersionRange, "A08", ...)
```

Do not replace:

- text outside formulas
- existing T08 columns
- sheet names
- comparison formulas outside the new pair
- partial strings

Implement this with a formula-aware helper that operates only on formula cells in the new range.

## 10.7 Header-driven formulas

Some formulas may read the version from the header cell rather than hard-code `T08`.

In those formulas, simply changing the new header to `A08` is enough.

Do not perform unnecessary formula rewriting.

## 10.8 Formula audit

For each new pair, report:

- formula cell count in T08 amount column
- formula cell count in new A08 amount column
- formula cell count in T08 % column
- formula cell count in new A08 % column
- count of formulas in A08 still containing quoted `"T08"`
- count containing quoted `"A08"`

Any unexplained residual `"T08"` criteria in the new A08 amount formula set should be treated as a validation warning/failure depending on context.

---

# 11. Build A08 in `Total PL`

Run this only after all three business sheets have valid A08 columns.

## 11.1 Locate August Total PL block

Use the same semantic discovery rules:

- year/period `2026.008`
- `T08`
- `%`
- `S08`
- `%`
- before September

Do not hard-code column letters.

## 11.2 Insert A08 pair

Insert two full columns after the August `S08 | %` pair and before September.

Copy the most structurally appropriate existing August pair for styles, widths, percentage formatting, row formulas, merged headers, etc.

## 11.3 Amount formulas must come from the three business totals

The business owner's rule is authoritative:

```text
Total PL A08 = VD Total A08 + MX Total A08 + DA Total A08
```

The important problem is **row mapping**.

Do not assume the same row number is used in all four sheets.

### Preferred row-mapping method

Use an existing analogous `Total PL` column formula to discover the correct row references.

For every `Total PL` row that should contain an amount:

1. inspect the formula in the analogous August `S08` or `T08` amount cell;
2. determine whether it already references `VD Total`, `MX Total`, and/or `DA Total` using the correct row mapping;
3. if yes, reuse that formula lineage and change only the referenced source columns to the corresponding A08 columns created in each business sheet;
4. preserve the referenced row numbers exactly.

Example concept only:

```text
='VD Total'!XQ168+'MX Total'!AB168+'DA Total'!AA168
```

must become conceptually:

```text
='VD Total'!<A08-col><same-row>+
 'MX Total'!<A08-col><same-row>+
 'DA Total'!<A08-col><same-row>
```

Actual row/column coordinates must be discovered from the workbook.

### Fallback

If the analogous Total PL formula does not expose a reliable cross-sheet mapping, do **not** invent one from visible labels.

Stop the run and produce:

- sheet
- row
- P&L label
- current T08 formula
- current S08 formula
- reason mapping could not be proven

This protects the financial statement from silent wrong totals.

## 11.4 Total PL percentage formulas

The new Total PL `%` column should use the same percentage logic as the analogous August percentage column.

Prefer native formula cloning so relative references move to the newly created A08 amount cells.

If it is based on same-sheet ratios, do not replace it with a hard-coded percentage.

## 11.5 Blank/static rows

If the analogous Total PL amount row is intentionally blank, text, subtotal label, or formatting-only row, preserve that behavior.

Do not force a three-sheet sum into every row.

---

# 12. Formula cloning helper design

Create a reusable formula engine with these capabilities:

## 12.1 Native clone

```python
clone_range_with_excel(source_range, destination_range)
```

Uses Excel `Copy` / destination behavior.

## 12.2 Exact quoted version rewrite

```python
rewrite_exact_version_criteria(
    range_obj,
    old_version="T08",
    new_version="A08"
)
```

Rules:

- formula cells only;
- exact quoted token only;
- destination range only;
- no global workbook replacement.

## 12.3 Formula classification

Classify formulas into:

- no formula
- header-driven formula
- literal-version formula
- cross-sheet formula
- percentage formula
- subtotal/parent sum formula
- unsupported/special formula

This classification is for validation/reporting, not for reimplementing Excel's formula engine.

## 12.4 Special formula types

Detect and preserve:

- array formulas
- dynamic arrays if any
- merged cells
- formulas containing external workbook links
- formulas containing structured references
- formulas containing named ranges

For special formulas, prefer native Excel copy over string reconstruction.

---

# 13. Calculation strategy

Structural column insertion can change dependencies.

Use this sequence:

1. Excel calculation manual during edits.
2. Complete all four sheet updates.
3. Restore calculation setting to automatic if that was the original state.
4. Run one controlled workbook calculation.
5. Because columns were structurally inserted, run a full calculation rebuild if needed:
   - `CalculateFullRebuild()`
6. Wait until Excel reports calculation state complete.
7. Apply a sensible timeout and log elapsed calculation time.

Do not repeatedly calculate after every cell/range update.

---

# 14. Validation engine

The final file is not valid merely because Excel saved it.

Implement all checks below.

## 14.1 Structural validation

Verify:

- all original sheet names still exist;
- no extra unintended sheets were created;
- workbook remains `.xlsb`;
- target sheets exist;
- each target sheet now contains exactly one `A08` pair in the correct August block;
- September block still exists after A08;
- original T08 and S08 columns still exist;
- August month header covers the new pair correctly;
- no accidental duplicate A08 exists.

## 14.2 Formula validation — business sheets

For each of:

- `VD Total`
- `MX Total`
- `DA Total`

check:

1. A08 amount formula count is logically consistent with T08 amount formula count.
2. A08 percentage formula count is logically consistent with T08 percentage formula count.
3. formulas are not replaced by static values where T08 had formulas.
4. formula errors are scanned for:
   - `#REF!`
   - `#VALUE!`
   - `#NAME?`
   - `#DIV/0!` where not intentionally present in the source pattern
5. no new external link source was introduced.
6. new A08 formulas use A08 criteria where appropriate.

## 14.3 `Total PL` reconciliation

For each row where the new `Total PL A08` cell is a business-total formula:

calculate:

```text
expected = mapped VD A08 + mapped MX A08 + mapped DA A08
actual   = Total PL A08
```

Validate:

```text
abs(actual - expected) <= tolerance
```

Default financial tolerance:

```text
0.01
```

For quantity or integer-like rows, exact or tighter checking may be used where appropriate.

Produce a mismatch table containing:

- Total PL row
- P&L label
- actual value
- expected value
- difference
- source references

Zero unresolved mismatches are required for success.

## 14.4 Percentage validation

For the new A08 percentage column:

- compare formula topology to the T08/S08 percent pattern;
- ensure formula references point to A08 cells, not accidentally back to T08/S08;
- scan for formula errors.

## 14.5 Existing control checks

Where the workbook already contains checker cells such as Summary reconciliation controls, capture their pre-edit value and post-edit value.

If a checker was `OK` before the edit and becomes non-OK after the edit, fail validation unless the new A08 column is explicitly designed to change that checker.

## 14.6 External links

Record external links before and after.

The automation must not silently:

- add a new external workbook link;
- change an existing link path;
- break an existing link.

## 14.7 File integrity by reopen

After save:

1. close workbook;
2. reopen final working file in Excel through COM;
3. confirm all target sheets are readable;
4. confirm A08 blocks are still present;
5. run critical reconciliation again;
6. close cleanly.

Only after this step may the file be published as final.

---

# 15. Preflight / dry-run mode

The first command the user should run is a dry run.

Example:

```bat
python app.py --file "C:\Path\Final PL.xlsb" --year 2026 --month 8 --dry-run
```

Dry run must not modify the workbook.

It should print a concise report like:

```text
SOURCE: Final PL.xlsb
MODE: DRY RUN
YEAR: 2026
MONTH: August
TARGET VERSION: T08
FORECAST VERSION: S08
ACTUAL VERSION: A08
PERIOD: 2026.008

VD Total
  August T08: col ...
  August %:   col ...
  August S08: col ...
  S08 %:      col ...
  Insert at:  col ...
  Existing A08: NO

MX Total
  ...

DA Total
  ...

Total PL
  ...

DRM / workbook writable check: PASS
External link update: DISABLED
Pivot refresh: DISABLED

READY TO EXECUTE: YES
```

If any discovery result is ambiguous, `READY TO EXECUTE` must be `NO`.

---

# 16. Execute mode

Example:

```bat
python app.py --file "C:\Path\Final PL.xlsb" --year 2026 --month 8 --execute
```

Execution flow:

1. preflight
2. fingerprint source
3. create run folder
4. `SaveCopyAs` working file
5. open working file
6. disable calculation/events/screen updates
7. update `VD Total`
8. validate local VD A08 block
9. update `MX Total`
10. validate local MX A08 block
11. update `DA Total`
12. validate local DA A08 block
13. update `Total PL`
14. restore calculation and calculate once
15. run full validation
16. save
17. close
18. reopen
19. post-reopen validation
20. close
21. publish final output
22. write run report

If any step fails:

- stop;
- close working workbook without further edits;
- do not touch source;
- retain failed working copy and logs;
- return a clear failure reason.

---

# 17. Simple non-technical launcher

Create `RUN_A08.bat`.

The first version may accept a file dropped onto the batch file or prompt for a file path.

Better production UX:

1. user double-clicks `RUN_A08.bat`;
2. script asks the user to select or paste the workbook path;
3. it runs dry-run first;
4. shows readiness result;
5. user confirms execution;
6. final file is written to `output`;
7. output folder opens automatically only after success.

Do not require the user to edit Python source code.

---

# 18. Logging and run manifest

Each run must produce:

## Human-readable report

`run_report.txt`

Include:

- source
- output
- run ID
- start/end time
- Excel version
- attach/open mode
- detected headers and columns
- formula counts
- validation summary
- Total PL mismatch count
- external link before/after count
- final success/failure

## Machine-readable report

`run_manifest.json`

Suggested shape:

```json
{
  "run_id": "...",
  "source": "...",
  "output": "...",
  "year": 2026,
  "month": 8,
  "period": "2026.008",
  "target_version": "T08",
  "forecast_version": "S08",
  "actual_version": "A08",
  "sheets": {},
  "validations": {},
  "status": "SUCCESS"
}
```

Never log confidential cell values in bulk. Log only the minimum values needed to prove reconciliation and diagnose errors.

---

# 19. Error handling

Create explicit errors such as:

- `ExcelNotInstalledError`
- `WorkbookNotFoundError`
- `DRMOpenBlockedError`
- `WorkbookReadOnlyError`
- `MissingSheetError`
- `AmbiguousMonthBlockError`
- `ExistingActualColumnError`
- `FormulaCloneError`
- `TotalPLMappingError`
- `CalculationTimeoutError`
- `ValidationError`
- `SaveError`
- `ReopenValidationError`

User-facing messages must be plain and actionable.

Example:

```text
FAILED SAFELY.
The original workbook was not changed.
Reason: August T08 block was found twice in MX Total and could not be uniquely linked to period 2026.008.
See: logs\<run-id>\run_report.txt
```

---

# 20. Tests before production use

## 20.1 Unit tests

Test pure logic without Excel where possible:

- month-to-version mapping
- period formatting
- exact version-token replacement
- detection candidate scoring
- idempotency checks
- numeric tolerance logic

## 20.2 Integration tests with a copy of the real workbook

Run in this order:

### Test 1 — dry run only

Expected:

- unique August block in all four target sheets
- no edits

### Test 2 — copy creation only

Expected:

- source hash unchanged
- working `.xlsb` opens

### Test 3 — VD Total only

Expected:

- A08 pair inserted in correct place
- formulas copied
- workbook opens

### Test 4 — MX Total only

Same validation.

### Test 5 — DA Total only

Same validation.

### Test 6 — all three business sheets

Expected:

- no cross-sheet corruption
- September columns shift correctly

### Test 7 — Total PL

Expected:

- A08 pair inserted
- every mapped Total PL row reconciles to business A08 sources

### Test 8 — full save/reopen

Expected:

- no Excel repair dialog
- no corruption warning
- all formulas and styles survive

### Test 9 — repeated run

Run automation again on already-updated output.

Expected:

- it detects existing A08 and stops safely;
- it does not insert another A08 pair.

### Test 10 — forced failure

Intentionally make one header ambiguous in a test copy.

Expected:

- automation stops;
- original remains unchanged;
- no final output is published.

---

# 21. Acceptance criteria

The solution is production-ready for A08 only when all criteria below pass.

## File safety

- Original file byte content is unchanged.
- Final file is `.xlsb`.
- Final file opens normally in the user's Excel/NASCA environment.
- Excel does not show a corruption/repair warning.

## Layout

- `VD Total` has one new `A08 | %` pair in August.
- `MX Total` has one new `A08 | %` pair in August.
- `DA Total` has one new `A08 | %` pair in August.
- `Total PL` has one new `A08 | %` pair in August.
- New pair is before September.
- August merged header/formatting is correct.

## Logic

- Business A08 formulas are cloned from the correct August T08 logic.
- A08 formulas use the Actual version where appropriate.
- Percentage formulas remain formulas and use correct A08 references.
- `Total PL A08` reconciles to `VD Total + MX Total + DA Total` for every mapped row within tolerance.

## Preservation

- T08 and S08 are unchanged apart from Excel's legitimate reference adjustment caused by column insertion.
- no unintended sheets are changed.
- existing external link sources are preserved.
- pivots are not refreshed unexpectedly.
- workbook calculation completes.
- existing control/checker results do not regress.

## Repeatability

- Second run detects existing A08 and performs no destructive action.
- logs are sufficient to reproduce what happened.

---

# 22. Important implementation philosophy for the coding agent

This workbook has roughly one million formula cells. Do not try to understand or rewrite one million formulas in Python.

The safe solution is to let Excel preserve its own model and perform a very small number of **native structural operations**:

1. discover the exact August blocks;
2. insert two columns in four sheets;
3. native-copy the correct existing formula pair;
4. change the version criterion only where necessary;
5. create `Total PL` A08 from the already-created business A08 totals;
6. calculate once;
7. reconcile everything;
8. save/reopen/validate.

Treat Excel as the calculation/file-format engine and Python as the **orchestrator and safety controller**.

---

# 23. Phase boundaries

## Phase 1 — A08 only

Build and prove the exact August use case described above.

Do not generalize prematurely.

## Phase 2 — generic January-September Actual insertion

Only after A08 is proven, parameterize:

- `A01` through `A09`
- source `T01` through `T09`
- same semantic month-block discovery

## Phase 3 — October-December

Do not implement until the business owner explicitly defines the Actual naming convention for October, November and December.

Target naming is already defined as:

- `T0A`
- `T0B`
- `T0C`

Actual naming still requires confirmation.

## Phase 4 — source-data/pivot automation

Only later, if needed, automate:

- DB import
- Guide mapping checks
- pivot refresh
- rawdata refresh
- Summary/executive report propagation

That is a different risk level and must not be mixed into the first A08 production release.

---

# 24. Coding-agent execution instruction

When given this plan, the coding agent should follow this sequence:

1. Inspect the supplied real workbook in **read-only / dry-run mode first**.
2. Implement the project structure and COM session layer.
3. Implement semantic August-block discovery and print the discovered coordinates.
4. Do not write anything until all four target sheets have a unique valid match.
5. Implement safe `SaveCopyAs` working-copy transaction.
6. Implement one business sheet at a time: VD, then MX, then DA.
7. Validate each one before proceeding.
8. Implement Total PL only after the three business A08 columns are valid.
9. Implement full save/reopen validation.
10. Provide:
    - source code
    - requirements
    - one-click batch launcher
    - README for a non-technical user
    - dry-run command
    - execute command
    - automated test suite
    - example run report
11. Run all possible tests on a **copy** of the supplied workbook.
12. Never claim success unless the final saved workbook reopens and all acceptance checks pass.

If workbook-specific behavior differs from this plan, the agent must **inspect and report the difference instead of improvising a destructive workaround**.

---

# 25. Definition of done

The task is done only when the user can take the original protected P&L workbook on the Windows PC, run one local command/double-click launcher, and receive a separate validated `.xlsb` file where:

```text
VD Total : August -> T08 | % | S08 | % | A08 | %
MX Total : August -> T08 | % | S08 | % | A08 | %
DA Total : August -> T08 | % | S08 | % | A08 | %
Total PL : August -> T08 | % | S08 | % | A08 | %
```

and:

```text
Total PL A08 = VD Total A08 + MX Total A08 + DA Total A08
```

with the original workbook untouched, no corruption warning, formulas preserved, and a validation report proving the result.
