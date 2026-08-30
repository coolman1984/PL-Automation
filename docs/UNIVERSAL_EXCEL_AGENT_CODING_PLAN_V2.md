# Universal Excel Agent — Coding Execution Plan V2

**Status:** Approved implementation roadmap  
**Date:** 2026-08-30  
**Current milestone:** Close M2, then build the generic mutation transaction  
**Primary platform:** Windows 10/11 with desktop Microsoft Excel  
**Primary fidelity engine:** Microsoft Excel COM  
**Priority:** Safety, fidelity, correctness, performance

V2 supersedes V1 for implementation order and acceptance gates. V1 remains in
the repository as design history. A coding agent must read `AGENTS.md`,
`docs/PROJECT_STATUS.md`, this plan, and only the files named by the active task.

## 1. Outcome

Build the next safe vertical slices of the universal Excel agent and deliver a
first reusable business recipe that can:

1. inspect a user-selected SAP export and target workbook;
2. create and verify backup and snapshot evidence;
3. refresh the `DB File` worksheet from the SAP export;
4. fill the helper formulas in columns `A:C` through the new data row count;
5. update and refresh PivotTables on `PV` using the resolved data source;
6. insert a configured Target / Actual / Forecast column group;
7. save, close, reopen, validate, and publish a separate output;
8. leave the original workbook byte-for-byte unchanged on success or failure.

The SAP recipe is not allowed to bypass the universal tools or call COM
directly. It remains locked until every dependency and acceptance gate in this
plan passes.

## 2. Verified baseline and current risks

### Implemented

- agent entry pack and project-status commands;
- typed request/result/error/plan contracts;
- capability catalogue with available versus planned tools;
- file signature, format, complexity, and protection probe;
- exact byte backup and SHA-256 manifest;
- workbook inventory/full snapshot;
- fake engine and contract-test harness;
- read-only Excel COM range adapter and engine routing;
- reusable transaction state enumeration and JSON journal;
- guarded P&L A08 transaction and publication workflow.

### Evidence observed on Windows

- 84 unit and contract tests passed before this V2 planning change;
- the real `read_range` integration test read `VD Total!A1:C10` successfully;
- the source XLSB SHA-256 remained
  `353EE458286A582A8D3308171CA3E1370C755BDD49FB937F2594C4B0C5F7ED69`;
- the Excel process count was identical before and after the read-only test.

### Stabilization items that must be closed first

1. `workflow.py` imported `candidate_has_business_lineage` from the wrong
   module. The focused fix is to import it from `total_pl_updater.py` and keep a
   regression test. Avoid line-ending or formatting churn.
2. The leading `★` in the real workbook path can raise a `UnicodeEncodeError`
   under a CP1252 console. CLI output must be UTF-8 safe.
3. On this Windows/Excel build, Python `faulthandler` reports a handled
   first-chance `0x80010108` while Excel COM proxies disconnect during normal
   shutdown. The acceptance harness must distinguish this diagnostic from a
   real crash by checking process exit code, source hash, and orphan Excel PIDs.
   Do not hide production exceptions inside application code.
4. Generic write branches exist behind the engine interface, but no mutation
   tool may be unlocked until the generic coordinator enforces backup, working
   copy, journal, reopen validation, and publication rules.

## 3. Non-negotiable safety rules

1. Never save, edit, rename, move, or delete the source workbook or SAP export.
2. Every mutation requires an explicit source path, output intent, transaction
   ID, verified backup, and source hash.
3. All writes target an Excel-created working copy.
4. Every sheet, range, table, PivotTable, and insertion anchor must be resolved
   exactly. Ambiguity stops the run.
5. Destructive tools require a before fingerprint and expected target size.
6. The SAP import must validate row and column counts before clearing old data.
7. No clipboard, `Select`, `Activate`, keystrokes, mouse automation, or active
   sheet/cell assumptions.
8. Bulk rectangular COM assignments or `Range.Copy(Destination=...)` are
   required; cell-by-cell COM loops are forbidden.
9. Pivot refresh is opt-in and restricted to the resolved PivotTables on `PV`.
   Do not call `RefreshAll` unless a separate approved plan requires it.
10. The result must be saved, closed, reopened, structurally validated, and
    compared with the plan before publication.
11. Failure moves the transaction to `FAILED_SAFE`; a failed working copy is
    retained as diagnostic evidence and never published as successful.
12. Existing VBA, links, charts, pivots, shapes, names, connections, protection,
    and unrelated formulas/styles must remain unchanged.

## 4. Architecture decisions

- Keep schema version `1.0` and the current public tool names during this
  milestone. Do not perform a tool-renaming migration while adding writes.
- Extend existing modules behind compatibility imports; do not rewrite the P&L
  workflow as part of the SAP recipe.
- The coordinator owns transaction sequencing. Tools perform one bounded Excel
  operation. Recipes compose tools and business validation only.
- Use Excel COM for the target XLSB and PivotTables. A fast OOXML engine is not
  eligible for this recipe.
- Introduce one tool at a time and change its catalogue status only after its
  complete release checklist passes.

Dependency order:

```text
M2 stabilization and real read acceptance
    -> M3 generic transaction coordinator
        -> M4 bounded mutation tools
            -> M5 SAP DB/PV recipe
                -> M6 performance and production acceptance
```

## 5. SAP DB/PV recipe contract

### Proposed recipe name

`sap_db_refresh_and_pv_update`

### Required plan inputs

```json
{
  "target_workbook": "absolute path to the P&L working source",
  "sap_export": "absolute path to the selected SAP download",
  "db_sheet": "DB File",
  "pv_sheet": "PV",
  "formula_template": "A21:C21",
  "formula_fill_start_row": 22,
  "import_first_column": "E",
  "import_last_column": "JJ",
  "import_start_row": 22,
  "clear_start_row": null,
  "sap_source_sheet": null,
  "sap_header_rows": null,
  "pivot_names": [],
  "version_columns": {
    "destination_sheet": null,
    "insert_after": null,
    "order": ["Target", "Actual", "Forecast"],
    "headers": [],
    "formula_templates": []
  },
  "dry_run": true
}
```

`null` values are unresolved requirements, not defaults. Execution must stop
until the dry-run or the user resolves them.

### Workbook facts the dry-run must resolve

1. Whether the exact target sheet is `DB File`, `DB`, or another name. Do not
   choose by similarity if more than one candidate exists.
2. Whether row 21 contains headers/templates and must be preserved. The safe
   proposed default is to clear `E22:JJ<old_last_row>`. Clearing from row 21 is
   allowed only after capturing row 21 and proving how it will be rebuilt.
3. Which SAP worksheet contains the export and whether its first row is a
   header. Do not select the newest downloaded file automatically; use the
   user-selected exact file.
4. The SAP source range and shape. `E:JJ` is 266 destination columns. The source
   width must equal 266 unless an explicit mapping is approved.
5. The new data last row. If `N` SAP rows are imported at row 22, the expected
   last row is `21 + N`.
6. The existing PivotTable names and current source definitions on `PV`.
7. The destination sheet, anchor, labels, formulas, formats, and percentage
   partners—if any—for the new Target / Actual / Forecast columns.

### Required operation order

```text
probe both files
-> verify source hashes
-> backup and snapshot target
-> inspect SAP shape and DB/PV structure
-> validate dry-run plan and request approval
-> create Excel working copy
-> clear the approved old DB range
-> copy SAP data to DB File!E22:JJ...
-> fill A21:C21 formulas down through the new last row
-> change only the approved PV PivotTable source ranges
-> refresh only those PivotTables
-> insert/configure Target, Actual, Forecast columns
-> validate target and non-target workbook facts
-> save/close/reopen/revalidate
-> publish separate output
```

### Import behavior

- Clear contents only; do not delete worksheet rows unless a separately approved
  structural plan requires row deletion.
- The clear range must be explicit and bounded by the old proven last row.
- Copy SAP values and required formats in bulk. Prefer
  `source_range.Copy(Destination=target_range)` when full Excel fidelity is
  required; otherwise assign `Value2` plus an explicit approved format copy.
- Never clear old data before source shape, destination shape, and disk space
  checks pass.
- If the SAP import contains fewer columns, extra columns, duplicate headers,
  or no data rows, stop without editing the working copy.

### Formula behavior for `A:C`

- Treat `A21:C21` as the formula template only after confirming all three cells
  contain the expected formulas.
- Preserve row 21.
- Fill formulas into `A22:C<new_last_row>` using Excel formula propagation so
  relative, absolute, structured, and cross-sheet references retain Excel
  semantics.
- Validate the first, middle, and last filled rows and verify no formulas extend
  beyond the imported data.

### Pivot behavior for `PV`

- Inventory PivotTable name, cache identity, source type, source address, and
  refresh settings before change.
- Preserve the existing source sheet and top-left source cell unless the plan
  explicitly changes them. Normally update only the final row and, if proven,
  the final column through `JJ`.
- If a PivotTable uses an external connection, Data Model, OLAP source, or
  shared cache whose safe update is unsupported, stop and report it.
- Validate the source after refresh, PivotTable count, cache relationships,
  and at least one configured control total/item count.

### Target / Actual / Forecast columns

This step is included but cannot be executed from the current wording alone.
The plan must resolve:

- destination worksheet(s);
- exact insertion anchor and order;
- period/version headers such as `Txx`, `Axx`, and `Sxx`;
- whether each amount column has a paired percentage column;
- source formula templates and business reconciliation rules;
- formatting source and merged-header behavior.

After resolution, use `insert_columns`, `copy_range`, `set_formula`, and
`format_range` through the coordinator. Never invent formulas or infer a month
from the current date.

## 6. Implementation tasks

### Task 1 — Stabilize imports and repository hygiene

**Description:** Keep the lineage helper in `total_pl_updater.py`, correct the
workflow import, and add an import regression test without formatting unrelated
lines.

**Acceptance criteria:**

- `src.workflow` imports successfully;
- the helper identity matches `src.total_pl_updater`;
- `git diff --check` is clean for files touched by this task.

**Verification:**

```text
python -m pytest tests/unit/test_workflow_imports.py -q
python -m pytest tests/unit tests/contract -q
```

**Files:** `src/workflow.py`, `tests/unit/test_workflow_imports.py`  
**Scope:** Small  
**Dependencies:** None

### Task 2 — Make CLI output Unicode-safe

**Description:** Ensure commands can print workbook paths containing `★` under
Windows console encodings without changing path identity.

**Acceptance criteria:**

- `--probe-only` succeeds for a path containing `★`;
- JSON output remains UTF-8 and machine-readable;
- no filename sanitization is used for file access.

**Verification:** targeted CLI test plus `python -X utf8 app.py --file PATH --probe-only`.

**Files:** `app.py`, one focused CLI test  
**Scope:** Small  
**Dependencies:** Task 1

### Task 3 — Create a deterministic read-only COM acceptance harness

**Description:** Add a PowerShell harness that runs only the gated read-range
test and records source hashes, exact parameters, exit code, and Excel PIDs.

**Acceptance criteria:**

- nonzero pytest exit fails the harness;
- source hash before/after must match;
- no new isolated Excel PID remains;
- a JSON acceptance report is retained;
- handled `faulthandler` diagnostics are not treated as proof of a crash when
  process exit is zero, but production code does not suppress exceptions.

**Files:** `tools/run_read_range_acceptance.ps1`, one harness test, integration README  
**Scope:** Medium  
**Dependencies:** Tasks 1–2

### Checkpoint A — Close M2

- all unit and contract tests pass;
- integration tests collect cleanly when gated off;
- real `read_range` acceptance passes through the harness;
- source hash is unchanged and no Excel process is orphaned;
- update `PROJECT_STATUS.md` and keep all mutation tools locked.

### Task 4 — Harden the transaction journal

**Description:** Extend `TransactionContext` with per-run workspace locking,
source/backup/working/output hashes, automatic atomic journal writes, and safe
resume boundaries.

**Acceptance criteria:** sequence numbers cannot repeat; invalid transitions
fail closed; interrupted journal writes do not corrupt the last valid state.

**Files:** `src/transaction_state.py`, one unit-test module  
**Scope:** Medium  
**Dependencies:** Checkpoint A

### Task 5 — Implement the generic coordinator dry-run path

**Description:** Coordinate probe, backup, snapshot, deterministic plan
validation, risk approval, and `FAILED_SAFE` handling without mutation.

**Acceptance criteria:** no working copy is edited during dry-run; missing or
ambiguous targets stop before approval; every state has evidence.

**Files:** `src/core/coordinator.py`, `src/plan_validation.py`, focused tests  
**Scope:** Medium  
**Dependencies:** Task 4

### Task 6 — Implement working-copy, reopen, validation, and publication

**Description:** Add the mutating half of the coordinator using existing backup,
SaveCopyAs, validation, and publication components.

**Acceptance criteria:** every injected failure transitions to `FAILED_SAFE`;
the source hash never changes; publication is possible only from `VALIDATED`.

**Files:** coordinator, transaction compatibility adapter, failure-injection tests  
**Scope:** Medium  
**Dependencies:** Task 5

### Checkpoint B — Generic transaction gate

- failure injection passes at every state transition;
- closed/reopened working copy is validated before publication;
- no generic mutation tool is available yet;
- P&L tests remain green.

### Task 7 — Release bounded `clear_range`

Implement clear-contents only with exact range, expected cell count, before
fingerprint, dry-run, backup requirement, and rollback tests. Row/column deletion
is out of scope.

**Files:** one range-tool module, registry/catalogue, focused tests  
**Scope:** Medium  
**Dependencies:** Checkpoint B

### Task 8 — Release bulk `write_range` and `copy_range`

Require rectangular shape equality, chunking thresholds, before/after evidence,
and explicit cross-workbook source identity. Test values, formulas, formats,
empty cells, dates, errors, and large blocks.

**Files:** range-tool module, executor adapter, focused tests  
**Scope:** Medium  
**Dependencies:** Task 7

### Task 9 — Release formula fill-down

Add a bounded formula propagation operation using Excel semantics. Validate
Formula versus Formula2, template fingerprints, target row count, and no fill
beyond the imported rows.

**Files:** formula-tool module, catalogue, focused tests  
**Scope:** Medium  
**Dependencies:** Task 8

### Task 10 — Release controlled column insertion and formatting copy

Insert an exact count at an exact anchor on a working copy; copy approved header,
formula, width, merge, and style patterns; validate shifted references.

**Files:** column/formatting tools, catalogue, focused tests  
**Scope:** Medium  
**Dependencies:** Task 9

### Task 11 — Release PivotTable source update and targeted refresh

Support only worksheet/table sources proven by integration tests. Preserve
shared caches unless the plan explicitly replaces them. External/Data Model
sources remain locked.

**Files:** pivot tool, Excel COM adapter, catalogue, focused tests  
**Scope:** Medium  
**Dependencies:** Task 8

### Checkpoint C — Core SAP dependencies

Every tool in Tasks 7–11 must pass its complete release checklist: typed
contract, dry-run, approval rule, evidence, validation, fake-engine tests, real
Excel integration, failure rollback, docs, and catalogue consistency.

### Task 12 — Build SAP recipe preflight and dry-run plan

Resolve both workbooks, source sheet/header rows, 266-column shape, DB old/new
last rows, formula template, PivotTables, and version-column configuration.
Return unresolved questions instead of guessing.

**Files:** `src/recipes/sap_db_refresh.py`, recipe models/config, focused tests  
**Scope:** Medium  
**Dependencies:** Checkpoint C

### Task 13 — Execute DB clear, SAP import, and formula fill

Compose only released tools. Validate imported row count, `E:JJ` width, formula
coverage in `A:C`, and untouched neighboring cells.

**Files:** SAP recipe, recipe validators, acceptance fixtures  
**Scope:** Medium  
**Dependencies:** Task 12

### Task 14 — Execute PV source update and targeted refresh

Update only approved PivotTables, wait for calculation/refresh completion, and
validate source address, cache identity, counts, and configured business totals.

**Files:** SAP recipe, pivot validators, acceptance tests  
**Scope:** Medium  
**Dependencies:** Task 13

### Task 15 — Insert Target / Actual / Forecast columns

Implement this recipe step only after all unresolved column details are present
in the approved plan. Use existing analogous columns as templates and validate
headers, formulas, formats, merges, widths, and reconciliations.

**Files:** SAP recipe, column-group configuration, acceptance tests  
**Scope:** Medium  
**Dependencies:** Tasks 10 and 14

### Task 16 — Reopen validation, report, and publication

Compare source and output structure, imported data, formulas, PivotTables,
version columns, and untouched workbook objects. Produce a readable report and
complete JSON journal before publishing a separate output.

**Files:** recipe validators, reporting adapter, acceptance tests  
**Scope:** Medium  
**Dependencies:** Task 15

## 7. Required SAP recipe acceptance matrix

| Case | Required evidence |
|---|---|
| Empty SAP export | Stop before clearing DB data |
| Wrong source width | Stop; expected 266 columns |
| Duplicate/missing headers | Stop with exact header evidence |
| Smaller new dataset | Old trailing rows are cleared; formulas stop at new end |
| Larger new dataset | Bulk import and formula fill reach exact new end |
| Missing `DB File` or `PV` | Stop; never choose a similar sheet silently |
| Multiple PivotTables | Update only approved names |
| External/Data Model pivot | Stop as unsupported until released |
| Failure after clear/import/fill/pivot | No publication; source unchanged |
| Successful run | Save, close, reopen, validate, publish separate output |

Use sanitized fixtures for automated tests and an authorized disposable copy
for real Excel acceptance. Never commit company workbooks or SAP exports.

## 8. Performance requirements

- Read/write `E:JJ` in rectangular blocks; no per-cell COM loop.
- Record elapsed time, rows, columns, cells touched, and peak process memory.
- Add configurable chunking only after measuring the real payload.
- Cancellation is allowed only between safe chunks and results in
  `FAILED_SAFE`, never a partial publication.
- Define supported row limits from benchmark evidence; do not claim unlimited
  workbook size.

## 9. Coding-agent operating protocol

For every task:

1. Read this task, its dependencies, and only the named files.
2. Inspect the current worktree and preserve unrelated user changes.
3. Add a failing focused test first.
4. Implement the smallest complete slice.
5. Run focused tests, then unit/contract tests.
6. Run real Excel tests only when their environment gates and workbook authority
   are satisfied.
7. Review the diff for secrets, workbook binaries, generated output, and
   unrelated formatting churn.
8. Update `PROJECT_STATUS.md`, `FILE_MAP.md`, catalogue, and recipe docs in the
   same capability change.
9. Keep a tool locked unless every release checklist item is proven.
10. Stop at the current task/checkpoint; do not implement later phases early.

## 10. Definition of V2 success

V2 is complete when a fresh coding agent can run the SAP DB/PV recipe on an
authorized disposable workbook and prove:

- the original target and SAP export hashes never changed;
- the approved old DB range was cleared and exactly the approved SAP rows were
  imported into `E:JJ`;
- formulas in `A:C` cover exactly the imported rows;
- approved PivotTables on `PV` use the validated new source and refresh cleanly;
- Target / Actual / Forecast columns match the approved structure and formulas;
- unrelated workbook content and advanced Excel objects are preserved;
- success is published only after close/reopen validation;
- every failure produces useful evidence and no published partial output.

