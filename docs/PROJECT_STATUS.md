# Project Status

**Last updated:** 2026-08-30  
**Milestone:** M4 — Checkpoint B closed; bounded mutation tools are being
released one at a time (`clear_range`, `write_range`, `copy_range` done)  
**Overall state:** the generic coordinator (dry-run + mutating halves) is
implemented and proven against real Windows Excel end-to-end, including a
real cell mutation reaching a published output while the source stayed
byte-identical; remaining mutation tools stay locked until each passes its
own release checklist

**Current verification:** 142 tests passed, 4 gated Windows/Excel integration
tests skip cleanly by default. **Three high-severity defects are open — see
"Known defects" below; Checkpoints B and C are NOT fully met.** The real
`read_range` acceptance harness
(`tools/run_read_range_acceptance.ps1`) was run against the actual source
workbook on an authorized Windows machine: source SHA-256 unchanged, no
orphaned Excel process, pytest exit 0. Report retained at
`work/acceptance/read_range_acceptance_20260830_144209.json`.

## Completed

- [x] Dependency-free Excel container/protection probe.
- [x] Fast-path candidate decision with safe fallback to Excel COM.
- [x] Byte-for-byte backup and SHA-256 manifest.
- [x] Inventory/full workbook snapshot with reusable style records.
- [x] Guarded P&L A08 COM workflow, validation, reporting, and idempotency.
- [x] Offline packaging and private runtime repair scripts.
- [x] Machine-readable initial tool catalogue.
- [x] Universal Excel coding plan.
- [x] Agent entry instructions and safety contract.
- [x] Typed request/result/error/plan contracts.
- [x] Fake engine and contract-test harness.
- [x] Safe executor for declared probe/backup tools.
- [x] Agent start, status, and tool-description commands.
- [x] Generated catalogue plus catalogue schema.
- [x] Read-only `read_range` contract and Excel COM adapter code.
- [x] Excel COM adapter implementing the engine contract (unit-tested with COM doubles).
- [x] Deterministic read-only COM acceptance harness (Task 3).
- [x] First universal read-only range tool accepted against real Windows Excel.
- [x] Checkpoint A — M2 closed: unit/contract tests pass, integration tests
  gate cleanly, real `read_range` acceptance passed with unchanged source hash
  and no orphaned Excel process.

- [x] Task 4: hardened transaction journal — per-source workspace locking with
  stale-lock reclaim, source/backup/working/output hash stages, automatic
  atomic journal writes on every event, and safe resume that fails closed for
  in-flight states (`EXECUTING`, `SAVED`, `REOPENED`). Fixed a real flake
  found while stress-running the suite during Task 8: `write_journal`'s
  `Path.replace` can transiently raise `PermissionError` on Windows (e.g.
  antivirus scanning a just-written temp file) when events fire in rapid
  succession; it now retries briefly before giving up.

- [x] Task 5: generic coordinator dry-run path (`src/core/coordinator.py`)
  composes only declared tools (`inspect_file`, `create_backup`,
  `snapshot_workbook`) through `execute_tool`, validates the plan, and stops
  at `APPROVED` without ever creating a working copy. Verified against real
  Windows Excel with a disposable workbook: state reached `approved`, source
  hash unchanged, no orphaned Excel process, full evidence journal.

- [x] Task 6: mutating coordinator half (`run_execute` in
  `src/core/coordinator.py`) — working copy via `SaveCopyAs`, execute plan
  steps through the same `execute_tool` an agent would use, pre-save and
  post-reopen generic preservation checks, publish only from `VALIDATED`.
  Every failure path (not-approved, source changed since backup, failed step,
  preservation violation, reopen/publish error) transitions to `FAILED_SAFE`
  and leaves the source byte-identical. Verified with failure-injection unit
  tests and twice against real Windows Excel: once end-to-end to a published
  output, once proving a plan naming a still-locked mutating tool
  (`write_range`) is refused before any edit.
- [~] Checkpoint B — **NOT met.** Originally recorded as closed, but a
  later review disproved its central criterion: `run_execute` is a bare
  `try/finally` with no `except`, so an unhandled failure (e.g. `Save()`
  raising) escapes with the transaction left in `EXECUTING` both in memory
  and on disk, never reaching `FAILED_SAFE`. See defects D1/D2 below.

- [x] Task 7: released `clear_range` — the first unlocked mutation tool.
  Clear-contents only (`Range.ClearContents`, never `.Clear()` or a row/column
  delete), requires an explicit working-copy target (never `source`),
  requires `arguments.expected_cell_count` to match the resolved range
  exactly before touching anything, returns the pre-clear values as
  `before_evidence` for rollback, and stays a dry-run no-op unless
  `dry_run=False`. Verified with focused unit tests (including a rollback
  proof using the recorded before-evidence) and against real Windows Excel:
  values cleared, cell formatting (fill color) preserved, confirming
  `ClearContents` semantics.

- [x] Task 8: released bulk `write_range` and `copy_range`. Both now
  capture pre-write/pre-copy values as `before_evidence` (rollback), enforce
  a `_MAX_CELLS_PER_RANGE_OPERATION` (200,000) cap in place of real chunking
  (not yet benchmarked), and `copy_range` additionally requires source/
  destination shape equality and explicitly refuses a source naming a
  workbook this engine session didn't open (`cross_workbook_copy_unsupported`)
  rather than guessing. Verified with focused unit tests and, against real
  Windows Excel, a full dry-run -> execute -> publish run that actually
  wrote a new cell value through to the published output while the source
  stayed byte-identical.

- [x] Task 9: released `fill_formula_down`. Requires an explicit single-row
  template target plus `expected_template_formulas` (exact fingerprint
  check before touching anything) and `expected_target_row_count`; the
  target must be column-aligned and contiguously directly below the
  template. Uses Excel's native `Range.FillDown` (via a combined
  template+target range) so relative/absolute/structured references get
  real Excel semantics rather than a hand-rolled reference rewriter, and
  re-reads the template row after the fill as a defense-in-depth check that
  the source row was never itself overwritten. Verified with focused unit
  tests and against real Windows Excel: a relative reference shifted
  correctly per row, an absolute reference stayed fixed, and the template
  row's formula was provably untouched.

- [x] Task 10: released `insert_columns` — exact anchor (verified against a
  required `expected_anchor_column` before touching anything) and exact
  count, using Excel's own `Insert(Shift, CopyOrigin)` so neighboring width/
  merge/style inherit through Excel's native semantics; "formatting copy"
  of header/formula content into the new columns is left to already-released
  `copy_range`/`write_range`/`fill_formula_down` composed by a future recipe,
  matching the plan's own architecture decision that tools stay bounded and
  recipes compose them. Adds a before/after formula-error scan
  (`count_formula_errors`, via `SpecialCells(xlFormulas, xlErrors)`) so a
  broken reference anywhere on the sheet fails the whole tool call (and
  therefore the whole transaction, discarding the disposable working copy)
  rather than silently publishing.
  **Bug found and fixed while verifying this against real Windows Excel:**
  pywin32's late-bound dynamic COM dispatch does not accept named keyword
  arguments on `Range.Resize` (a parameterized *property*, unlike a true
  method such as `Insert`), so `.Resize(ColumnSize=2)` always raised
  `TypeError: ... unexpected keyword argument 'ColumnSize'`. This exact
  idiom was already present, unfixed, in the "completed" P&L A08 production
  workflow (`business_sheet_updater.insert_two_columns`,
  `total_pl_updater`'s Total PL column insert) — meaning a real execute run
  of the monthly close would have failed at the column-insert step the
  first time anyone actually ran it against real Excel. Both call sites now
  span the target range via `Range(col1, col2)` instead of `Resize`, and
  were reverified against real Excel.

- [x] Task 11: released `update_pivot_source`. Requires an exact
  `target.sheet`/`target.object_name` PivotTable, an `expected_current_source`
  fingerprint check, and refuses external/Data Model sources
  (`unsupported_pivot_source`) and any PivotTable whose cache is shared with
  another PivotTable unless `allow_shared_cache_replacement=true` is set
  explicitly. Uses `ChangePivotCache` + a targeted `RefreshTable()` (never
  `RefreshAll`).
  **Two real bugs found and fixed while verifying this against real
  Windows Excel with actual PivotTables:** (1) `PivotCache.SourceData`
  always reports addresses in Excel's internal R1C1 notation regardless of
  what notation was used to set the source, so naive string-equality
  checks against a caller's A1-style `expected_current_source`/`new_source`
  would spuriously fail on every real call; (2) a caller round-tripping
  the tool's own previously-reported R1C1 string back as
  `expected_current_source` would also fail, because `Application.Range()`
  cannot parse R1C1-style strings. Fixed by resolving both notations to a
  canonical `(sheet, first_row, first_col, last_row, last_col)` tuple
  (`resolve_source_bounds`, with an R1C1 regex parser as a fast path before
  falling back to `Application.Range`) and comparing bounds, not raw
  strings. Reverified end-to-end against real Excel: a source update that
  added a new data row correctly flowed through to the PivotTable's
  displayed output and grand total, and the shared-cache guard correctly
  blocked/allowed based on the acknowledgement flag.
- [~] Checkpoint C — **NOT met.** Originally recorded as closed. The tools
  from Tasks 7–11 do each have a typed contract, a dry-run path, evidence
  and unit tests, but a subsequent review found that several of their
  release-checklist guards do not actually hold. `write_range` has no
  shape guard at all (D3), and the `update_pivot_source` and
  `insert_columns` validations fail open (D4/D5). The earlier "reverified
  against real Windows Excel" claim was too strong: each tool was smoke
  tested in isolation against a bare workbook, which is exactly the
  configuration in which D4 and D5 do not manifest. See "Known defects".

## In progress

- [ ] Task 12: build the SAP recipe preflight and dry-run plan. **This
  needs real input from the user first** — the plan explicitly forbids
  guessing: the exact SAP export file, the exact `DB File`/`PV` sheet
  names, the SAP source range/header rows, the existing PivotTable names on
  `PV`, and the Target/Actual/Forecast column layout are all currently
  `null`/unresolved.

## Planned and locked

- [ ] Row/column and worksheet structure tools.
- [ ] Formatting, tables, filters, validation, comments, and hyperlinks.
- [ ] Charts, pivots, names, connections, shapes, and controlled refresh.
- [ ] Deterministic plan executor and recipe composition.
- [ ] Large-workbook chunking, progress, cancellation, and benchmarks.
- [ ] Real Windows acceptance matrix for protected and advanced workbooks.

## Known defects

Found by review after Checkpoint C was (prematurely) recorded as closed.
Every item marked **[proven]** was reproduced by executing the real
production code; the rest are code-level findings not yet reproduced.
**No mutation tool should be used in production until D1–D6 are fixed.**

### Blocking — publication / data integrity

- **D1 [proven] — a dry-run-only plan is published as a successful update.**
  `ToolRequest.dry_run` defaults to `True`, and `plan_validation` only
  demands approval when `dry_run is False`. A plan of real mutating steps
  that simply omits `dry_run` validates without approval, every step
  returns `ok=True, changed=False`, preservation checks trivially pass, and
  `run_execute` publishes `<name>__UPDATED.xlsb` — verified byte-identical
  to the untouched source. `run_execute` checks only `result.ok`, never
  `dry_run`, `result.changed`, or `expected_effect`.
- **D2 [proven] — unhandled exceptions escape `run_execute` without
  reaching `FAILED_SAFE`.** No `except` clause; `sha256_file`,
  `fingerprint()`, `save_and_close()`, `record_hash`, `transition` and the
  `mkdir` calls are all unguarded. A `Save()` failure leaves the
  transaction at `EXECUTING` in memory *and in the on-disk journal*, with
  no `failed_safe` event. Same shape in `run_dry_run`.
- **D3 [proven] — `write_range` never checks the payload against the target
  range.** `_shape` proves the payload is rectangular but nothing compares
  it to the resolved target, and there is no `expected_*` argument.
  Verified on real Excel: a `[["TOTAL"]]` payload into `A1:E10` silently
  **broadcast across all 50 cells** while the result reported
  `shape={rows:1,columns:1}` and `cells_touched=1`. `FakeEngine` raises on
  the same call, so no unit test can see this.

### Blocking — guards that fail open

- **D4 [proven] — `insert_columns`' formula-error guard is inert in the
  production path.** The coordinator's `open_working_copy_for_edit` sets
  `Calculation = xlCalculationManual`; `SpecialCells` reads cached values,
  which never update. Verified: `errors_before=0, errors_after=0` while a
  real error existed (`errors=1` once recalculated). Additionally
  `count_formula_errors` returns `0` from a blanket `except`, and
  `Range.Count` on a multi-area result counts only the first area.
- **D5 [proven] — `update_pivot_source` is broken both ways for
  Table-backed pivots.** `SourceData` for a ListObject source is the table
  *name*, which the R1C1 parser cannot parse, so `source_bounds` is `None`.
  A *correct* caller is falsely rejected; a caller passing an unresolvable
  address also yields `None`, and `None != None` is False, so the guard
  passes vacuously — verified the pivot was then actually mutated. The
  post-mutation check has the identical vacuous comparison.
- **D6 — `write_range` and `copy_range` lack the working-copy-only target
  guard** that `clear_range`, `fill_formula_down` and `insert_columns` all
  have; `ExcelComEngine._resolve_target` also whitelists `"source"`.
  Currently mitigated only because the coordinator binds the engine to the
  working copy. `copy_range` additionally accepts `workbook_id="source"` as
  its *source* and silently reads the working copy instead.

### Serious — approval and fidelity

- **D7 [proven] — `step.tool` and `step.request.tool` may differ.**
  `validate_plan` resolves the spec from `step.tool` while the coordinator
  dispatches `step.request`. Declaring `step.tool="read_range"` with
  `request.tool="clear_range"` passes validation with no approval required
  and really clears cells.
- **D8 — `run_execute` never validates the plan it is given.** It accepts a
  `plan` independent of `dry_run.context`, with no `validate_plan` call and
  no transaction-id match, so an approved benign plan can be swapped for an
  unvalidated destructive one.
- **D9 — the source → working-copy transfer is never verified.** All three
  fingerprints come from the working copy, so if `SaveCopyAs` drops VBA,
  links or pivot caches, the loss is already in the baseline and every
  preservation check compares loss to loss and passes.
- **D10 — attach mode can copy unsaved in-memory edits.** No `Saved` check
  in `transaction_adapter` (the snapshot path in `tool_executor` has one).
  Both source-hash checks hash the on-disk file and so are blind to it.
- **D11 [proven] — the "format-agnostic" coordinator cannot complete on any
  non-`.xlsb` source.** `save_working_copy` hard-rejects a non-`.xlsb`
  working path *after* already writing the file, leaving an orphan.

### Lower

- **D12 — `copy_range(mode="formats")` is broken:** `PasteSpecial(Paste="formats")`
  passes a string where `XlPasteType` is an integer enum, and
  `source_range.Copy()` leaves `CutCopyMode` armed, so a later
  `Range.Insert` can insert the copied cells instead of blank ones.
- **D13 — multi-area / whole-column addresses defeat the cell-count
  guards.** `Range.Value2` returns only the first area while
  `ClearContents()` clears all of them, so `clear_range` can wipe more
  cells than it counted and than its rollback evidence covers.
- **D14 — cell-count caps are enforced on only 2 of 6 mutating tools**
  (`write_range`, `copy_range`), and even there only after the expensive
  read has already materialised the range.
- **D15 — dead check:** `coordinator.py` builds the context with
  `transaction_id=plan.transaction_id` then tests the two for inequality.
- **D16 — two unit tests encode non-Excel semantics.** `FakeEngine`'s
  `fill_formula_down` does not shift relative references and its
  `insert_columns` does not rewrite cross-sheet references, so the
  corresponding assertions describe behaviour that would be wrong in real
  Excel.

## Known blockers

- Real Excel COM acceptance requires Windows desktop Excel and authorized
  access to representative workbooks.
- NASCA/DRM must be handled through the user's already-authorized Excel session.
- No library can promise lossless mutation of every proprietary Excel feature;
  unsupported feature detection must stop the run.

## Next single task

**Do not start Task 12.** Fix the blocking defects D1–D6 above first and
re-earn Checkpoints B and C; the SAP recipe composes exactly the tools
those defects live in, so building on them now would bake the faults into
the recipe. Suggested order: D1 and D2 (coordinator: require `dry_run`
off for mutating steps, verify `changed`, add a catch-all that fails
safe), then D3/D6 (`write_range` shape + target guards), then D4/D5 (the
two fail-open validations).

After that, Task 12 (SAP recipe preflight and dry-run plan) still cannot
start until the user supplies the unresolved plan inputs: the exact SAP
export file, the exact `DB File`/`PV` sheet names, the SAP source range and
header-row convention, the existing PivotTable names on `PV`, and the
Target/Actual/Forecast column layout (destination sheet, insertion anchor,
headers, formula templates). Guessing any of these is explicitly against
the plan's rules. Note `Untitled.txt` in the repo root holds the user's
original hand-written statement of this requirement.
