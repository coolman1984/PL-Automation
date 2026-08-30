# Project Status

**Last updated:** 2026-08-30  
**Milestone:** M4 universal Excel tool set complete in code; Windows re-acceptance pending
**Release state:** **NOT PRODUCTION READY** until the corrected mutation paths
pass the real Windows/Excel acceptance matrix.

## Verification

- 166 automated tests pass; 4 Windows/Excel integration tests are gated and
  skip cleanly outside the authorized Windows environment.
- The earlier real `read_range` acceptance remains valid: source SHA-256 was
  unchanged and no orphaned Excel process remained.
- Static checks pass: `pyflakes`, `compileall`, and `git diff --check`.
- The former defects D1-D16 now have code fixes or corrected test semantics.
  Real Excel re-acceptance is still required for calculation, PivotTable,
  format-copy, SaveCopyAs fidelity, and failure-injection paths.

## Completed

- [x] Dependency-free file/protection probe and safe engine routing.
- [x] Exact backup, SHA-256 evidence, workbook inventory, and full style snapshot.
- [x] Read-only Excel COM range adapter and real read acceptance harness.
- [x] Transaction journal, per-source lock, stale-lock recovery, and safe resume.
- [x] Generic dry-run and working-copy coordinator.
- [x] Bounded mutation implementations: clear, write, copy, formula fill,
  column insertion, and targeted PivotTable source update.
- [x] Advanced declared tools: exact formula matrices, range formatting, row
  insertion, sheet/table/filter/validation/comment/hyperlink/chart/name
  management, targeted connection refresh, calculation, and validation.
- [x] All 28 callable capabilities have machine-readable JSON contracts;
  `publish_workbook` and `restore_backup` remain deliberately locked.
- [x] Coordinator remediation: approved-plan digest, transaction match,
  revalidation at execution, explicit mutation intent, real-change check, and
  catch-all `FAILED_SAFE` handling.
- [x] Range remediation: exact payload/target shape equality, working-copy-only
  writes, finite single-area ranges, and limits before expensive reads.
- [x] Excel remediation: forced worksheet calculation before formula-error
  checks, multi-area error counting, fail-closed validation errors, table-name
  PivotTable resolution, correct format-copy enum, and clipboard-state cleanup.
- [x] Fidelity remediation: unsaved attached workbooks are refused; source and
  working-copy structure are compared immediately after `SaveCopyAs`; generic
  Excel extensions no longer fail after the copy was already created.
- [x] Documentation and catalogue now require explicit approval for every
  available mutation capability.

## Pending acceptance — blocks production

- [ ] Re-run the coordinator failure matrix on Windows Excel and confirm every
  injected failure persists `FAILED_SAFE` with no published output.
- [ ] Re-test `write_range` with a 1x1 payload against a larger target and prove
  that Excel is never called.
- [ ] Re-test `insert_columns` with real formulas while calculation starts in
  manual mode and prove new formula errors block publication.
- [ ] Re-test PivotTables backed by both worksheet ranges and Excel Tables,
  including unresolvable and shared-cache cases.
- [ ] Re-test format-only copy and prove `CutCopyMode` is cleared.
- [ ] Exercise every new advanced tool against real desktop Excel, including
  object-existence and before-fingerprint failures, calculation timeouts, and
  targeted connection refresh (never `RefreshAll`).
- [ ] Run two consecutive production-like transactions with unchanged source
  hashes, valid outputs, no orphaned Excel processes, and stable manifests.

## SAP recipe

Task 12 remains blocked until both the Windows re-acceptance above passes and
the user supplies the exact SAP export, target sheets, source/header ranges,
PivotTable names, and Target/Actual/Forecast column layout. These values must
not be guessed.

## Next single task

Run the Windows/Excel remediation acceptance matrix. Do not start the SAP
recipe or claim production readiness before every critical gate passes.
