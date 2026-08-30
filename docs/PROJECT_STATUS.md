# Project Status

**Last updated:** 2026-08-30  
**Milestone:** M2 — read-only Excel engine in progress  
**Overall state:** agent onboarding and read-only foundation are implemented;
production generic write tools are still locked

**Current verification:** 84 tests passed; two Windows/Excel integration tests
are gated until an authorized Windows machine is available.

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

## In progress

- [x] Excel COM adapter implementing the engine contract (unit-tested with COM doubles).
- [ ] First universal read-only range tool accepted against real Windows Excel.

## Planned and locked

- [ ] General range write and formula tools (read-only range is implemented but
  real Windows acceptance is pending).
- [ ] Row/column and worksheet structure tools.
- [ ] Formatting, tables, filters, validation, comments, and hyperlinks.
- [ ] Charts, pivots, names, connections, shapes, and controlled refresh.
- [ ] Deterministic plan executor and recipe composition.
- [ ] Large-workbook chunking, progress, cancellation, and benchmarks.
- [ ] Real Windows acceptance matrix for protected and advanced workbooks.

## Known blockers

- Real Excel COM acceptance requires Windows desktop Excel and authorized
  access to representative workbooks.
- NASCA/DRM must be handled through the user's already-authorized Excel session.
- No library can promise lossless mutation of every proprietary Excel feature;
  unsupported feature detection must stop the run.

## Next single task

Run the read-only `read_range` acceptance test on an authorized Windows Excel
machine. After it passes, start the generic write transaction layer; keep all
mutation tools locked until backup, rollback, and reopen-validation gates pass.
