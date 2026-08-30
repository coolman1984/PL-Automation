# Project Status

**Last updated:** 2026-08-30  
**Milestone:** M1 — agent-ready foundation complete  
**Overall state:** agent onboarding and contract foundation are implemented;
production generic write tools are still locked

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

## In progress

- [x] Excel COM adapter implementing the engine contract (unit-tested with COM doubles).
- [ ] First universal read-only range tool accepted against real Windows Excel.

## Planned and locked

- [ ] General range read/write and formula tools.
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

Implement the read-only Excel COM adapter and `read_range` tool, with a gated
Windows integration test. Keep all mutation tools locked until the adapter can
inspect exact targets and return JSON-safe evidence.
