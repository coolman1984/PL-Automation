# File Map

| Path | Responsibility | Status |
|---|---|---|
| `app.py` | Thin command-line entry point | active |
| `src/file_probe.py` | Signature, protection, and route probe | available |
| `src/backup_bundle.py` | Verified byte backup and manifest | available |
| `src/workbook_snapshot.py` | Excel workbook inventory and style snapshot | available |
| `src/excel_session.py` | COM lifecycle and Excel state restoration | available |
| `src/file_transaction.py` | Working-copy and source-integrity controls | available |
| `src/workflow.py` | Current P&L transaction orchestration | available |
| `src/tool_registry.py` | Capability catalogue and readiness state | active |
| `src/agent_contracts.py` | Universal request/result/error/plan models | M1 |
| `src/engine_contract.py` | Engine-independent workbook interface | M1 |
| `src/tool_executor.py` | Safe dispatch for declared tools | M1 |
| `src/fake_engine.py` | Deterministic contract-test engine | M1 |
| `src/agent_entry.py` | Start/status/description output | M1 |
| `src/plan_validation.py` | Deterministic pre-execution plan checks | M1 |
| `src/transaction_state.py` | Generic fail-closed states, workspace locking, hash stages, and auto-persisted journal with safe resume | M3 |
| `src/core/coordinator.py` | Generic coordinator: dry-run half (probe/backup/snapshot/plan-validate/approve) plus the mutating half (working copy/execute/save/reopen/validate/publish), all fail-closed | M3 |
| `src/core/transaction_adapter.py` | Workflow-agnostic Excel session adapters (working copy, reopen, generic preservation checks) the coordinator composes | M3 |
| `src/engines/excel_com.py` / `src/fake_engine.py` | Also implement `clear_range` (ClearContents-only, bounded) | M4 |
| `src/tool_executor.py` | Also implements `clear_range`, hardened `write_range`/`copy_range` (shape checks, before-evidence, `_MAX_CELLS_PER_RANGE_OPERATION` cap, cross-workbook guard), `fill_formula_down`, `insert_columns`, and `update_pivot_source` | M4 |
| `src/business_sheet_updater.py`, `src/total_pl_updater.py` | Fixed a real pywin32 kwarg bug in the production column-insert step (`Resize(ColumnSize=...)` -> `Range(col1, col2)`) found while verifying the new generic `insert_columns` tool against real Excel | M1 |
| `src/engines/excel_com.py` | Explicit-target Excel COM adapter | M2 |
| `src/engines/router.py` | Probe-to-engine decision policy | M2 |
| `src/engines/` | COM and fast-engine adapters | planned |
| `src/tools/` | Universal Excel operations | planned |
| `src/recipes/` | Business-specific compositions | planned |
| `docs/UNIVERSAL_EXCEL_AGENT_CODING_PLAN_V2.md` | Canonical implementation order, gates, and SAP refresh recipe specification | active |
| `tasks/plan.md` | Short pointer to the canonical V2 plan and current starting task | active |
| `tasks/todo.md` | Execution checklist for V2 Tasks 1-16 and release checkpoints | active |
| `tests/unit/` | Pure logic and safety tests | active |
| `tests/contract/` | Tool and engine contract tests | M1 |
| `tests/integration/` | Gated real Excel tests | active |
| `schemas/` | Versioned machine-readable contracts | M1 |
| `schemas/tool_catalog.json` | Generated agent capability catalogue | M1 |
| `schemas/tool_catalog.schema.json` | Catalogue validation schema | M1 |
| `tests/contract/` | Engine/tool boundary contract tests | M1 |
| `vendor/wheelhouse/` | Private/offline build cache | local only |
