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
| `src/engines/excel_com.py` | Explicit-target Excel COM adapter | M2 |
| `src/engines/router.py` | Probe-to-engine decision policy | M2 |
| `src/engines/` | COM and fast-engine adapters | planned |
| `src/tools/` | Universal Excel operations | planned |
| `src/recipes/` | Business-specific compositions | planned |
| `tests/unit/` | Pure logic and safety tests | active |
| `tests/contract/` | Tool and engine contract tests | M1 |
| `tests/integration/` | Gated real Excel tests | active |
| `schemas/` | Versioned machine-readable contracts | M1 |
| `schemas/tool_catalog.json` | Generated agent capability catalogue | M1 |
| `schemas/tool_catalog.schema.json` | Catalogue validation schema | M1 |
| `tests/contract/` | Engine/tool boundary contract tests | M1 |
| `vendor/wheelhouse/` | Private/offline build cache | local only |
