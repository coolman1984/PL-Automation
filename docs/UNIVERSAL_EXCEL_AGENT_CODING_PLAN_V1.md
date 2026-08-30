# Universal Excel Agent — Coding Execution Plan V1

**Status:** M1 complete; M2 in progress  
**Target:** Windows 10/11 with desktop Microsoft Excel  
**Primary runtime:** Private portable Python runtime  
**Primary fidelity engine:** Microsoft Excel COM  
**Design priority:** Safety first, fidelity second, speed third, simplicity always

## 1. Mission

Turn this project from a single P&L recipe into a small, clear, universal Excel
control toolkit that a newly arrived AI coding agent can understand in minutes.
The toolkit must support advanced workbook automation without changing the
original file, silently dropping Excel features, or claiming success without
evidence.

"Any workbook" means any supported Excel workbook that the signed-in user and
desktop Excel are authorized to open. It does not mean bypassing passwords,
DRM, corporate policy, file corruption, or Excel's own limits.

## 2. Definition of success

A fresh coding agent must be able to:

1. Read one entry file and understand the architecture, safety rules, current
   status, and next task.
2. Discover all callable tools from a machine-readable catalogue.
3. Convert a user request into a validated operation plan.
4. Create and verify a backup before any mutation.
5. Select the safest compatible engine automatically.
6. Modify only a working copy.
7. validate the requested result and workbook integrity after reopening it.
8. Publish a separate output or restore safely.
9. Produce a readable report and a complete JSON event record.
10. Refuse unsupported or ambiguous work with a precise reason.

## 3. Existing baseline — preserve and extend

Already implemented and tested locally:

- file signature and protection probe;
- guarded Excel COM session modes: attach, open, and automatic routing;
- exact byte backup with SHA-256 manifest;
- workbook inventory/full snapshot with style deduplication;
- working-copy transaction and source fingerprint checks;
- P&L A08 recipe, validation, reporting, and idempotency checks;
- tool catalogue that distinguishes available and planned capabilities;
- portable/offline build and repair scripts;
- 84 passing tests and two Windows COM tests intentionally gated.

Do not rewrite these components without a failing test or a documented design
reason. Generalize them behind stable interfaces.

## 4. Non-negotiable safety invariants

1. Never save, rename, move, or edit the source workbook.
2. No mutation tool can run without a verified backup and transaction ID.
3. All writes target a working copy created from Excel `SaveCopyAs` when Excel
   is the selected engine.
4. The source SHA-256 must match before and after the run.
5. A plan must resolve exact workbook, sheet, range, and expected effect before
   execution.
6. Ambiguous matches stop; the agent must never guess a sheet, block, formula,
   table, or pivot target.
7. Every mutation records before/after evidence sufficient for validation and,
   where practical, reversal.
8. The result must be saved, closed, reopened, and validated before publication.
9. Failure never publishes an output as successful.
10. Protected content is used only through the user's authorized Excel session;
    the project must never attempt to defeat protection or DRM.
11. Existing macros, links, charts, pivots, shapes, connections, and unknown
    workbook parts must be preserved unless the explicit task changes them.
12. Planned tools remain non-callable until their tests and acceptance gates pass.

## 5. Simplicity contract for coding agents

The repository must expose these six authoritative files:

| File | Purpose |
|---|---|
| `AGENTS.md` | First-read rules, exact commands, forbidden actions |
| `docs/START_HERE_AGENT.md` | Ten-minute architecture and workflow overview |
| `docs/PROJECT_STATUS.md` | Completed, in progress, blocked, next work |
| `docs/FILE_MAP.md` | One-line purpose and ownership for every module |
| `schemas/tool_catalog.json` | Machine-readable capabilities and status |
| `docs/RECIPES.md` | Minimal examples of planning and executing tasks |

Rules:

- one canonical fact, linked from other documents instead of duplicated;
- each module has one responsibility and a public interface;
- business recipes never call COM directly;
- engine implementations never contain report-specific business logic;
- all public operations use the same request/result/error envelope;
- generated files, caches, workbooks, backups, logs, and wheel binaries stay out
  of source control.

## 6. Target architecture

```text
app.py                         thin command-line entry point
src/
  agent/                       request planning and capability discovery
    planner.py
    plan_models.py
    policy.py
  core/                        engine-independent safety and orchestration
    coordinator.py
    transaction.py
    backup.py
    evidence.py
    validation.py
    errors.py
  engines/                     workbook technology adapters
    base.py
    router.py
    excel_com.py
    openxml_fast.py
    read_only_probe.py
  tools/                       small universal Excel operations
    workbook.py
    worksheet.py
    range.py
    formula.py
    formatting.py
    table.py
    chart.py
    pivot.py
    names_links.py
    objects.py
    macro.py
  recipes/                     task-specific composition only
    pnl_a08.py
  inspection/
    file_probe.py
    workbook_snapshot.py
    diff.py
  reporting/
    journal.py
    run_report.py
schemas/
tests/
  unit/
  contract/
  fixtures/
  integration/
  acceptance/
```

Migrate incrementally. Keep compatibility imports until every existing test and
launcher uses the new structure.

## 7. Universal operation contract

Every callable tool must accept and return JSON-serializable data.

### Request envelope

```json
{
  "schema_version": "1.0",
  "transaction_id": "run-id",
  "tool": "range.write_values",
  "target": {
    "workbook_id": "working-copy-id",
    "sheet": "Sheet1",
    "address": "A1:C10"
  },
  "arguments": {},
  "preconditions": [],
  "expected_effect": {},
  "dry_run": true
}
```

### Result envelope

```json
{
  "ok": true,
  "tool": "range.write_values",
  "changed": true,
  "affected_ranges": ["Sheet1!A1:C10"],
  "before_evidence": {},
  "after_evidence": {},
  "warnings": [],
  "metrics": {"elapsed_ms": 0, "cells_touched": 30}
}
```

### Error envelope

```json
{
  "ok": false,
  "code": "ambiguous_target",
  "message": "More than one matching table was found.",
  "recoverable": true,
  "details": {},
  "suggested_action": "Specify the worksheet and table name."
}
```

No tool returns raw COM objects. No tool depends on the active sheet, active
cell, clipboard, selection, or visible UI state.

## 8. Engine routing policy

| Workbook condition | Default route | Rule |
|---|---|---|
| NASCA/DRM or user-open protected file | Excel COM attach | Require exact authorized open workbook |
| XLSB, XLSM, XLAM, advanced XLSX | Excel COM | Preserve full Excel fidelity |
| Simple XLSX with no advanced parts | Fast engine candidate | Only after feature scan and round-trip tests |
| Unknown, damaged, mismatched extension | Stop | Report evidence; never force open/write |
| Read-only inventory request | Probe or Excel COM | Use the least expensive safe reader |

The fast engine is an optimization, never a compatibility promise. The official
openpyxl documentation warns that not every Excel item is preserved and shapes
can be lost when an existing file is opened and saved. Therefore it must not be
used for mutation when unsupported parts exist.

## 9. Implementation phases

### Phase 0 — Freeze and characterize the baseline

Tasks:

- tag the current V7 behavior in tests;
- create small sanitized fixture workbooks for XLSX/XLSM and a locally created
  XLSB integration fixture;
- record existing CLI output and P&L recipe behavior as golden tests;
- add a source-control check that rejects Excel files, backups, secrets, wheels,
  runtime output, and caches.

Exit gate:

- all current tests pass unchanged;
- fixtures contain no company data;
- baseline behavior has reproducible evidence.

### Phase 1 — Agent entry pack — **COMPLETE**

Create the six authoritative files in section 5. Generate the tool catalogue
from Python metadata so documentation cannot drift from executable tools.

Add commands:

```text
python app.py --agent-start
python app.py --project-status
python app.py --list-tools --format json
python app.py --describe-tool TOOL_NAME
```

Exit gate:

- a clean-session coding agent can identify architecture, safety rules, test
  command, next milestone, and available tools by reading `AGENTS.md` only;
- documentation link checker and catalogue consistency tests pass.

### Phase 2 — Stable contracts and compatibility layer — **IN PROGRESS**

Create:

- typed plan, target, request, result, evidence, and error models;
- `ExcelEngine` protocol/abstract base class;
- `ToolDefinition` metadata with risk and availability state;
- compatibility adapters around existing modules;
- schema versioning and JSON Schema files.

Required engine methods:

```text
open_readonly, create_working_copy, open_working_copy, close,
save, recalculate, reopen, inspect, execute_tool
```

Exit gate:

- existing P&L workflow runs through the new interfaces;
- no recipe imports a COM implementation;
- contract tests run against a fake engine without Excel.

### Phase 3 — Transaction coordinator

Implement this state machine:

```text
RECEIVED -> PROBED -> BACKED_UP -> SNAPSHOTTED -> PLANNED -> APPROVED
-> WORKING_COPY_READY -> EXECUTING -> SAVED -> REOPENED -> VALIDATED
-> PUBLISHED
```

Any failure moves to `FAILED_SAFE`; publication is impossible from that state.

Add:

- append-only event journal with sequence numbers;
- atomic JSON writes;
- per-run workspace and lock;
- source/backup/working/output hashes;
- resumable checkpoints only at safe state boundaries;
- cleanup policy that never deletes diagnostic evidence automatically.

Exit gate:

- injected failure tests at every transition prove no source mutation and no
  false publication;
- interrupted runs are reported clearly and can be retried safely.

### Phase 4 — Core workbook and range tools

Implement and unlock in this order:

1. `workbook.inspect`
2. `worksheet.list/get/create/copy/rename/move/hide/delete`
3. `range.read_values/read_formulas`
4. `range.write_values/write_formulas/clear`
5. `range.copy/move/insert/delete`
6. `row.insert/delete/copy/set_height/hide`
7. `column.insert/delete/copy/set_width/hide/autofit`
8. `merge.create/remove`
9. `freeze_panes.set/clear`

Deletion and clear tools require an explicit expected target size and before
fingerprint. Structural operations require formula-reference validation.

Performance rules:

- transfer rectangular values/formulas through array assignments, not one COM
  call per cell;
- process large ranges in configurable blocks;
- always fully qualify workbook, worksheet, and range objects;
- capture and restore Excel calculation, alerts, events, and screen updating in
  `finally` blocks;
- use `Value2` for raw value transfer and treat dates/currency deliberately;
- preserve `Formula` versus `Formula2` semantics for legacy and dynamic arrays.

Exit gate:

- every tool has unit, contract, COM integration, dry-run, idempotency, invalid
  target, protected target, and rollback tests;
- no tool appears as available before all mandatory tests pass.

### Phase 5 — Formatting and data features

Implement:

- font, fill, border, alignment, number format, protection, and named style;
- exact style copy from source range;
- conditional formatting inventory/copy/create/delete;
- data validation inventory/copy/create/delete;
- comments/notes and hyperlinks;
- tables, filters, sorting, subtotal, grouping, and outline levels;
- row heights, column widths, merged areas, print area, page setup, and panes.

Use style hashes to compare before/after without repeating identical style JSON.

Exit gate:

- visual/fidelity snapshot diff passes for every supported feature;
- unrelated styles and workbook objects remain unchanged.

### Phase 6 — Advanced Excel objects

Implement separate tools for:

- charts and series;
- pivot tables, pivot caches, slicers, and refresh controls;
- workbook/worksheet names;
- external links and connections;
- shapes, pictures, text boxes, and controls;
- VBA module inventory and explicitly authorized macro execution;
- workbook calculation and refresh orchestration.

Rules:

- do not create or edit VBA code in the first release;
- macro execution is disabled by default and requires explicit user approval;
- refreshes are opt-in because they may access external systems;
- never break a shared pivot cache without an explicit plan.

Exit gate:

- open/save/reopen inventory proves object counts and identities are preserved;
- each mutation validates both the target object and unaffected neighboring
  objects.

### Phase 7 — Planning and recipe system

Create a deterministic plan format containing:

- user intent summary;
- resolved targets;
- assumptions and unresolved ambiguity;
- ordered tools;
- preconditions and expected effects;
- validation rules;
- risk score and approval requirement;
- estimated cells and objects touched.

The AI agent proposes the plan; deterministic Python validates and executes it.
The agent must not generate arbitrary Python or COM statements during a normal
run.

Move the P&L A08 logic into the first recipe using only universal tools. Add
recipe composition, parameter validation, dry run, and recipe-specific
acceptance rules.

Exit gate:

- the P&L result matches the frozen V7 behavior;
- at least five generic recipes use no direct engine calls;
- invalid plans are rejected before a working copy is edited.

### Phase 8 — Large workbook performance

Add:

- range chunk planner based on cell count and measured payload size;
- sparse used-range detection and bounded searches;
- bulk read/write metrics and adaptive block sizing;
- progress events and cancellation at safe boundaries;
- timeouts only around operations that can be abandoned safely;
- calculation mode planning: targeted range, sheet, workbook, or full rebuild;
- memory and disk preflight checks;
- benchmark fixtures at 100k, 500k, and 1M rows where the format permits.

Never split a structural operation in a way that leaves a partially inserted
table, broken merge, or half-rewritten formula block.

Exit gate:

- benchmarks define supported limits rather than advertising unlimited size;
- memory remains bounded during bulk data operations;
- cancellation produces a failed working copy, never a published partial result.

### Phase 9 — Packaging and one-click operation

Provide:

- private portable runtime; no global installation or administrator rights;
- verified offline package folder with hashes;
- one-click self-check, repair, prepare, dry-run, and execute launchers;
- clear detection of Python, Excel bitness, COM availability, permissions, disk
  space, workbook lock, and output path;
- source and packaged-mode acceptance tests on clean Windows machines.

Exit gate:

- works offline on a clean approved Windows test PC;
- package does not rely on PATH, Store aliases, internet, or globally installed
  packages;
- a nontechnical user can run it from one launcher and receive a clear result.

### Phase 10 — Production acceptance

Run the matrix below on sanitized files and authorized disposable copies:

| Dimension | Required cases |
|---|---|
| Format | XLSX, XLSM, XLSB |
| Complexity | simple, formulas, charts, pivots, links, VBA, objects |
| Protection | none, sheet/workbook protection, authorized NASCA session |
| Size | small, medium, 500k-row, near Excel worksheet limit |
| Excel state | closed, already open, read-only, locked, unsaved changes |
| Failure | disk full, permission denied, Excel crash, timeout, validation mismatch |

Production gate:

- zero source mutations across the complete failure matrix;
- zero silent feature loss;
- every success has reopened validation and hashes;
- every failure has a useful error report and recoverable backup;
- security, privacy, performance, and business-owner sign-off are recorded.

## 10. Required validation layers

Every plan chooses relevant validators from all five layers:

1. **File integrity:** readable, expected format, stable source hash, output exists.
2. **Workbook structure:** sheets, names, VBA flag, links, connections, objects.
3. **Target effect:** expected cells, formulas, styles, objects, or dimensions.
4. **Business rule:** totals, reconciliations, period mappings, uniqueness.
5. **Non-regression:** untouched areas and advanced feature inventory unchanged.

A byte hash of the output proves identity to the validated working copy; it does
not prove business correctness. Both are required.

## 11. Tool release checklist

A tool changes from `planned` to `available` only when all boxes are true:

- [ ] typed request, result, and error contract;
- [ ] dry-run output;
- [ ] precise target resolution;
- [ ] risk classification and approval rule;
- [ ] before/after evidence;
- [ ] validation function;
- [ ] unit tests;
- [ ] fake-engine contract tests;
- [ ] real Excel COM integration test;
- [ ] large-range test when applicable;
- [ ] protection/permission failure test;
- [ ] interruption/rollback test;
- [ ] documentation and one minimal recipe example;
- [ ] catalogue generated and consistent.

## 12. Coding rules

- Python standard library first; add a dependency only with a documented need.
- Keep `pywin32` as the Windows COM dependency and `PyYAML` only while YAML
  configuration remains necessary.
- Use dataclasses/type hints and small pure functions around COM boundaries.
- Centralize COM constants and tolerant property reads.
- Never use broad `except Exception` without recording the original exception and
  converting it to a stable domain error at the boundary.
- Release COM references in deterministic reverse order and always restore Excel
  application state.
- No clipboard automation, keystrokes, mouse automation, `Select`, or `Activate`.
- No hidden network calls, telemetry, automatic uploads, or credential handling.
- No destructive cleanup of source, backup, failed run, or user files.
- Keep each commit focused; update tests, catalogue, file map, and project status
  in the same commit as a capability change.

## 13. Recommended implementation sequence for the coding agent

For each phase:

1. Read `AGENTS.md`, project status, this plan, and only the referenced modules.
2. Inspect current tests and repository state; do not assume the plan is newer
   than the code.
3. Mark one milestone `in_progress` in project status.
4. Write or update failing tests first.
5. Implement the smallest complete vertical slice.
6. Run unit and contract tests, then gated Windows tests when available.
7. Run static checks and secret/source-workbook checks.
8. Update capability status only after acceptance passes.
9. Update project status with evidence, remaining gaps, and exact next action.
10. Stop when blocked by Excel/Windows/authorization instead of faking evidence.

Do not build all tools at once. Complete and unlock one coherent tool group at a
time.

## 14. Current milestone

**Milestone M2 — Read-only Excel engine foundation**

Deliverables:

- `src/engines/excel_com.py` implementing the explicit-target engine contract;
- `src/engines/router.py` implementing probe-to-engine selection;
- `src/transaction_state.py` providing the reusable fail-closed state journal;
- read-only `read_range` behind the engine contract;
- gated Windows integration coverage for a real Excel workbook;
- no generic mutation tool unlocked until the read-only acceptance gate passes.

M1 is complete. A clean-session agent can answer these questions from the
repository without searching the whole codebase:

1. What is safe to run now?
2. What is planned but locked?
3. How is an original workbook protected?
4. Which engine will be used and why?
5. What test proves the next change?
6. What is the single next implementation task?

## 15. Official technical references

- Excel Range object: https://learn.microsoft.com/en-us/office/vba/api/excel.range(object)
- Excel Range.Value2: https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2
- Formula versus Formula2: https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2
- Workbooks.Open: https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open
- Workbook.SaveCopyAs: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas
- Excel performance guidance: https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions
- openpyxl preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- Python on Windows and embedded distribution: https://docs.python.org/3/using/windows.html
