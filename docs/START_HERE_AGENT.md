# Start Here — Universal Excel Agent

## What this project is

This is a guarded Excel automation engine. It can inspect workbooks, create
verified recovery evidence, and run approved recipes through desktop Excel.
The long-term target is a universal tool set for cells, formulas, formatting,
structure, tables, charts, pivots, links, objects, and calculations.

## Read in this order

1. `AGENTS.md` — non-negotiable rules.
2. `docs/PROJECT_STATUS.md` — what is actually complete today.
3. `docs/FILE_MAP.md` — where each responsibility lives.
4. `docs/UNIVERSAL_EXCEL_AGENT_CODING_PLAN_V2.md` — current phase-by-phase plan.
   V1 is retained as design history.
5. `python app.py --list-tools --format json` — callable versus locked tools.
6. `python app.py --run-tool REQUEST_JSON` — execute one declared request file.

## Mental model

The agent is a planner, not a free-form macro writer. It selects declared tools
and supplies explicit targets, preconditions, expected effects, and validators.
Python owns safety and evidence. Excel COM owns high-fidelity workbook writes
when advanced or protected features are involved.

## First response to a new task

1. Identify the exact workbook and user-intended output.
2. Run the dependency-free file probe.
3. Decide whether the request is read-only or mutating.
4. For mutation, create a verified backup and snapshot before planning changes.
5. Resolve exact sheets/ranges/objects and list uncertainties.
6. Produce a dry-run plan and request approval when risk requires it.
7. Execute only declared, available tools.
8. Reopen and validate the output; publish only on success.

## Current boundary

The declared core tool set now covers cells, formulas, formatting, rows,
sheets, tables, filters, validation, notes, hyperlinks, PivotTables, charts,
names, connections, calculation, and workbook validation. Use
`python app.py --list-tools --format json` as the source of truth for exact
contracts. The implementation still requires Windows/Excel re-acceptance
before production use. A protected workbook requires the user to open it
manually in authorized Excel first. Never bypass that authorization.
