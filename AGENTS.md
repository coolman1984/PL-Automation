# Instructions for Every Coding Agent

Read this file first, then read `docs/START_HERE_AGENT.md` and
`docs/PROJECT_STATUS.md`. Do not scan the whole repository before identifying
the single milestone you will implement.

## Mission

This project controls Microsoft Excel safely. It works on a verified working
copy, never on the user's original workbook. The generic engine is separate
from report-specific recipes.

## Required workflow

```text
understand -> probe -> plan -> backup -> snapshot -> dry-run -> approve
-> execute on working copy -> save/close -> reopen -> validate -> publish
```

Stop on ambiguity, unsupported features, missing authorization, failed backup,
source hash change, or failed validation. Never guess.

## Commands

```text
python app.py --agent-start
python app.py --project-status
python app.py --list-tools --format json
python app.py --describe-tool TOOL_NAME
python app.py --run-tool REQUEST_JSON
python app.py --file PATH --probe-only
python -m pytest tests/unit
```

Excel COM integration tests are gated and require an authorized Windows Excel
machine. Do not claim that a COM feature works based on a non-Windows test.

## Forbidden actions

- Do not edit, overwrite, rename, move, or delete the source workbook.
- Do not bypass NASCA, Office passwords, DRM, or corporate controls.
- Do not use `Select`, `Activate`, clipboard automation, keystrokes, or mouse
  automation for workbook operations.
- Do not add a tool to the available catalogue before its release checklist,
  tests, and validation are complete.
- Do not commit workbooks, backups, generated output, secrets, caches, or binary
  wheel packages.
- Do not write arbitrary Python or COM code from an AI request at run time.

## Change discipline

Implement one vertical slice at a time. Add/update tests first, run the narrow
test set, then the full suite. Update `docs/PROJECT_STATUS.md`,
`docs/FILE_MAP.md`, and the tool catalogue in the same change. Keep the current
P&L recipe working while general capabilities are added.
