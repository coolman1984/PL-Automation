# V7 — General Excel Agent Foundation

## Ready now

- `LIST_AGENT_TOOLS.bat` shows the agent's complete capability catalogue and
  distinguishes ready tools from planned tools.
- `PREPARE_WORKBOOK.bat` creates a verified byte-for-byte backup before any
  future editing and then produces `workbook_snapshot.json` through Excel.
- Automatic snapshot mode captures every cell and exact reusable style when the
  workbook fits the configured limit; otherwise it keeps the complete workbook,
  object, row-height, and column-width inventory and records a warning.
- The snapshot records formulas, values, fonts, colors, fills, borders,
  alignment, number formats, validation, comments, merges, shapes, charts,
  tables, pivots, conditional formats, links, connections, names, VBA presence,
  sheet states, and protection facts when available through Excel.
- The original workbook file remains the recovery source of truth because JSON
  cannot independently reconstruct every binary Excel object.

## Verified locally

- 59 tests passed; 1 Windows Excel COM integration test skipped.
- The real NASCA file was identified correctly.
- Its test backup was byte-identical by SHA-256.
- Offline Windows packages resolve for Python 3.12, 3.13, and 3.14.

## Still blocked until Windows

- Build and run the portable executable.
- Capture a real snapshot from the authorized Excel/NASCA session.
- Execute and validate the P&L recipe on a disposable copy.
- Approve the generic write tools only after transaction and recovery tests.
