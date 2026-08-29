# General Excel Agent — Updated Roadmap

## Status legend

- ✅ implemented and locally tested
- ⚠️ implemented but requires Windows Excel proof
- ⬜ not implemented

## Phase 1 — safety foundation

- ✅ File type, protection, and engine routing.
- ✅ Byte-for-byte backup with source/backup SHA-256 comparison.
- ✅ Collision-safe backup directories and atomic JSON manifests.
- ⚠️ Read-only workbook inventory through Excel COM.
- ⚠️ Full cell/style snapshot with deduplicated reusable styles.
- ✅ Automatic size gate that never silently truncates a full snapshot.
- ✅ Machine-readable tool catalogue with readiness status.
- ✅ Portable/offline build system and packages for Python 3.12–3.14.

## Phase 2 — controlled cell operations

- ⬜ Read ranges in bulk.
- ⬜ Write values and formulas in bulk to a working copy only.
- ⬜ Copy ranges using Excel-native behavior.
- ⬜ Apply fonts, fills, borders, alignment, and number formats.
- ⬜ Insert/delete rows and columns with explicit boundaries.
- ⬜ Find/replace with preview and match limits.

## Phase 3 — workbook objects

- ⬜ Create/rename/move/hide sheets.
- ⬜ Create and resize tables.
- ⬜ Create and modify charts.
- ⬜ Manage names, validation, conditional formats, comments, and hyperlinks.
- ⬜ Controlled refresh of queries, links, connections, and pivots.
- ⬜ Calculation completion and formula-error scan.

## Phase 4 — guarded agent transaction

- ⬜ Accept a versioned JSON operation plan.
- ⬜ Validate every operation before editing.
- ⬜ Create backup and Excel working copy automatically.
- ⬜ Execute only allowlisted tools; reject arbitrary code.
- ⬜ Save, close, reopen, compare, and publish atomically.
- ⬜ Record every request, result, warning, and changed range.
- ⬜ Restore a verified backup to a new recovery path.

## Phase 5 — acceptance

- ⬜ Build the executable on Windows 10/11 x64.
- ⬜ Run self-check with Excel installed.
- ⬜ Prove snapshots on simple XLSX, complex XLSM/XLSB, and NASCA files.
- ⬜ Run two stable end-to-end executions with no manual repair.
- ⬜ Test forced failure, cancellation, corruption detection, and recovery.
- ⬜ Finance-owner approval for the P&L recipe.

## Current decision

The safety/read foundation is ready for Windows proof. Generic workbook write
tools remain intentionally unavailable until the guarded transaction layer is
implemented and tested; the catalogue reports them as `planned` so an AI agent
cannot mistake them for safe callable tools.
