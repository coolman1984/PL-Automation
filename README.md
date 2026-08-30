# Excel Agent Foundation + P&L August Actual Automation

> **Current status — NOT PRODUCTION READY (2026-08-30).** The D1-D16 safety
> remediation is implemented and 157 automated tests pass, but the corrected
> mutation paths still require real Windows Excel re-acceptance. Follow
> `docs/PROJECT_STATUS.md` before use.

This tool inserts the Actual columns for **August 2026 (A08)** into the protected
P&L workbook by driving desktop Microsoft Excel through COM only.

V7 also starts the general Excel-agent layer. `PREPARE_WORKBOOK.bat` creates a
verified byte-for-byte backup and a read-only workbook inventory before any
future agent operation. `LIST_AGENT_TOOLS.bat` exposes the machine-readable
capability catalogue and clearly separates available tools from planned tools.

The original workbook is never saved, edited, or moved. The tool works on a
byte-for-byte Excel-created working copy and publishes a separate validated
output file only after every safety gate passes.

## What it does

For each sheet `VD Total`, `MX Total`, `DA Total`, and `Total PL` inside the
August block:

```text
before : ... | T08 | % | S08 | % | September ...
after  : ... | T08 | % | S08 | % | A08 | % | September ...
```

`Total PL` A08 amounts are formula-linked to the three business sheets' new A08
columns and reconciled to `Total PL A08 = VD A08 + MX A08 + DA A08`.

## Smart quick check and engine routing

Before Excel is opened, the launcher now performs a dependency-free read-only
signature check. It reads only the container header/member names and never
writes to the source.

| Detected file | Automatic safe route |
|---|---|
| NASCA DRM / Office-encrypted workbook | Manually open in authorized Excel, then COM attach |
| Normal unprotected XLSB | Isolated hidden Excel COM open |
| Simple unprotected XLSX | Marked as a future fast-engine candidate |
| XLSX with VBA, drawings, charts, pivots, links, connections, or embedded objects | Excel COM |
| Unknown/mismatched container | Stop for review |

Important: protection is not the only decision. This P&L workbook is XLSB and
contains formulas and advanced Excel structures. Free fast XLSB tools are
currently readers or incomplete writers; the emerging `rxlsb` project lists
formula and chart support as planned, while `pyxlsb` and Calamine are read-only.
For that reason, the production write route for XLSB remains Excel COM even
when the file is not protected. This prevents speed from being purchased by
losing formulas, styles, links, charts, or workbook integrity.

For XML workbooks, openpyxl itself warns that it does not preserve every Excel
item and can lose shapes when an existing file is opened and saved. The router
therefore marks only simple XLSX files as candidates; the P&L updater does not
execute that future engine yet.

References:

- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/acc8aa92-1f02-4167-99f5-84f9f676b95a
- https://github.com/willtrnr/pyxlsb
- https://github.com/tafia/calamine
- https://github.com/itcraft-cn/rxlsb
- https://openpyxl.readthedocs.io/en/stable/tutorial.html

## Requirements

- Windows with Microsoft Excel desktop installed
- The source `.xlsb` workbook

The production portable release contains its own Python runtime and packages.
The user's PC does not need Python, pip, internet access, administrator rights,
or PATH changes. Python 3.12, 3.13, or 3.14 is required only on the Windows
build computer; the matching locked packages are already included offline.

## One-time setup

### Production user

Extract the complete portable ZIP and run `SELF_CHECK.bat`. No installation is
required.

### Developer/source mode

```bat
cd pl_actual_automation
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

For the locked portable build and offline repair process, follow
`PORTABLE_DEPLOYMENT.md`.

## Usage

### General Excel safety preparation

Drag any recognized Excel workbook onto `PREPARE_WORKBOOK.bat`. It creates:

- an exact original-file backup;
- `backup_manifest.json` with source and backup SHA-256 evidence;
- `workbook_snapshot.json` with sheets, used ranges, widths/heights inventory,
  objects, links, pivots, connections, VBA presence, and protection facts.

Automatic mode is the safe default: it records every cell when the workbook is
within the configured safety limit, otherwise it creates the complete workbook,
object, row-height, and column-width inventory and records a warning. Full mode
additionally requires every cell to fit the explicit limit and records every cell,
formula, value, row height, column width, and deduplicated font/fill/border/
alignment/number-format style. It stops instead of silently truncating when the
configured cell limit would be exceeded.

### Easy way

Drag the protected workbook onto `RUN_A08.bat`.
It runs a dry-run first, asks for explicit confirmation, then executes,
and finally opens the `output\` folder on success.

For NASCA-protected workbooks, the safest operating procedure is to open the
source manually in desktop Excel, confirm it is saved, and run with
`--mode attach`.  This lets the tool use the already-authorized Excel session.

### Command line

```text
python app.py --file PATH --year 2026 --month 8 --dry-run [--mode auto|attach|open]
python app.py --file PATH --year 2026 --month 8 --execute  [--mode auto|attach|open]
python app.py --file PATH --probe-only
python app.py --list-tools
python app.py --file PATH --backup-only
python app.py --file PATH --prepare --snapshot-mode inventory
python app.py --file PATH --snapshot --snapshot-mode full --max-snapshot-cells 250000
```

| Option | Meaning |
|---|---|
| `--dry-run` | Read-only discovery proof. Prints resolved codes, discovered block coordinates, and any problems. Edits nothing. |
| `--execute` | Full guarded transaction on a working copy. |
| `--probe-only` | Fast signature/protection/format check. Does not import or open Excel. |
| `--mode auto` (default) | Attach to running Excel if the exact file is open there; otherwise open it in an isolated hidden instance. |
| `--mode attach` | Require the exact workbook to be open in your Excel session. |
| `--mode open` | Always open an isolated hidden Excel instance. |
| `--config PATH` | Configuration YAML; default `config.yaml`. |
| `--verbose` | Verbose logging to console and run log. |

### Exit codes

```text
0  success / dry-run ready
2  CLI or configuration error
3  preflight not ready
4  safe execution failure
5  validation failure
6  publication failure
```

## Safety model (what you can rely on)

1. Only Excel COM ever changes or saves workbook content.
2. Sequence: `SaveCopyAs` a working copy from the untouched source → edit the
   copy → validate → save → close → reopen and re-validate → byte-copy publish.
3. SHA-256 of the source is captured before and after; a mismatch aborts everything.
4. Any ambiguity (duplicate August block, existing A08, unprovable Total PL row,
   failed validation) stops the run with nothing published.
5. Re-running on an already-updated workbook refuses before editing anything.
6. Failed runs keep diagnostics under `work\<run-id>\` and the failed working
   file under `failed_runs\<run-id>\`.

## Run evidence

Every execution writes into `work\<run-id>\`:

- `automation.log` - step-by-step structured log
- `run_report.txt` - human-readable report including validation results
- `run_manifest.json` - machine-readable manifest with fingerprints, discovery
  evidence, updates, validations, and error details when applicable

A successful final workbook lands in `output\<name>__A08_UPDATED.xlsb` with a
filename collision guard, and its hash equals the validated working copy's hash.

## Development

The universal Excel-agent implementation roadmap is in
`docs/UNIVERSAL_EXCEL_AGENT_CODING_PLAN_V1.md`. Coding agents should execute it
phase by phase and must not unlock a planned tool before its release checklist
passes.

Unit tests (no Excel required):

```bat
.venv\Scripts\python -m pytest tests\unit
```

Excel COM integration tests are gated behind environment variables so they never
run by accident; see `tests/integration/README.md`.

Implementation follows `docs/PL_A08_AUTOMATION_DETAILED_EXECUTION_PLAN_V1.md`
(module map in section 4, safety invariants in section 3).
