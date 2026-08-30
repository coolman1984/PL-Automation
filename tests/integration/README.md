# Integration tests (real Excel COM)

These tests touch **real Excel**. They exist for plan tasks that cannot be proven
any other way (Tasks 12-21 verification steps).

## Safety gates

The whole directory is skipped unless **all** of these hold:

1. Environment variable `PL_COM_TESTS=1`
2. Running on Windows with `pywin32` importable
3. Source provided via `PL_COM_WORKBOOK=<absolute path to .xlsb>`
4. Any test that performs an *execution* run also needs `PL_COM_EXECUTE=1`

## Commands

```bat
cd pl_actual_automation
set PL_COM_TESTS=1
set PL_COM_WORKBOOK=D:\full\path\to\★Final PL Statement S08 T09 V4(1).xlsb
set PL_COM_READ_SHEET=Data
set PL_COM_READ_ADDRESS=A1:C10
.venv\Scripts\python -m pytest tests\integration -v

rem full execution acceptance (approved disposable run only):
set PL_COM_EXECUTE=1
.venv\Scripts\python -m pytest tests\integration -v
```

The real production run remains the exclusive job of `app.py --execute` /
`RUN_A08.bat`; these tests exist to prove behavior on disposable copies derived
through the approved SaveCopyAs transaction.

## Deterministic read-only acceptance harness (V2 Task 3)

`tools\run_read_range_acceptance.ps1` runs only `test_read_range.py` and
records the acceptance evidence required by the V2 plan: source SHA-256
before/after, exact parameters, pytest exit code and log, and Excel PIDs
before/after. A JSON report is retained under `work\acceptance\`. A handled
COM shutdown diagnostic (`0x80010108`) in the output is recorded but is not
treated as a crash when pytest exits zero.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\run_read_range_acceptance.ps1 `
  -Workbook "D:\full\path\to\★Final PL Statement S08 T09 V4(1).xlsb" `
  -Sheet "VD Total" -Address "A1:C10"
```

The harness fails (exit 1) when pytest fails, the source hash changes, or a
new isolated Excel process remains.
