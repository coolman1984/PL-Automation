"""Task 3 acceptance: the read-only COM acceptance harness is well-formed.

The harness cannot run real Excel inside unit tests, so this module verifies
that the script parses cleanly and enforces the evidence rules required by
the V2 plan (hash equality, exit-code gate, orphan Excel PID check, and a
retained JSON report).
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]
HARNESS = PROJECT_ROOT / "tools" / "run_read_range_acceptance.ps1"


def _script_text() -> str:
    assert HARNESS.exists(), f"Missing harness script: {HARNESS}"
    return HARNESS.read_text(encoding="utf-8-sig")


def test_harness_script_parses_as_powershell():
    if not sys.platform.startswith("win"):
        pytest.skip("PowerShell syntax check requires Windows")
    checker = (
        "$errors = $null; "
        "$null = [System.Management.Automation.Language.Parser]::ParseFile("
        f"'{HARNESS}', [ref]$null, [ref]$errors); "
        "if ($errors -and $errors.Count -gt 0) { "
        "$errors | ForEach-Object { Write-Output $_.Message }; exit 1 } else { exit 0 }"
    )
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", checker],
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_harness_enforces_required_evidence_rules():
    text = _script_text()
    # Hash evidence before and after, and equality gate.
    assert "Get-FileHash" in text
    assert "source_hash_unchanged" in text
    # Nonzero pytest exit fails the harness.
    assert "pytest_exit_zero" in text
    # Orphan Excel PID detection against the pre-run baseline.
    assert "Get-Process -Name \"EXCEL\"" in text
    assert "no_orphaned_excel_pids" in text
    # Retained JSON acceptance report.
    assert "ConvertTo-Json" in text
    assert "read_range_acceptance_" in text
    # COM shutdown diagnostics are recorded, and only treated as a crash on
    # nonzero pytest exit.
    assert "com_shutdown_diagnostic_seen" in text
    assert "com_shutdown_diagnostic_treated_as_crash" in text
