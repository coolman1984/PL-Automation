"""Task 2 acceptance: CLI output is Unicode-safe under legacy console encodings.

The real production workbook path starts with a star character (U+2605). Under
a CP1252 Windows console this previously aborted the probe with
``UnicodeEncodeError`` before any check result was printed.
"""

from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]
APP = PROJECT_ROOT / "app.py"


def _make_minimal_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as package:
        package.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" />',
        )
        package.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />',
        )


@pytest.mark.skipif(
    not sys.platform.startswith("win"),
    reason="Legacy console encoding behavior is Windows-specific",
)
def test_probe_only_survives_star_path_under_cp1252(tmp_path):
    workbook = tmp_path / "★probe_acceptance.xlsx"
    _make_minimal_xlsx(workbook)
    assert workbook.exists()

    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "cp1252"  # force the legacy console codec
    env.pop("PYTHONUTF8", None)

    proc = subprocess.run(
        [sys.executable, str(APP), "--file", str(workbook), "--probe-only"],
        capture_output=True,
        env=env,
        cwd=str(PROJECT_ROOT),
        timeout=120,
    )
    stdout = proc.stdout.decode("utf-8", errors="replace")
    stderr = proc.stderr.decode("utf-8", errors="replace")
    output = stdout + stderr

    assert "UnicodeEncodeError" not in output, output
    assert "Traceback" not in output, output
    assert proc.returncode == 0, output
    assert "★" in stdout  # the path itself must survive, not be dropped
