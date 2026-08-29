"""Shared fixtures for Excel COM integration tests.

These tests drive real Microsoft Excel and therefore never run by accident:

- they are skipped unless ``PL_COM_TESTS=1``;
- the workbook under test must be provided through ``PL_COM_WORKBOOK``;
- execute-level transactions additionally require ``PL_COM_EXECUTE=1`` so a
  human has explicitly approved touching a disposable working copy chain.

Run from the project root:

    set PL_COM_TESTS=1
    set PL_COM_WORKBOOK=D:\\path\\to\\source.xlsb
    python -m pytest tests\\integration -v

To also permit full execution runs:

    set PL_COM_EXECUTE=1
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip() == "1"


def _on_windows_with_excel() -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def pytest_collection_modifyitems(config, items):
    if _env_flag("PL_COM_TESTS") and _on_windows_with_excel():
        return
    reason = "COM integration requires PL_COM_TESTS=1, Windows, and pywin32"
    skip_marker = pytest.mark.skip(reason=reason)
    for item in items:
        if Path(item.fspath).parts[-2:-1] == ("integration",) or "integration" in str(item.fspath):
            item.add_marker(skip_marker)


@pytest.fixture(scope="session")
def com_ready() -> None:
    if not _on_windows_with_excel():
        pytest.skip("Excel/pywin32 unavailable")
    try:
        import pythoncom

        pythoncom.CoInitialize()
    except Exception as exc:  # pragma: no cover - environment specific
        pytest.skip(f"COM initialization failed: {exc}")


@pytest.fixture(scope="session")
def source_workbook_path(com_ready) -> Path:
    value = os.environ.get("PL_COM_WORKBOOK", "").strip()
    if not value:
        pytest.skip("Set PL_COM_WORKBOOK to the absolute path of the source workbook")
    path = Path(value)
    assert path.suffix.casefold() == ".xlsb"
    assert path.exists()
    return path


@pytest.fixture()
def execute_approved() -> None:
    if not _env_flag("PL_COM_EXECUTE"):
        pytest.skip("Execute runs need explicit approval: set PL_COM_EXECUTE=1")


@pytest.fixture()
def config():
    from src.config import load_config

    return load_config(PROJECT_ROOT / "config.yaml", year=2026, month=8, execution=True)
