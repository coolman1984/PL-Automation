"""Integration Tests 13-15 from the execution plan: idempotency & forced stops."""

from __future__ import annotations

import sys

import pytest

if not sys.platform.startswith("win"):
    pytest.skip(
        "Excel COM integration tests require Windows desktop Excel",
        allow_module_level=True,
    )
pytest.importorskip("pythoncom")

from src.file_transaction import sha256_file
from src.workflow import preflight


def _hash(path) -> str:
    return sha256_file(path)


def test_preflight_discovers_four_blocks(source_workbook_path, config):
    result = preflight(source_workbook_path, config, mode="open")
    assert result.mode_used == "open"
    assert sorted(result.blocks) == sorted(
        (*config.target_sheets, config.total_sheet)
    )
    assert result.ready, result.problems


def test_source_hash_never_changes_during_dry_runs(
    source_workbook_path, config
):
    before = _hash(source_workbook_path)
    for _ in range(2):
        result = preflight(source_workbook_path, config, mode="open")
        assert result.ready
    after = _hash(source_workbook_path)
    assert before == after, "A read-only dry-run changed the source bytes"


def test_second_execution_is_refused_by_idempotency(
    source_workbook_path, config, tmp_path, execute_approved
):
    """After one successful run, re-running must stop before any edit."""
    from src.workflow import run_execute

    first_exit = run_execute(
        source_workbook_path, config, mode="open", project_root=tmp_path
    )
    assert first_exit == 0
    published = sorted((tmp_path / "output").glob("*.xlsb"))
    assert len(published) == 1
    updated_workbook = published[0]
    fingerprint_before = _hash(updated_workbook)
    # Re-running against the produced A08 workbook must refuse before editing.
    second_exit = run_execute(
        updated_workbook, config, mode="open", project_root=tmp_path
    )
    assert second_exit != 0
    assert _hash(updated_workbook) == fingerprint_before
