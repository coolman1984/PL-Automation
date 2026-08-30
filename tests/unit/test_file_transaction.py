"""Unit tests for the pure filesystem pieces of the transaction."""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from src.errors import PublicationError, WorkbookFormatError
from src.file_transaction import (
    create_run_paths,
    make_run_id,
    publish_validated_file,
    publish_validated_workbook,
    save_working_copy,
)


def test_make_run_id_format():
    run_id = make_run_id("A08")
    assert re.fullmatch(r"\d{4}-\d{2}-\d{2}_\d{6}_A08_[0-9a-f]{8}", run_id)


def test_create_run_paths_builds_expected_layout(tmp_path):
    source = tmp_path / "★Final PL Statement S08 T09 V4(1).xlsb"
    paths = create_run_paths(tmp_path, source, "A08")
    assert paths.run_dir.is_dir()
    assert paths.run_dir.parent.name == "work"
    assert (tmp_path / "output").is_dir()
    assert (tmp_path / "failed_runs").is_dir()
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", source.stem).strip(" .")
    assert paths.working_path.name == f"{safe_name}__A08_WORKING.xlsb"
    assert paths.final_path.name == f"{safe_name}__A08_UPDATED.xlsb"
    assert paths.report_path.parent == paths.run_dir
    assert paths.manifest_path.parent == paths.run_dir


def test_publish_validated_workbook_roundtrip_matches_hash(tmp_path):
    working = tmp_path / "work.xlsb"
    payload = b"deterministic workbook bytes"
    working.write_bytes(payload)
    final = tmp_path / "final.xlsb"
    published = publish_validated_workbook(working, final)
    assert Path(published) == final
    assert final.read_bytes() == payload


def test_publish_collision_guard_picks_free_name(tmp_path):
    working = tmp_path / "work.xlsb"
    working.write_bytes(b"payload")
    existing = tmp_path / "target.xlsb"
    existing.write_bytes(b"older-run")
    published = publish_validated_workbook(working, existing)
    assert Path(published).name.startswith("target_2")
    assert existing.read_bytes() == b"older-run"


def test_publish_refuses_missing_destination_parent(tmp_path):
    working = tmp_path / "work.xlsb"
    working.write_bytes(b"payload")
    # shutil.copy2 does not create parent directories, so a missing parent is a
    # hard publication error rather than something to silently paper over.
    missing_parent = tmp_path / "does-not-exist-dir" / "out.xlsb"
    with pytest.raises(PublicationError):
        publish_validated_workbook(working, missing_parent)


def test_publish_refuses_missing_working_file(tmp_path):
    missing = tmp_path / "does-not-exist.xlsb"
    with pytest.raises(PublicationError):
        publish_validated_workbook(missing, tmp_path / "out.xlsb")


def test_generic_publish_supports_xlsx_and_requires_matching_extension(tmp_path):
    working = tmp_path / "work.xlsx"
    working.write_bytes(b"xlsx bytes")
    final = tmp_path / "final.xlsx"

    published = publish_validated_file(working, final)

    assert Path(published) == final
    assert final.read_bytes() == b"xlsx bytes"
    with pytest.raises(PublicationError):
        publish_validated_file(working, tmp_path / "wrong.xlsb")


def test_generic_working_copy_accepts_xlsx_without_leaving_an_orphan_error(tmp_path):
    working = tmp_path / "working.xlsx"

    class Workbook:
        def SaveCopyAs(self, path):
            Path(path).write_bytes(b"xlsx bytes")

    save_working_copy(Workbook(), working)

    assert working.read_bytes() == b"xlsx bytes"


def test_unsupported_working_copy_extension_is_rejected_before_excel_writes(tmp_path):
    calls = []

    class Workbook:
        def SaveCopyAs(self, path):
            calls.append(path)

    with pytest.raises(WorkbookFormatError, match="supported Excel format"):
        save_working_copy(Workbook(), tmp_path / "working.txt")

    assert calls == []
