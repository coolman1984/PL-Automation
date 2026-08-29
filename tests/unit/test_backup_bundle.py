from __future__ import annotations

import json

from src.backup_bundle import create_backup_bundle


def test_backup_bundle_is_byte_identical_and_has_manifest(tmp_path):
    source = tmp_path / "protected.xlsb"
    source.write_bytes(b"<## NASCA DRM FILE - VER1.00 ##>\x00payload")

    bundle = create_backup_bundle(source, tmp_path / "backups", reason="test")

    assert bundle.backup_file.read_bytes() == source.read_bytes()
    assert bundle.source_sha256 == bundle.backup_sha256
    manifest = json.loads(bundle.manifest_file.read_text(encoding="utf-8"))
    assert manifest["verified"] is True
    assert manifest["reason"] == "test"
    assert manifest["file_probe"]["protection"] == "nasca_drm"
    assert manifest["source"]["sha256"] == manifest["backup"]["sha256"]


def test_backup_bundle_never_overwrites_existing_bundle(tmp_path):
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"PK\x03\x04not-a-real-package")
    first = create_backup_bundle(source, tmp_path / "backups")
    second = create_backup_bundle(source, tmp_path / "backups")
    assert first.directory != second.directory
    assert first.backup_file.exists()
    assert second.backup_file.exists()
