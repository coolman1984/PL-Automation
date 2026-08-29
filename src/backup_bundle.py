"""Verified, collision-safe backup bundles for any Excel source file."""

from __future__ import annotations

import json
import os
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .file_probe import probe_excel_file
from .file_transaction import sha256_file


@dataclass(frozen=True)
class BackupBundle:
    backup_id: str
    directory: Path
    backup_file: Path
    manifest_file: Path
    source_sha256: str
    backup_sha256: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "backup_id": self.backup_id,
            "directory": str(self.directory),
            "backup_file": str(self.backup_file),
            "manifest_file": str(self.manifest_file),
            "source_sha256": self.source_sha256,
            "backup_sha256": self.backup_sha256,
        }


def _safe_name(value: str) -> str:
    cleaned = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value).strip(" .")
    return cleaned or "workbook"


def _atomic_json(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temporary, path)


def create_backup_bundle(source: Path, backup_root: Path, *, reason: str = "pre_operation") -> BackupBundle:
    source = Path(source).expanduser().resolve()
    backup_root = Path(backup_root).expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise FileNotFoundError(f"Source file was not found: {source}")

    before_stat = source.stat()
    source_hash = sha256_file(source)
    now = datetime.now(timezone.utc)
    backup_id = f"{now.strftime('%Y%m%dT%H%M%SZ')}_{uuid.uuid4().hex[:8]}"
    directory = backup_root / f"{_safe_name(source.stem)}_{backup_id}"
    directory.mkdir(parents=True, exist_ok=False)
    backup_file = directory / f"{_safe_name(source.stem)}__ORIGINAL{source.suffix}"
    partial = directory / f".{backup_file.name}.partial"

    try:
        shutil.copy2(source, partial)
        os.replace(partial, backup_file)
        backup_hash = sha256_file(backup_file)
        after_stat = source.stat()
        after_hash = sha256_file(source)
        if source_hash != backup_hash:
            raise OSError("Backup hash differs from the source hash")
        if before_stat.st_size != after_stat.st_size or source_hash != after_hash:
            raise OSError("Source file changed while the backup was being created")

        probe = probe_excel_file(source)
        manifest_file = directory / "backup_manifest.json"
        manifest = {
            "schema_version": "1.0",
            "backup_id": backup_id,
            "created_utc": now.isoformat(),
            "reason": reason,
            "verified": True,
            "source": {
                "path": str(source),
                "file_name": source.name,
                "size_bytes": before_stat.st_size,
                "modified_utc": datetime.fromtimestamp(before_stat.st_mtime, tz=timezone.utc).isoformat(),
                "sha256": source_hash,
            },
            "backup": {
                "path": str(backup_file),
                "file_name": backup_file.name,
                "size_bytes": backup_file.stat().st_size,
                "sha256": backup_hash,
            },
            "file_probe": {
                "container": probe.container,
                "workbook_format": probe.workbook_format,
                "protection": probe.protection,
                "recommended_engine": probe.recommended_engine,
                "recommended_com_mode": probe.recommended_com_mode,
            },
        }
        _atomic_json(manifest_file, manifest)
        return BackupBundle(backup_id, directory, backup_file, manifest_file, source_hash, backup_hash)
    except Exception:
        if partial.exists():
            partial.unlink()
        raise
