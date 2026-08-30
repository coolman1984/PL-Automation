"""Safe source/working/final workbook transaction."""

from __future__ import annotations

import hashlib
import re
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path

from .constants import XLSB_EXTENSION
from .errors import CopyCreationError, PublicationError, SourceChangedError, WorkbookFormatError, WorkbookNotFoundError
from .models import RunPaths, WorkbookFingerprint


def sha256_file(path: Path, chunk_size: int = 4 * 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def make_run_id(actual_version: str, now: datetime | None = None) -> str:
    timestamp = (now or datetime.now(timezone.utc)).astimezone().strftime("%Y-%m-%d_%H%M%S")
    return f"{timestamp}_{actual_version}_{uuid.uuid4().hex[:8]}"


def _safe_stem(stem: str) -> str:
    result = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", stem).strip(" .")
    return result or "workbook"


def create_run_paths(project_root: Path, source_path: Path, actual_version: str) -> RunPaths:
    run_id = make_run_id(actual_version)
    run_dir = project_root / "work" / run_id
    run_dir.mkdir(parents=True, exist_ok=False)
    output_dir = project_root / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    failed_dir = project_root / "failed_runs"
    failed_dir.mkdir(parents=True, exist_ok=True)
    stem = _safe_stem(source_path.stem)
    working_path = run_dir / f"{stem}__{actual_version}_WORKING.xlsb"
    final_path = output_dir / f"{stem}__{actual_version}_UPDATED.xlsb"
    return RunPaths(
        run_id=run_id,
        run_dir=run_dir,
        working_path=working_path,
        report_path=run_dir / "run_report.txt",
        manifest_path=run_dir / "run_manifest.json",
        final_path=final_path,
    )


def assert_source_candidate(path: Path) -> None:
    if not path.exists():
        raise WorkbookNotFoundError(f"Workbook was not found: {path}")
    if not path.is_file():
        raise WorkbookFormatError(f"Workbook path is not a file: {path}")
    if path.suffix.casefold() != XLSB_EXTENSION:
        raise WorkbookFormatError(
            f"Phase 1 requires an Excel binary workbook (.xlsb): {path.name}"
        )


def save_working_copy(source_workbook: object, working_path: Path) -> None:
    try:
        source_workbook.SaveCopyAs(str(working_path))
    except Exception as exc:  # pragma: no cover - requires Excel
        raise CopyCreationError(f"Excel SaveCopyAs failed for {working_path}: {exc}") from exc
    if not working_path.exists() or working_path.stat().st_size == 0:
        raise CopyCreationError(f"Excel did not create a usable working copy: {working_path}")
    if working_path.suffix.casefold() != XLSB_EXTENSION:
        raise WorkbookFormatError(f"Working copy is not .xlsb: {working_path}")


def assert_source_unchanged(before: WorkbookFingerprint, source_path: Path) -> None:
    if not source_path.exists():
        raise SourceChangedError(f"Source workbook disappeared: {source_path}")
    current_stat = source_path.stat()
    current_hash = sha256_file(source_path) if before.sha256 else None
    if (
        current_stat.st_size != before.size_bytes
        or current_hash != before.sha256
    ):
        raise SourceChangedError(
            "The original workbook changed during the run",
            evidence={"before_sha256": before.sha256, "after_sha256": current_hash},
        )


def _collision_safe_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 1000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise PublicationError(f"Could not find a free output filename near {path}")


def publish_validated_workbook(closed_working_path: Path, final_path: Path) -> str:
    if not closed_working_path.exists():
        raise PublicationError(f"Closed validated working file is missing: {closed_working_path}")
    if closed_working_path.suffix.casefold() != XLSB_EXTENSION:
        raise PublicationError("Only a closed .xlsb working file may be published")
    final_path = _collision_safe_path(final_path)
    try:
        shutil.copy2(closed_working_path, final_path)
    except Exception as exc:
        raise PublicationError(f"Could not publish final workbook: {exc}") from exc
    if sha256_file(closed_working_path) != sha256_file(final_path):
        raise PublicationError("Published workbook hash differs from validated working copy")
    return str(final_path)


def publish_validated_file(closed_working_path: Path, final_path: Path) -> str:
    """Publish any validated Excel-format working copy with hash verification.

    The existing XLSB-only helper remains unchanged for the P&L recipe.  This
    generic helper is deliberately limited to Excel extensions and still
    requires the caller to prove that the file is closed and validated.
    """
    allowed = {".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xls"}
    closed_working_path = Path(closed_working_path)
    final_path = Path(final_path)
    if not closed_working_path.exists() or not closed_working_path.is_file():
        raise PublicationError(f"Validated working file is missing: {closed_working_path}")
    if closed_working_path.suffix.casefold() not in allowed:
        raise PublicationError(f"Unsupported Excel output extension: {closed_working_path.suffix}")
    if final_path.suffix.casefold() != closed_working_path.suffix.casefold():
        raise PublicationError("Output extension must match the validated working file")
    final_path = _collision_safe_path(final_path)
    try:
        shutil.copy2(closed_working_path, final_path)
    except Exception as exc:
        raise PublicationError(f"Could not publish final Excel file: {exc}") from exc
    if sha256_file(closed_working_path) != sha256_file(final_path):
        raise PublicationError("Published Excel file hash differs from validated working copy")
    return str(final_path)


def retain_failed_workbook(paths: RunPaths) -> Path | None:
    if not paths.working_path.exists():
        return None
    failed_dir = paths.run_dir.parent.parent / "failed_runs" / paths.run_id
    failed_dir.mkdir(parents=True, exist_ok=True)
    target = failed_dir / paths.working_path.name
    try:
        shutil.copy2(paths.working_path, target)
    except Exception:
        return None
    return target
