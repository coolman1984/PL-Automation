"""Pre/post-edit workbook fingerprints and preservation checks."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .constants import XL_EXCEL12
from .file_transaction import sha256_file
from .models import ValidationCheck, WorkbookFingerprint


def get_sheet_names(workbook: object) -> list[str]:
    return [str(workbook.Sheets(index).Name) for index in range(1, int(workbook.Sheets.Count) + 1)]


def get_external_links(workbook: object) -> list[str]:
    try:
        links = workbook.LinkSources(1)  # xlLinkTypeExcelLinks
    except Exception:
        return []
    if links is None:
        return []
    if isinstance(links, str):
        return [links]
    try:
        return [str(item) for item in links]
    except TypeError:
        return [str(links)]


def get_pivot_counts(workbook: object) -> dict[str, int]:
    counts: dict[str, int] = {}
    for name in get_sheet_names(workbook):
        worksheet = workbook.Worksheets(name)
        try:
            counts[name] = int(worksheet.PivotTables().Count)
        except Exception:
            # A sheet with no PivotTables may raise under dynamic dispatch.
            counts[name] = 0
    return counts


def get_connection_count(workbook: object) -> int | None:
    try:
        return int(workbook.Connections.Count)
    except Exception:
        return None


def detect_vba_project(workbook: object) -> bool | None:
    try:
        _ = workbook.VBProject
        return True
    except Exception as exc:
        text = str(exc).lower()
        if "programmatic access" in text or "access denied" in text or "trust" in text:
            return None
        return False


def collect_fingerprint(workbook: object, path: Path) -> WorkbookFingerprint:
    stat = path.stat()
    try:
        file_format = workbook.FileFormat
    except Exception:
        file_format = None
    return WorkbookFingerprint(
        path=str(path.resolve()),
        file_name=path.name,
        size_bytes=int(stat.st_size),
        modified_utc=datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat(),
        sha256=sha256_file(path),
        file_format=file_format,
        sheet_names=get_sheet_names(workbook),
        sheet_count=int(workbook.Sheets.Count),
        external_links=get_external_links(workbook),
        defined_name_count=_safe_count(workbook, "Names"),
        pivot_counts=get_pivot_counts(workbook),
        connection_count=get_connection_count(workbook),
        has_vba_project=detect_vba_project(workbook),
    )


def _safe_count(obj: object, attribute: str) -> int | None:
    try:
        return int(getattr(obj, attribute).Count)
    except Exception:
        return None


def compare_preservation(
    before: WorkbookFingerprint, after: WorkbookFingerprint
) -> list[ValidationCheck]:
    checks = [
        ValidationCheck(
            "sheet_names_preserved",
            before.sheet_names == after.sheet_names,
            True,
            "Sheet names are unchanged" if before.sheet_names == after.sheet_names else "Sheet names changed",
            {"before": before.sheet_names, "after": after.sheet_names},
        ),
        ValidationCheck(
            "sheet_count_preserved",
            before.sheet_count == after.sheet_count,
            True,
            "Sheet count is unchanged" if before.sheet_count == after.sheet_count else "Sheet count changed",
            {"before": before.sheet_count, "after": after.sheet_count},
        ),
        ValidationCheck(
            "xlsb_format_preserved",
            after.file_format == XL_EXCEL12,
            True,
            "Workbook remains Excel binary format" if after.file_format == XL_EXCEL12 else "Workbook format is not Excel binary",
            {"before": before.file_format, "after": after.file_format},
        ),
        ValidationCheck(
            "external_links_preserved",
            before.external_links == after.external_links,
            True,
            "External link sources are unchanged" if before.external_links == after.external_links else "External link sources changed",
            {"before": before.external_links, "after": after.external_links},
        ),
    ]
    for name, before_value, after_value in (
        ("defined_name_count_preserved", before.defined_name_count, after.defined_name_count),
        ("connection_count_preserved", before.connection_count, after.connection_count),
        ("pivot_counts_preserved", before.pivot_counts, after.pivot_counts),
        ("vba_presence_preserved", before.has_vba_project, after.has_vba_project),
    ):
        if before_value is None or after_value is None:
            checks.append(
                ValidationCheck(name, True, False, "Optional preservation fact was unavailable", {})
            )
        else:
            checks.append(
                ValidationCheck(
                    name,
                    before_value == after_value,
                    True,
                    "Optional preservation fact is unchanged" if before_value == after_value else "Optional preservation fact changed",
                    {"before": before_value, "after": after_value},
                )
            )
    return checks


def capture_control_cells(workbook: object, configured_controls: list[dict[str, Any]] | None = None) -> dict[str, object]:
    """Capture explicitly configured control cells without guessing workbook controls."""
    result: dict[str, object] = {}
    for item in configured_controls or []:
        sheet = str(item["sheet"])
        address = str(item["address"])
        key = f"{sheet}!{address}"
        result[key] = workbook.Worksheets(sheet).Range(address).Value2
    return result
