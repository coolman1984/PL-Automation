"""Dependency-free Excel container/protection probe and safe route selection.

The probe never opens Excel and never writes to the source.  It inspects only
the file signature and, for normal Office packages, the ZIP member names.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path


NASCA_PREFIX = b"<## NASCA DRM FILE - VER1.00 ##>"
OLE_PREFIX = bytes.fromhex("D0CF11E0A1B11AE1")
ZIP_PREFIXES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
_ENCRYPTED_PACKAGE_UTF16 = "EncryptedPackage".encode("utf-16le")
_ENCRYPTION_INFO_UTF16 = "EncryptionInfo".encode("utf-16le")


@dataclass(frozen=True)
class FileProbeResult:
    path: str
    extension: str
    container: str
    workbook_format: str
    protection: str
    recognized: bool
    recommended_engine: str
    recommended_com_mode: str | None
    fast_read_candidate: bool
    fast_edit_candidate: bool
    reason: str
    warnings: tuple[str, ...] = field(default_factory=tuple)

    @property
    def requires_manual_excel_open(self) -> bool:
        return self.recommended_com_mode == "attach"


def _read_prefix(path: Path, size: int = 4 * 1024 * 1024) -> bytes:
    with path.open("rb") as handle:
        return handle.read(size)


def _extension_warning(extension: str, workbook_format: str) -> tuple[str, ...]:
    expected = {
        "xlsb": ".xlsb",
        "xlsx": ".xlsx",
        "xlsm": ".xlsm",
    }.get(workbook_format)
    if expected and extension != expected:
        return (
            f"Filename extension {extension or '<none>'} does not match detected "
            f"workbook format {workbook_format}",
        )
    return ()


def _probe_zip_package(path: Path, extension: str) -> FileProbeResult:
    try:
        with zipfile.ZipFile(path, "r") as package:
            names = set(package.namelist())
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        return FileProbeResult(
            str(path), extension, "zip", "unknown", "unknown", False,
            "manual_review", None, False, False,
            f"Office package could not be inspected safely: {exc}",
        )
    lower_names = {name.casefold() for name in names}
    if "xl/workbook.bin" in lower_names:
        warnings = _extension_warning(extension, "xlsb")
        return FileProbeResult(
            str(path), extension, "opc_zip", "xlsb", "none", True,
            "excel_com", "open", True, False,
            "Normal unprotected XLSB package. Fast readers are suitable for "
            "inspection, but safe full-fidelity editing remains assigned to Excel.",
            warnings,
        )
    if "xl/workbook.xml" in lower_names:
        workbook_format = "xlsm" if "xl/vbaproject.bin" in lower_names else "xlsx"
        warnings = list(_extension_warning(extension, workbook_format))
        risky_prefixes = (
            "xl/activex/",
            "xl/charts/",
            "xl/drawings/",
            "xl/embeddings/",
            "xl/externallinks/",
            "xl/pivotcache/",
            "xl/pivottables/",
            "xl/slicers/",
            "xl/timelines/",
        )
        risky_exact = {"xl/connections.xml", "xl/vbaproject.bin"}
        risky_parts = sorted(
            name for name in lower_names
            if name in risky_exact or name.startswith(risky_prefixes)
        )
        fast_candidate = workbook_format == "xlsx" and not risky_parts
        if risky_parts:
            warnings.append(
                "Complex workbook parts detected; fast editing is disabled: "
                + ", ".join(risky_parts[:10])
            )
        return FileProbeResult(
            str(path), extension, "opc_zip", workbook_format, "none", True,
            "fast_ooxml_candidate" if fast_candidate else "excel_com",
            None if fast_candidate else "open",
            True,
            fast_candidate,
            "Normal XML workbook package. It may use a fast non-COM engine only "
            "after a workbook-feature gate proves the file is simple and supported.",
            tuple(warnings),
        )
    return FileProbeResult(
        str(path), extension, "zip", "unknown", "none", False,
        "manual_review", None, False, False,
        "ZIP file does not expose an Excel workbook part",
    )


def probe_excel_file(path: Path) -> FileProbeResult:
    path = Path(path)
    extension = path.suffix.casefold()
    prefix = _read_prefix(path)

    if prefix.startswith(NASCA_PREFIX):
        return FileProbeResult(
            str(path), extension, "nasca_wrapper", "excel_workbook", "nasca_drm",
            True, "excel_com", "attach", False, False,
            "NASCA DRM signature detected. Open the exact file manually in "
            "authorized desktop Excel, then attach to that session.",
        )

    if prefix.startswith(ZIP_PREFIXES):
        return _probe_zip_package(path, extension)

    if prefix.startswith(OLE_PREFIX):
        encrypted = (
            _ENCRYPTED_PACKAGE_UTF16 in prefix
            or _ENCRYPTION_INFO_UTF16 in prefix
        )
        if encrypted:
            return FileProbeResult(
                str(path), extension, "ole_compound", "encrypted_ooxml",
                "office_encrypted", True, "excel_com", "attach", False, False,
                "Microsoft Office encrypted package detected. Manual authorized "
                "Excel opening is required before attachment.",
            )
        return FileProbeResult(
            str(path), extension, "ole_compound", "legacy_xls_or_unknown",
            "unknown", True, "excel_com", "open", False, False,
            "OLE workbook/container detected. This project does not rewrite it "
            "outside desktop Excel.",
        )

    return FileProbeResult(
        str(path), extension, "unknown", "unknown", "unknown", False,
        "manual_review", None, False, False,
        "File signature is not a recognized Excel container",
    )


def resolve_com_mode(requested_mode: str, probe: FileProbeResult) -> str:
    """Honor explicit user modes; make auto deterministic from the probe."""
    if requested_mode != "auto":
        return requested_mode
    return probe.recommended_com_mode or "auto"


def render_probe_report(probe: FileProbeResult) -> str:
    lines = [
        "EXCEL FILE QUICK CHECK",
        f"File: {probe.path}",
        f"Container: {probe.container}",
        f"Detected format: {probe.workbook_format}",
        f"Protection: {probe.protection}",
        f"Recommended engine: {probe.recommended_engine}",
        f"Recommended Excel mode: {probe.recommended_com_mode or 'not required/undecided'}",
        f"Fast read candidate: {'YES' if probe.fast_read_candidate else 'NO'}",
        f"Fast edit candidate: {'YES' if probe.fast_edit_candidate else 'NO'}",
        f"Decision: {probe.reason}",
    ]
    if probe.warnings:
        lines.append("Warnings:")
        lines.extend(f"- {warning}" for warning in probe.warnings)
    return "\n".join(lines) + "\n"
