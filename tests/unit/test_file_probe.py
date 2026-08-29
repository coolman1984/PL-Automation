from __future__ import annotations

import zipfile

from src.file_probe import (
    NASCA_PREFIX,
    OLE_PREFIX,
    probe_excel_file,
    resolve_com_mode,
)


def test_nasca_signature_routes_to_manual_excel_attach(tmp_path):
    path = tmp_path / "protected.xlsb"
    path.write_bytes(NASCA_PREFIX + b"encrypted payload")

    result = probe_excel_file(path)

    assert result.protection == "nasca_drm"
    assert result.recommended_engine == "excel_com"
    assert result.recommended_com_mode == "attach"
    assert not result.fast_edit_candidate
    assert resolve_com_mode("auto", result) == "attach"


def test_normal_xlsb_routes_to_isolated_excel_for_safe_write(tmp_path):
    path = tmp_path / "normal.xlsb"
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types/>")
        package.writestr("xl/workbook.bin", b"binary workbook")

    result = probe_excel_file(path)

    assert result.workbook_format == "xlsb"
    assert result.protection == "none"
    assert result.fast_read_candidate
    assert not result.fast_edit_candidate
    assert resolve_com_mode("auto", result) == "open"


def test_normal_xlsx_is_a_fast_engine_candidate(tmp_path):
    path = tmp_path / "normal.xlsx"
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types/>")
        package.writestr("xl/workbook.xml", "<workbook/>")

    result = probe_excel_file(path)

    assert result.workbook_format == "xlsx"
    assert result.recommended_engine == "fast_ooxml_candidate"
    assert result.fast_edit_candidate


def test_complex_xlsx_is_kept_on_excel_engine(tmp_path):
    path = tmp_path / "complex.xlsx"
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types/>")
        package.writestr("xl/workbook.xml", "<workbook/>")
        package.writestr("xl/drawings/drawing1.xml", "<drawing/>")
        package.writestr("xl/externalLinks/externalLink1.xml", "<link/>")

    result = probe_excel_file(path)

    assert result.recommended_engine == "excel_com"
    assert result.recommended_com_mode == "open"
    assert not result.fast_edit_candidate
    assert result.warnings


def test_office_encrypted_ole_routes_to_attach(tmp_path):
    path = tmp_path / "encrypted.xlsx"
    path.write_bytes(
        OLE_PREFIX
        + b"padding"
        + "EncryptionInfo".encode("utf-16le")
        + "EncryptedPackage".encode("utf-16le")
    )

    result = probe_excel_file(path)

    assert result.protection == "office_encrypted"
    assert result.recommended_com_mode == "attach"


def test_explicit_mode_is_never_silently_overridden(tmp_path):
    path = tmp_path / "protected.xlsb"
    path.write_bytes(NASCA_PREFIX + b"encrypted payload")
    result = probe_excel_file(path)

    assert resolve_com_mode("open", result) == "open"
