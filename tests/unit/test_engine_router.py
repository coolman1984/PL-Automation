from pathlib import Path

from src.engines.router import decide_engine
from src.file_probe import probe_excel_file


def test_router_stops_unknown_container(tmp_path: Path):
    path = tmp_path / "unknown.bin"
    path.write_bytes(b"not excel")

    decision = decide_engine(probe_excel_file(path))

    assert decision.engine == "none"
    assert decision.mode == "stop"


def test_router_selects_com_for_nasca_signature(tmp_path: Path):
    path = tmp_path / "protected.xlsb"
    path.write_bytes(b"<## NASCA DRM FILE - VER1.00 ##>payload")

    decision = decide_engine(probe_excel_file(path))

    assert decision.engine == "excel_com"
    assert decision.mode == "attach"
    assert decision.capabilities.supports_xlsb

