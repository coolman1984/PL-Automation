from pathlib import Path

from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool


def test_inspect_file_is_read_only(tmp_path: Path):
    source = tmp_path / "sample.txt"
    source.write_bytes(b"not an excel file")
    request = ToolRequest.new("inspect_file", arguments={"file": str(source)})

    result = execute_tool(request)

    assert result.ok is True
    assert result.changed is False
    assert result.after_evidence["recognized"] is False


def test_backup_tool_dry_run_never_creates_backup(tmp_path: Path):
    source = tmp_path / "sample.xlsx"
    source.write_bytes(b"PK\x03\x04")
    request = ToolRequest.new("create_backup", arguments={"file": str(source)})

    result = execute_tool(request, project_root=tmp_path)

    assert result.ok is True
    assert result.changed is False
    assert not (tmp_path / "backups").exists()


def test_planned_tool_is_never_executed():
    request = ToolRequest.new(
        "write_range",
        target=TargetRef("working-copy", sheet="Data", address="A1"),
        arguments={"values": [[1]]},
        dry_run=False,
    )

    result = execute_tool(request, engine=FakeEngine({"Data": {}}))

    assert result.ok is False
    assert result.error.code == "tool_not_available"


def test_available_read_only_preparation_tools_have_safe_dry_run_handlers(tmp_path: Path):
    source = tmp_path / "sample.xlsb"
    source.write_bytes(b"not an excel file")

    for tool in ("snapshot_workbook", "prepare_workbook"):
        result = execute_tool(
            ToolRequest.new(tool, arguments={"file": str(source)}),
            project_root=tmp_path,
        )
        assert result.ok is True
        assert result.changed is False
        assert "Dry run" in result.warnings[0]
