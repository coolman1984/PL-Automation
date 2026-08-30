from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool


def test_locked_range_request_never_reaches_the_engine():
    engine = FakeEngine({"Data": {}})
    request = ToolRequest.new(
        "publish_workbook",
        target=TargetRef("working-copy", sheet="Data", address="A1:B2"),
        arguments={},
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "tool_not_available"
    assert not any(call[0] == "write_formulas" for call in engine.calls)


def test_declared_write_range_request_is_dry_run_before_write():
    engine = FakeEngine({"Data": {}})
    request = ToolRequest.new(
        "write_range",
        target=TargetRef("working-copy", sheet="Data", address="A1:B2"),
        arguments={"values": [[1, 2], [3, 4]]},
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is False
    assert "Dry run" in result.warnings[0]
    assert not any(call[0] == "write_values" for call in engine.calls)
