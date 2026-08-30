from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool


def test_declared_range_request_is_dry_run_before_write():
    engine = FakeEngine({"Data": {}})
    request = ToolRequest.new(
        "write_range",
        target=TargetRef("working-copy", sheet="Data", address="A1:B2"),
        arguments={"values": [[1, 2], [3, 4]]},
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "tool_not_available"
    assert not any(call[0] == "write_values" for call in engine.calls)

