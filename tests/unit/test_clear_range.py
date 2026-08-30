"""Task 7: bounded clear_range — clear contents only, exact range, rollback."""

from __future__ import annotations

from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool
from src.tool_registry import describe_tool


def test_clear_range_is_declared_available_and_requires_backup():
    spec = describe_tool("clear_range")
    assert spec is not None
    assert spec["status"] == "available"
    assert spec["mutates_workbook"] is True
    assert spec["requires_backup"] is True


def test_clear_range_dry_run_never_mutates():
    engine = FakeEngine({"Data": {"1,1": "a", "1,2": "b"}})
    request = ToolRequest.new(
        "clear_range",
        target=TargetRef("working-copy", sheet="Data", address="A1:B1"),
        arguments={"expected_cell_count": 2},
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is False
    assert "Dry run" in result.warnings[0]
    assert engine.read_values(TargetRef("working-copy", sheet="Data", address="A1:B1")) == [["a", "b"]]
    assert not any(call[0] == "clear_range" for call in engine.calls)


def test_clear_range_refuses_on_cell_count_mismatch():
    engine = FakeEngine({"Data": {"1,1": "a", "1,2": "b"}})
    request = ToolRequest.new(
        "clear_range",
        target=TargetRef("working-copy", sheet="Data", address="A1:B1"),
        arguments={"expected_cell_count": 99},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "cell_count_mismatch"
    assert engine.read_values(TargetRef("working-copy", sheet="Data", address="A1:B1")) == [["a", "b"]]
    assert not any(call[0] == "clear_range" for call in engine.calls)


def test_clear_range_refuses_a_source_target():
    engine = FakeEngine({"Data": {"1,1": "a"}})
    request = ToolRequest.new(
        "clear_range",
        target=TargetRef("source", sheet="Data", address="A1"),
        arguments={"expected_cell_count": 1},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "invalid_target_workbook"
    assert not any(call[0] == "clear_range" for call in engine.calls)


def test_clear_range_clears_contents_and_provides_rollback_evidence():
    target = TargetRef("working-copy", sheet="Data", address="A1:B1")
    engine = FakeEngine({"Data": {"1,1": "a", "1,2": "b"}})
    request = ToolRequest.new(
        "clear_range",
        target=target,
        arguments={"expected_cell_count": 2},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is True
    assert result.before_evidence["values"] == [["a", "b"]]
    assert result.after_evidence["cleared_cell_count"] == 2
    assert result.metrics.cells_touched == 2
    assert engine.read_values(target) == [[None, None]]

    # Rollback: the recorded before-evidence is enough to restore the range.
    engine.write_values(target, result.before_evidence["values"])
    assert engine.read_values(target) == [["a", "b"]]


def test_clear_range_requires_an_engine():
    request = ToolRequest.new(
        "clear_range",
        target=TargetRef("working-copy", sheet="Data", address="A1"),
        arguments={"expected_cell_count": 1},
        dry_run=False,
    )

    result = execute_tool(request)

    assert result.ok is False
    assert result.error.code == "engine_required"
