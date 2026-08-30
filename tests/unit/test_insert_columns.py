"""Task 10: controlled column insertion — exact anchor, exact count."""

from __future__ import annotations

from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool
from src.tool_registry import describe_tool


def test_insert_columns_is_declared_available_and_requires_backup():
    spec = describe_tool("insert_columns")
    assert spec is not None
    assert spec["status"] == "available"
    assert spec["requires_backup"] is True


def test_insert_columns_dry_run_never_mutates():
    engine = FakeEngine({"Data": {"1,4": "D-value", "1,5": "E-value"}})
    request = ToolRequest.new(
        "insert_columns",
        target=TargetRef("working-copy", sheet="Data", address="D1"),
        arguments={"count": 2, "expected_anchor_column": "D"},
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is False
    assert "Dry run" in result.warnings[0]
    assert engine.read_values(TargetRef("working-copy", sheet="Data", address="D1")) == [["D-value"]]
    assert not any(call[0] == "insert_columns" for call in engine.calls)


def test_insert_columns_refuses_wrong_anchor():
    engine = FakeEngine({"Data": {"1,4": "D-value"}})
    request = ToolRequest.new(
        "insert_columns",
        target=TargetRef("working-copy", sheet="Data", address="D1"),
        arguments={"count": 1, "expected_anchor_column": "E"},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "anchor_column_mismatch"
    assert not any(call[0] == "insert_columns" for call in engine.calls)


def test_insert_columns_shifts_existing_content_right():
    engine = FakeEngine({"Data": {"1,4": "D-value", "1,5": "E-value"}})
    request = ToolRequest.new(
        "insert_columns",
        target=TargetRef("working-copy", sheet="Data", address="D1"),
        arguments={"count": 2, "expected_anchor_column": "D"},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is True
    assert result.after_evidence["inserted_count"] == 2
    # D and E shift right by 2 -> F and G; the new D:E is blank.
    assert engine.read_values(TargetRef("working-copy", sheet="Data", address="D1")) == [[None]]
    assert engine.read_values(TargetRef("working-copy", sheet="Data", address="F1")) == [["D-value"]]
    assert engine.read_values(TargetRef("working-copy", sheet="Data", address="G1")) == [["E-value"]]


def test_insert_columns_refuses_a_source_target():
    engine = FakeEngine({"Data": {"1,4": "D-value"}})
    request = ToolRequest.new(
        "insert_columns",
        target=TargetRef("source", sheet="Data", address="D1"),
        arguments={"count": 1, "expected_anchor_column": "D"},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "invalid_target_workbook"
