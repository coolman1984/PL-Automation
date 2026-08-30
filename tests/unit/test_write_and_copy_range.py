"""Task 8: bulk write_range and copy_range — shape equality, evidence, caps."""

from __future__ import annotations

from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool
from src.tool_registry import describe_tool


def test_write_range_and_copy_range_are_now_available():
    for name in ("write_range", "copy_range"):
        spec = describe_tool(name)
        assert spec is not None
        assert spec["status"] == "available"
        assert spec["requires_backup"] is True


def test_write_range_dry_run_reports_before_evidence_without_writing():
    engine = FakeEngine({"Data": {"1,1": "old"}})
    request = ToolRequest.new(
        "write_range",
        target=TargetRef("working-copy", sheet="Data", address="A1"),
        arguments={"values": [["new"]]},
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is False
    assert result.before_evidence["values"] == [["old"]]
    assert engine.read_values(TargetRef("working-copy", sheet="Data", address="A1")) == [["old"]]
    assert not any(call[0] == "write_values" for call in engine.calls)


def test_write_range_executes_and_captures_rollback_evidence():
    target = TargetRef("working-copy", sheet="Data", address="A1")
    engine = FakeEngine({"Data": {"1,1": "old"}})
    request = ToolRequest.new(
        "write_range",
        target=target,
        arguments={"values": [["new"]]},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is True
    assert result.before_evidence["values"] == [["old"]]
    assert engine.read_values(target) == [["new"]]

    engine.write_values(target, result.before_evidence["values"])
    assert engine.read_values(target) == [["old"]]


def test_write_range_rejects_a_payload_over_the_cell_limit(monkeypatch):
    import src.tool_executor as tool_executor_module

    monkeypatch.setattr(tool_executor_module, "_MAX_CELLS_PER_RANGE_OPERATION", 2)
    engine = FakeEngine({"Data": {}})
    request = ToolRequest.new(
        "write_range",
        target=TargetRef("working-copy", sheet="Data", address="A1:C1"),
        arguments={"values": [[1, 2, 3]]},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "range_too_large"
    assert not any(call[0] == "write_values" for call in engine.calls)


def test_copy_range_requires_matching_shape():
    engine = FakeEngine({"Data": {"1,1": "a", "1,2": "b"}})
    request = ToolRequest.new(
        "copy_range",
        target=TargetRef("working-copy", sheet="Data", address="A2:B2"),
        arguments={"source": {"workbook_id": "working-copy", "sheet": "Data", "address": "A1:A1"}},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "shape_mismatch"
    assert not any(call[0].startswith("copy_range") for call in engine.calls)


def test_copy_range_copies_values_and_captures_rollback_evidence():
    engine = FakeEngine({"Data": {"1,1": "source-value", "2,1": "old-dest"}})
    destination = TargetRef("working-copy", sheet="Data", address="A2")
    request = ToolRequest.new(
        "copy_range",
        target=destination,
        arguments={
            "source": {"workbook_id": "working-copy", "sheet": "Data", "address": "A1"},
            "mode": "values",
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is True
    assert result.before_evidence["values"] == [["old-dest"]]
    assert result.after_evidence["copied_cell_count"] == 1
    assert engine.read_values(destination) == [["source-value"]]

    engine.write_values(destination, result.before_evidence["values"])
    assert engine.read_values(destination) == [["old-dest"]]


def test_copy_range_refuses_a_genuinely_foreign_workbook_id():
    engine = FakeEngine({"Data": {"1,1": "a"}})
    request = ToolRequest.new(
        "copy_range",
        target=TargetRef("working-copy", sheet="Data", address="A2"),
        arguments={"source": {"workbook_id": "some_other_open_workbook.xlsx", "sheet": "Data", "address": "A1"}},
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "cross_workbook_copy_unsupported"
