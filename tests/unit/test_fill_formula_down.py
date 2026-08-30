"""Task 9: bounded formula fill-down — template fingerprint, exact row count."""

from __future__ import annotations

from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool
from src.tool_registry import describe_tool


def _engine_with_template():
    return FakeEngine(
        {"DB": {"21,1": None, "21,2": None}},
        {"DB": {"21,1": "=A20+1", "21,2": "=B20*2"}},
    )


def test_fill_formula_down_is_declared_available_and_requires_backup():
    spec = describe_tool("fill_formula_down")
    assert spec is not None
    assert spec["status"] == "available"
    assert spec["requires_backup"] is True


def test_fill_formula_down_dry_run_never_mutates():
    engine = _engine_with_template()
    request = ToolRequest.new(
        "fill_formula_down",
        target=TargetRef("working-copy", sheet="DB", address="A22:B23"),
        arguments={
            "template": {"workbook_id": "working-copy", "sheet": "DB", "address": "A21:B21"},
            "expected_template_formulas": ["=A20+1", "=B20*2"],
            "expected_target_row_count": 2,
        },
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is False
    assert "Dry run" in result.warnings[0]
    assert not any(call[0] == "fill_formula_down" for call in engine.calls)


def test_fill_formula_down_refuses_wrong_template_fingerprint():
    engine = _engine_with_template()
    request = ToolRequest.new(
        "fill_formula_down",
        target=TargetRef("working-copy", sheet="DB", address="A22:B23"),
        arguments={
            "template": {"workbook_id": "working-copy", "sheet": "DB", "address": "A21:B21"},
            "expected_template_formulas": ["=A20+999", "=B20*2"],
            "expected_target_row_count": 2,
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "template_fingerprint_mismatch"
    assert not any(call[0] == "fill_formula_down" for call in engine.calls)


def test_fill_formula_down_refuses_row_count_mismatch():
    engine = _engine_with_template()
    request = ToolRequest.new(
        "fill_formula_down",
        target=TargetRef("working-copy", sheet="DB", address="A22:B23"),
        arguments={
            "template": {"workbook_id": "working-copy", "sheet": "DB", "address": "A21:B21"},
            "expected_template_formulas": ["=A20+1", "=B20*2"],
            "expected_target_row_count": 99,
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "row_count_mismatch"


def test_fill_formula_down_refuses_non_contiguous_target():
    engine = _engine_with_template()
    request = ToolRequest.new(
        "fill_formula_down",
        target=TargetRef("working-copy", sheet="DB", address="A25:B26"),
        arguments={
            "template": {"workbook_id": "working-copy", "sheet": "DB", "address": "A21:B21"},
            "expected_template_formulas": ["=A20+1", "=B20*2"],
            "expected_target_row_count": 2,
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "not_contiguous"


def test_fill_formula_down_refuses_column_mismatch():
    engine = _engine_with_template()
    request = ToolRequest.new(
        "fill_formula_down",
        target=TargetRef("working-copy", sheet="DB", address="B22:C23"),
        arguments={
            "template": {"workbook_id": "working-copy", "sheet": "DB", "address": "A21:B21"},
            "expected_template_formulas": ["=A20+1", "=B20*2"],
            "expected_target_row_count": 2,
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "column_mismatch"


def test_fill_formula_down_fills_and_preserves_the_template_row():
    engine = _engine_with_template()
    template = TargetRef("working-copy", sheet="DB", address="A21:B21")
    target = TargetRef("working-copy", sheet="DB", address="A22:B23")
    request = ToolRequest.new(
        "fill_formula_down",
        target=target,
        arguments={
            "template": {"workbook_id": "working-copy", "sheet": "DB", "address": "A21:B21"},
            "expected_template_formulas": ["=A20+1", "=B20*2"],
            "expected_target_row_count": 2,
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is True
    assert result.after_evidence["filled_row_count"] == 2
    # The fake only proves bounded dispatch. Relative-reference shifting is
    # deliberately accepted only from the real Excel integration test.
    assert any(call[0] == "fill_formula_down" for call in engine.calls)
    # The template row itself must be untouched.
    assert engine.read_formulas(template) == [["=A20+1", "=B20*2"]]
