"""Task 11: PivotTable source update and targeted refresh."""

from __future__ import annotations

from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool
from src.tool_registry import describe_tool


def _engine_with_pivot(source_type="database", source_data="'DB File'!$A$1:$C$10", shared_with=None):
    return FakeEngine(
        {"PV": {}},
        {"PV": {}},
        {
            "PV!Sales": {
                "name": "Sales",
                "sheet": "PV",
                "cache_index": 0,
                "source_type": source_type,
                "source_data": source_data,
                "shared_with": shared_with or [],
            }
        },
    )


def test_update_pivot_source_is_declared_available_and_requires_backup():
    spec = describe_tool("update_pivot_source")
    assert spec is not None
    assert spec["status"] == "available"
    assert spec["requires_backup"] is True


def test_update_pivot_source_dry_run_never_mutates():
    engine = _engine_with_pivot()
    request = ToolRequest.new(
        "update_pivot_source",
        target=TargetRef("working-copy", sheet="PV", object_name="Sales"),
        arguments={
            "expected_current_source": "'DB File'!$A$1:$C$10",
            "new_source": "'DB File'!$A$1:$C$20",
        },
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is False
    assert "Dry run" in result.warnings[0]
    assert not any(call[0] == "update_pivot_source" for call in engine.calls)


def test_update_pivot_source_refuses_current_source_mismatch():
    engine = _engine_with_pivot()
    request = ToolRequest.new(
        "update_pivot_source",
        target=TargetRef("working-copy", sheet="PV", object_name="Sales"),
        arguments={
            "expected_current_source": "'DB File'!$A$1:$C$999",
            "new_source": "'DB File'!$A$1:$C$20",
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "current_source_mismatch"


def test_update_pivot_source_refuses_unsupported_source_type():
    engine = _engine_with_pivot(source_type="external")
    request = ToolRequest.new(
        "update_pivot_source",
        target=TargetRef("working-copy", sheet="PV", object_name="Sales"),
        arguments={
            "expected_current_source": "'DB File'!$A$1:$C$10",
            "new_source": "'DB File'!$A$1:$C$20",
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "unsupported_pivot_source"


def test_update_pivot_source_refuses_shared_cache_without_explicit_ack():
    engine = _engine_with_pivot(shared_with=["PV!SalesByRegion"])
    request = ToolRequest.new(
        "update_pivot_source",
        target=TargetRef("working-copy", sheet="PV", object_name="Sales"),
        arguments={
            "expected_current_source": "'DB File'!$A$1:$C$10",
            "new_source": "'DB File'!$A$1:$C$20",
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is False
    assert result.error.code == "shared_cache_not_acknowledged"


def test_update_pivot_source_proceeds_when_shared_cache_is_acknowledged():
    engine = _engine_with_pivot(shared_with=["PV!SalesByRegion"])
    request = ToolRequest.new(
        "update_pivot_source",
        target=TargetRef("working-copy", sheet="PV", object_name="Sales"),
        arguments={
            "expected_current_source": "'DB File'!$A$1:$C$10",
            "new_source": "'DB File'!$A$1:$C$20",
            "allow_shared_cache_replacement": True,
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.changed is True
    assert result.after_evidence["source_data"] == "'DB File'!$A$1:$C$20"


def test_update_pivot_source_updates_and_refreshes():
    engine = _engine_with_pivot()
    request = ToolRequest.new(
        "update_pivot_source",
        target=TargetRef("working-copy", sheet="PV", object_name="Sales"),
        arguments={
            "expected_current_source": "'DB File'!$A$1:$C$10",
            "new_source": "'DB File'!$A$1:$C$20",
        },
        dry_run=False,
    )

    result = execute_tool(request, engine=engine)

    assert result.ok is True
    assert result.before_evidence["source_data"] == "'DB File'!$A$1:$C$10"
    assert result.after_evidence["source_data"] == "'DB File'!$A$1:$C$20"
    assert any(call == ("update_pivot_source", "PV!Sales") for call in engine.calls)
