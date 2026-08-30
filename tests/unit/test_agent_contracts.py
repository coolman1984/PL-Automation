import json

import pytest

from src.agent_contracts import (
    OperationPlan,
    TargetRef,
    ToolError,
    ToolRequest,
    ToolResult,
    as_json,
)


def test_request_round_trip_is_json_safe_and_explicit():
    request = ToolRequest.new(
        "read_range",
        target=TargetRef("working-copy", sheet="Data", address="A1:B2"),
        arguments={"include_formats": True},
    )

    restored = ToolRequest.from_dict(json.loads(as_json(request)))

    assert restored.tool == "read_range"
    assert restored.target == request.target
    assert restored.dry_run is True


def test_request_rejects_unknown_fields_and_invalid_target():
    with pytest.raises(ValueError, match="unknown fields"):
        ToolRequest.from_dict({"tool": "read_range", "transaction_id": "x", "extra": 1})
    with pytest.raises(ValueError, match="target must identify"):
        TargetRef("workbook")


def test_result_requires_error_only_for_failures():
    result = ToolResult.failure("write_range", ToolError("blocked", "Needs approval"))

    assert result.to_dict()["error"]["code"] == "blocked"
    with pytest.raises(ValueError):
        ToolResult(ok=False, tool="write_range")


def test_plan_with_unresolved_items_requires_approval():
    with pytest.raises(ValueError, match="unresolved"):
        OperationPlan("run-1", "change workbook", unresolved=("Which sheet?",))

