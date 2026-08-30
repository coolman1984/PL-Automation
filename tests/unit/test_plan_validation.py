from src.agent_contracts import OperationPlan, PlanStep, TargetRef, ToolRequest
from src.plan_validation import validate_plan


def test_plan_validation_rejects_locked_tool_before_execution():
    transaction_id = "run-1"
    request = ToolRequest(
        schema_version="1.0",
        transaction_id=transaction_id,
        tool="publish_workbook",
        target=TargetRef("working-copy", sheet="Data", address="A1"),
        arguments={},
        dry_run=True,
    )
    plan = OperationPlan(
        transaction_id,
        "set one formula",
        steps=(PlanStep("step-1", "publish_workbook", "write", request),),
    )

    errors = validate_plan(plan)

    assert any("locked" in error for error in errors)


def test_plan_validation_accepts_read_only_inspection():
    transaction_id = "run-2"
    request = ToolRequest(
        schema_version="1.0",
        transaction_id=transaction_id,
        tool="inspect_file",
        arguments={"file": "book.xlsx"},
        dry_run=True,
    )
    plan = OperationPlan(
        transaction_id,
        "inspect workbook",
        steps=(PlanStep("step-1", "inspect_file", "probe", request),),
    )

    assert validate_plan(plan) == []


def test_plan_validation_rejects_declared_and_dispatched_tool_mismatch():
    transaction_id = "run-mismatch"
    request = ToolRequest(
        schema_version="1.0",
        transaction_id=transaction_id,
        tool="clear_range",
        target=TargetRef("working-copy", sheet="Data", address="A1"),
        arguments={"expected_cell_count": 1},
        dry_run=False,
    )
    plan = OperationPlan(
        transaction_id,
        "mismatched dispatch",
        requires_approval=True,
        steps=(PlanStep("step-1", "read_range", "unsafe mismatch", request),),
    )

    errors = validate_plan(plan)

    assert any("declares read_range but request dispatches clear_range" in error for error in errors)
