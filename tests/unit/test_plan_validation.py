from src.agent_contracts import OperationPlan, PlanStep, TargetRef, ToolRequest
from src.plan_validation import validate_plan


def test_plan_validation_rejects_locked_tool_before_execution():
    transaction_id = "run-1"
    request = ToolRequest(
        schema_version="1.0",
        transaction_id=transaction_id,
        tool="set_formula",
        target=TargetRef("working-copy", sheet="Data", address="A1"),
        arguments={"values": [[1]]},
        dry_run=True,
    )
    plan = OperationPlan(
        transaction_id,
        "set one formula",
        steps=(PlanStep("step-1", "set_formula", "write", request),),
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

