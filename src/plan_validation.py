"""Deterministic checks applied before an agent operation plan can run."""

from __future__ import annotations

from .agent_contracts import OperationPlan
from .tool_registry import describe_tool


def validate_plan(plan: OperationPlan) -> list[str]:
    """Return all blocking plan errors; an empty list means structurally valid."""
    errors: list[str] = []
    if not plan.steps:
        errors.append("plan must contain at least one step")
    seen_ids: set[str] = set()
    for step in plan.steps:
        if step.step_id in seen_ids:
            errors.append(f"duplicate step_id: {step.step_id}")
        seen_ids.add(step.step_id)
        if step.tool != step.request.tool:
            errors.append(
                f"step {step.step_id} declares {step.tool} but request dispatches "
                f"{step.request.tool}"
            )
        spec = describe_tool(step.tool)
        if spec is None:
            errors.append(f"unknown tool: {step.tool}")
            continue
        if spec["status"] != "available":
            errors.append(f"tool is locked: {step.tool}")
        if step.request.transaction_id != plan.transaction_id:
            errors.append(f"step {step.step_id} uses a different transaction_id")
        if spec["mutates_workbook"] and not spec["requires_backup"]:
            errors.append(f"mutating tool is missing backup requirement: {step.tool}")
        if spec["mutates_workbook"] and not plan.requires_approval:
            errors.append(f"mutating execution requires explicit approval: {step.tool}")
        if spec["mutates_workbook"] and step.request.expected_effect.get("changed") is not True:
            errors.append(
                f"mutating step must explicitly expect changed=true: {step.tool}"
            )
    if plan.unresolved:
        errors.append("plan contains unresolved items")
    return errors


def available_tool_names() -> tuple[str, ...]:
    """Expose names for planners without exposing executable handlers."""
    from .tool_registry import tool_catalog

    return tuple(item["name"] for item in tool_catalog(include_planned=False)["tools"])
