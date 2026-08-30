"""Task 5: the generic coordinator dry-run path never mutates anything."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from src.agent_contracts import OperationPlan, PlanStep, ToolRequest, ToolResult
from src.core.coordinator import run_dry_run
from src.transaction_state import TransactionState


def _fake_snapshot_execute_tool(calls: list):
    from src.tool_executor import execute_tool as real_execute_tool

    def fake(request, **kwargs):
        if request.tool == "snapshot_workbook":
            calls.append(request)
            return ToolResult.success(
                "snapshot_workbook",
                after_evidence={
                    "snapshot": {
                        "snapshot_file": "fake_snapshot.json",
                        "snapshot_mode": "inventory",
                        "cell_count": 0,
                    }
                },
            )
        return real_execute_tool(request, **kwargs)

    return fake


def _plan(transaction_id: str, *, requires_approval: bool = False, unresolved: tuple = ()) -> OperationPlan:
    step = PlanStep(
        step_id="step-1",
        tool="inspect_file",
        purpose="probe the target workbook",
        request=ToolRequest.new("inspect_file", arguments={"file": "unused"}),
    )
    # ToolRequest.new mints its own transaction_id; align it with the plan's.
    step = PlanStep(
        step_id="step-1",
        tool="inspect_file",
        purpose="probe the target workbook",
        request=ToolRequest(
            schema_version=step.request.schema_version,
            transaction_id=transaction_id,
            tool=step.request.tool,
            arguments=step.request.arguments,
            dry_run=True,
        ),
    )
    return OperationPlan(
        transaction_id=transaction_id,
        intent="test plan",
        unresolved=unresolved,
        requires_approval=requires_approval or bool(unresolved),
        steps=(step,),
    )


def _xlsx_stub(tmp_path: Path) -> Path:
    """A minimal zip package that ``probe_excel_file`` recognizes as .xlsx."""
    source = tmp_path / "book.xlsx"
    with zipfile.ZipFile(source, "w") as package:
        package.writestr("xl/workbook.xml", "<workbook/>")
    return source


def test_successful_dry_run_reaches_approved_without_mutation(tmp_path):
    source = _xlsx_stub(tmp_path)
    source_bytes_before = source.read_bytes()
    plan = _plan("run-test-1")
    calls: list = []

    result = run_dry_run(
        source,
        plan,
        tmp_path / "artifacts",
        execute_tool_fn=_fake_snapshot_execute_tool(calls),
    )

    assert result.context.state == TransactionState.APPROVED
    assert result.ok
    assert len(calls) == 1
    assert source.read_bytes() == source_bytes_before
    assert (tmp_path / "artifacts" / "backups").exists()
    assert not any((tmp_path / "artifacts").glob("**/*WORKING*"))


def test_dry_run_stops_before_approval_when_plan_has_unresolved_items(tmp_path):
    source = _xlsx_stub(tmp_path)
    plan = _plan("run-test-2", unresolved=("which sheet?",))

    result = run_dry_run(
        source,
        plan,
        tmp_path / "artifacts",
        execute_tool_fn=_fake_snapshot_execute_tool([]),
    )

    assert result.context.state == TransactionState.FAILED_SAFE
    assert result.plan_errors
    assert not result.awaiting_approval


def test_dry_run_waits_for_approval_when_plan_requires_it(tmp_path):
    source = _xlsx_stub(tmp_path)
    plan = _plan("run-test-3", requires_approval=True)

    result = run_dry_run(
        source,
        plan,
        tmp_path / "artifacts",
        execute_tool_fn=_fake_snapshot_execute_tool([]),
        approved=False,
    )

    assert result.awaiting_approval
    assert result.context.state == TransactionState.PLANNED


def test_dry_run_reaches_approved_when_explicitly_approved(tmp_path):
    source = _xlsx_stub(tmp_path)
    plan = _plan("run-test-4", requires_approval=True)

    result = run_dry_run(
        source,
        plan,
        tmp_path / "artifacts",
        execute_tool_fn=_fake_snapshot_execute_tool([]),
        approved=True,
    )

    assert result.context.state == TransactionState.APPROVED
    assert not result.awaiting_approval


def test_missing_source_file_fails_closed_before_backup(tmp_path):
    plan = _plan("run-test-5")

    result = run_dry_run(
        tmp_path / "does_not_exist.xlsx",
        plan,
        tmp_path / "artifacts",
        execute_tool_fn=_fake_snapshot_execute_tool([]),
    )

    assert result.context.state == TransactionState.FAILED_SAFE
    assert not (tmp_path / "artifacts" / "backups").exists()


def test_plan_referencing_a_locked_tool_fails_before_approval(tmp_path):
    source = _xlsx_stub(tmp_path)
    step = PlanStep(
        step_id="step-1",
        tool="set_formula",
        purpose="set a formula",
        request=ToolRequest(
            schema_version="1.0",
            transaction_id="run-test-6",
            tool="set_formula",
            arguments={},
            dry_run=True,
        ),
    )
    plan = OperationPlan(transaction_id="run-test-6", intent="test", steps=(step,))

    result = run_dry_run(
        source,
        plan,
        tmp_path / "artifacts",
        execute_tool_fn=_fake_snapshot_execute_tool([]),
    )

    assert result.context.state == TransactionState.FAILED_SAFE
    assert any("locked" in error for error in result.plan_errors)


def test_dry_run_holds_and_releases_the_workspace_lock(tmp_path):
    from src.transaction_state import TransactionContext

    source = _xlsx_stub(tmp_path)
    plan = _plan("run-test-7")

    result = run_dry_run(
        source,
        plan,
        tmp_path / "artifacts",
        execute_tool_fn=_fake_snapshot_execute_tool([]),
    )

    other = TransactionContext(source, tmp_path / "artifacts")
    other.acquire_workspace_lock()  # must not raise: the coordinator released its lock
    other.release_workspace_lock()
    assert result.context.state == TransactionState.APPROVED
