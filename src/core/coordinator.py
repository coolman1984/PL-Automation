"""Generic coordinator: the non-mutating half of a universal transaction.

This composes only already-declared tools (via ``execute_tool``) and the
generic transaction journal.  It never opens a working copy, never calls a
mutating engine operation, and never chooses a target by similarity.  The
mutating half (working copy, save, reopen, validate, publish) is Task 6.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from ..agent_contracts import OperationPlan, ToolRequest, ToolResult
from ..file_transaction import publish_validated_file, sha256_file
from ..plan_validation import validate_plan
from ..tool_executor import execute_tool as _real_execute_tool
from ..transaction_state import TransactionContext, TransactionState
from .transaction_adapter import compare_generic_preservation
from .transaction_adapter import create_working_copy as _real_create_working_copy
from .transaction_adapter import open_working_copy_for_edit as _real_open_working_copy
from .transaction_adapter import reopen_working_copy as _real_reopen_working_copy

ExecuteToolFn = Callable[..., ToolResult]
CreateWorkingCopyFn = Callable[[Path, Path, str], None]
OpenWorkingCopyFn = Callable[[Path], Any]
ReopenWorkingCopyFn = Callable[[Path], Any]


@dataclass
class DryRunResult:
    """Outcome of the non-mutating coordinator phase."""

    context: TransactionContext
    plan_errors: tuple[str, ...] = ()
    awaiting_approval: bool = False

    @property
    def ok(self) -> bool:
        return (
            self.context.state == TransactionState.APPROVED
            and not self.plan_errors
            and not self.awaiting_approval
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "transaction_id": self.context.transaction_id,
            "state": self.context.state.value,
            "plan_errors": list(self.plan_errors),
            "awaiting_approval": self.awaiting_approval,
            "journal_path": str(self.context.journal_path),
        }


def run_dry_run(
    source_path: Path,
    plan: OperationPlan,
    artifact_root: Path,
    *,
    approved: bool = False,
    project_root: Path | None = None,
    snapshot_mode: str = "auto",
    max_snapshot_cells: int = 250_000,
    excel_mode: str = "auto",
    execute_tool_fn: ExecuteToolFn = _real_execute_tool,
) -> DryRunResult:
    """Probe, back up, snapshot, and validate a plan without any mutation.

    The transaction never advances past ``APPROVED``. No working copy is
    created and no engine write operation is invoked; only read-only
    evidence-gathering tools (``inspect_file``, ``create_backup``,
    ``snapshot_workbook``) run, each through the same declared-tool executor
    an agent would use, so this coordinator cannot invent a capability that
    isn't in the catalogue.
    """
    source_path = Path(source_path).expanduser().resolve()
    artifact_root = Path(artifact_root).expanduser().resolve()
    root = Path(project_root).expanduser().resolve() if project_root else artifact_root

    context = TransactionContext(source_path, artifact_root, transaction_id=plan.transaction_id)
    context.acquire_workspace_lock()
    try:
        probe_result = execute_tool_fn(
            ToolRequest.new("inspect_file", arguments={"file": str(source_path)})
        )
        if not probe_result.ok or not probe_result.after_evidence.get("recognized"):
            context.fail_safe(
                "Source workbook probe failed or was not recognized",
                evidence=probe_result.to_dict(),
            )
            return DryRunResult(context)
        context.record_hash("source", source_path)
        context.transition(TransactionState.PROBED, evidence=probe_result.after_evidence)

        backup_result = execute_tool_fn(
            ToolRequest.new(
                "create_backup",
                arguments={
                    "file": str(source_path),
                    "backup_root": str(artifact_root / "backups"),
                    "reason": "coordinator_dry_run",
                },
                dry_run=False,
            ),
            project_root=root,
        )
        if not backup_result.ok:
            context.fail_safe("Backup failed", evidence=backup_result.to_dict())
            return DryRunResult(context)
        context.record_hash("backup", Path(backup_result.after_evidence["backup_file"]))
        context.transition(TransactionState.BACKED_UP, evidence=backup_result.after_evidence)

        snapshot_result = execute_tool_fn(
            ToolRequest.new(
                "snapshot_workbook",
                arguments={
                    "file": str(source_path),
                    "artifact_root": str(artifact_root),
                    "mode": excel_mode,
                    "snapshot_mode": snapshot_mode,
                    "max_snapshot_cells": max_snapshot_cells,
                },
                dry_run=False,
            ),
            project_root=root,
        )
        if not snapshot_result.ok:
            context.fail_safe("Snapshot failed", evidence=snapshot_result.to_dict())
            return DryRunResult(context)
        context.transition(TransactionState.SNAPSHOTTED, evidence=snapshot_result.after_evidence)

        plan_errors = list(validate_plan(plan))
        if plan.transaction_id != context.transaction_id:
            plan_errors.append("plan transaction_id does not match this transaction")
        if plan_errors:
            context.fail_safe("Plan validation failed", evidence={"errors": plan_errors})
            return DryRunResult(context, plan_errors=tuple(plan_errors))
        context.transition(
            TransactionState.PLANNED,
            evidence={"intent": plan.intent, "risk": plan.risk, "step_count": len(plan.steps)},
        )

        if plan.requires_approval and not approved:
            return DryRunResult(context, awaiting_approval=True)
        context.transition(TransactionState.APPROVED, evidence={"approved_explicitly": bool(approved)})
        return DryRunResult(context)
    finally:
        context.release_workspace_lock()


@dataclass
class ExecuteResult:
    """Outcome of the mutating coordinator phase."""

    context: TransactionContext
    output_path: Path | None = None
    step_results: tuple[ToolResult, ...] = ()

    @property
    def ok(self) -> bool:
        return self.context.state == TransactionState.PUBLISHED and self.output_path is not None

    def to_dict(self) -> dict[str, Any]:
        return {
            "transaction_id": self.context.transaction_id,
            "state": self.context.state.value,
            "output_path": str(self.output_path) if self.output_path else None,
            "journal_path": str(self.context.journal_path),
        }


def run_execute(
    dry_run: DryRunResult,
    plan: OperationPlan,
    *,
    output_dir: Path | None = None,
    excel_mode: str = "open",
    execute_tool_fn: ExecuteToolFn = _real_execute_tool,
    create_working_copy_fn: CreateWorkingCopyFn = _real_create_working_copy,
    open_working_copy_fn: OpenWorkingCopyFn = _real_open_working_copy,
    reopen_working_copy_fn: ReopenWorkingCopyFn = _real_reopen_working_copy,
) -> ExecuteResult:
    """Run the mutating half of the transaction from an ``APPROVED`` context.

    Every failure — a bad step, a preservation violation, a source that
    changed underneath us, a publish error — transitions the context to
    ``FAILED_SAFE`` and stops. The source is only ever opened read-only.
    Publication is reachable only by passing through ``VALIDATED``.
    """
    context = dry_run.context
    if context.state != TransactionState.APPROVED:
        context.fail_safe(
            "Execute requires an approved transaction",
            evidence={"actual_state": context.state.value},
        )
        return ExecuteResult(context)

    context.acquire_workspace_lock()
    step_results: list[ToolResult] = []
    working_handle = None
    reopened_handle = None
    try:
        current_source_hash = sha256_file(context.source_path)
        if current_source_hash != context.source_sha256:
            context.fail_safe(
                "Source workbook changed since it was probed and backed up",
                evidence={"expected": context.source_sha256, "actual": current_source_hash},
            )
            return ExecuteResult(context)

        working_dir = context.artifact_root / context.transaction_id / "working"
        working_dir.mkdir(parents=True, exist_ok=True)
        working_path = working_dir / f"{context.source_path.stem}__WORKING{context.source_path.suffix}"
        try:
            create_working_copy_fn(context.source_path, working_path, excel_mode)
        except Exception as exc:
            context.fail_safe(f"Working copy creation failed: {exc}")
            return ExecuteResult(context)
        context.record_hash("working", working_path)
        context.transition(TransactionState.WORKING_COPY_READY, evidence={"working_path": str(working_path)})

        try:
            working_handle = open_working_copy_fn(working_path)
        except Exception as exc:
            context.fail_safe(f"Could not open the working copy for edit: {exc}")
            return ExecuteResult(context)
        before_edit_fp = working_handle.fingerprint()

        context.transition(TransactionState.EXECUTING, evidence={"step_count": len(plan.steps)})
        for step in plan.steps:
            result = execute_tool_fn(step.request, engine=working_handle.engine)
            step_results.append(result)
            if not result.ok:
                context.fail_safe(
                    f"Plan step failed: {step.step_id}",
                    evidence={"step_id": step.step_id, "tool": step.tool, "error": result.to_dict()},
                )
                working_handle.discard()
                working_handle = None
                return ExecuteResult(context, step_results=tuple(step_results))

        after_edit_fp = working_handle.fingerprint()
        pre_save_checks = compare_generic_preservation(before_edit_fp, after_edit_fp)
        failed = [check for check in pre_save_checks if check.required and not check.passed]
        if failed:
            context.fail_safe(
                "Pre-save preservation check failed; refusing to save",
                evidence={"failed_checks": [check.name for check in failed]},
            )
            working_handle.discard()
            working_handle = None
            return ExecuteResult(context, step_results=tuple(step_results))

        working_handle.save_and_close()
        working_handle = None
        context.transition(TransactionState.SAVED, evidence={"step_count": len(step_results)})

        try:
            reopened_handle = reopen_working_copy_fn(working_path)
        except Exception as exc:
            context.fail_safe(f"Reopen failed: {exc}")
            return ExecuteResult(context, step_results=tuple(step_results))
        context.transition(TransactionState.REOPENED, evidence={"reopened": True})

        post_reopen_fp = reopened_handle.fingerprint()
        post_checks = compare_generic_preservation(after_edit_fp, post_reopen_fp)
        reopened_handle.close()
        reopened_handle = None
        failed = [check for check in post_checks if check.required and not check.passed]
        if failed:
            context.fail_safe(
                "Post-reopen validation failed; refusing to publish",
                evidence={"failed_checks": [check.name for check in failed]},
            )
            return ExecuteResult(context, step_results=tuple(step_results))

        context.record_hash("working", working_path)
        context.transition(TransactionState.VALIDATED, evidence={"post_reopen_checks": len(post_checks)})

        # Reconfirm the source one final time immediately before publication.
        current_source_hash = sha256_file(context.source_path)
        if current_source_hash != context.source_sha256:
            context.fail_safe(
                "Source workbook changed during execution; refusing to publish",
                evidence={"expected": context.source_sha256, "actual": current_source_hash},
            )
            return ExecuteResult(context, step_results=tuple(step_results))

        output_root = Path(output_dir) if output_dir else context.artifact_root / context.transaction_id / "output"
        output_root.mkdir(parents=True, exist_ok=True)
        output_path = output_root / f"{context.source_path.stem}__UPDATED{context.source_path.suffix}"
        try:
            published = Path(publish_validated_file(working_path, output_path))
        except Exception as exc:
            context.fail_safe(f"Publication failed: {exc}")
            return ExecuteResult(context, step_results=tuple(step_results))
        context.record_hash("output", published)
        context.transition(TransactionState.PUBLISHED, evidence={"output_path": str(published)})
        return ExecuteResult(context, output_path=published, step_results=tuple(step_results))
    finally:
        if working_handle is not None:
            try:
                working_handle.discard()
            except Exception:
                pass
        if reopened_handle is not None:
            try:
                reopened_handle.close()
            except Exception:
                pass
        context.release_workspace_lock()
