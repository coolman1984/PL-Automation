"""Task 6: the mutating half of the coordinator fails closed at every stage.

A fake, in-process Excel COM double drives the real production fingerprinting
code (``workbook_audit.collect_fingerprint``) so these tests exercise the
actual preservation-comparison logic, not a re-implemented stand-in.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from src.agent_contracts import OperationPlan, PlanStep, TargetRef, ToolRequest, ToolResult
from src.core.coordinator import DryRunResult, ExecuteResult, run_execute
from src.core.transaction_adapter import ReopenedHandle, WorkingCopyHandle
from src.fake_engine import FakeEngine
from src.transaction_state import TransactionContext, TransactionState


class _Count:
    """Mimics a COM collection that only ever exposes .Count."""

    def __init__(self, count: int = 0):
        self.Count = count


class _FakeSheet:
    def __init__(self, name: str):
        self.Name = name

    def PivotTables(self):
        return _Count(0)


class _FakeWorkbook:
    """Exposes exactly the attributes ``workbook_audit`` touches."""

    def __init__(self, path: Path, sheet_names: list[str], file_format: int = 50, mutated: bool = False):
        self._path = path
        self._sheet_names = list(sheet_names)
        self.FileFormat = file_format
        self.Connections = _Count(0)
        self.Names = _Count(0)
        self._mutated = mutated

    @property
    def Sheets(self):
        return _SheetsCollection(self._sheet_names)

    def Worksheets(self, name):
        return _FakeSheet(name)

    def LinkSources(self, _kind):
        return None

    @property
    def VBProject(self):
        raise AttributeError("no VBA project")

    def Save(self):
        Path(self._path).write_bytes(Path(self._path).read_bytes())

    def Close(self, SaveChanges=False):
        pass


class _SheetsCollection:
    def __init__(self, names: list[str]):
        self._names = names
        self.Count = len(names)

    def __call__(self, index):
        return _FakeSheet(self._names[index - 1])


def _xlsx_stub(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("xl/workbook.xml", "<workbook/>")


def _approved_dry_run(tmp_path: Path, transaction_id: str) -> DryRunResult:
    source = tmp_path / "book.xlsx"
    _xlsx_stub(source)
    context = TransactionContext(source, tmp_path / "artifacts", transaction_id=transaction_id)
    context.record_hash("source", source)
    context.transition(TransactionState.PROBED)
    context.transition(TransactionState.BACKED_UP)
    context.transition(TransactionState.SNAPSHOTTED)
    context.transition(TransactionState.PLANNED)
    context.transition(TransactionState.APPROVED)
    return DryRunResult(context)


def _harmless_plan(transaction_id: str) -> OperationPlan:
    step = PlanStep(
        step_id="step-1",
        tool="read_range",
        purpose="read a cell",
        request=ToolRequest(
            schema_version="1.0",
            transaction_id=transaction_id,
            tool="read_range",
            target=TargetRef("working-copy", sheet="Sheet1", address="A1"),
            dry_run=False,
        ),
    )
    return OperationPlan(transaction_id=transaction_id, intent="test", steps=(step,))


def _fake_ops(sheet_names=("Sheet1",), fail_step: bool = False):
    """Build fake create/open/reopen callables sharing one workbook path."""
    state = {}

    def create_working_copy(source_path: Path, working_path: Path, mode: str) -> None:
        working_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, working_path)
        state["path"] = working_path

    def open_working_copy(working_path: Path) -> WorkingCopyHandle:
        workbook = _FakeWorkbook(working_path, list(sheet_names))
        engine = FakeEngine({name: {"1,1": "hello"} for name in sheet_names})
        state["workbook"] = workbook
        return WorkingCopyHandle(engine=engine, _session=_NoopSession(), _workbook=workbook, _path=working_path)

    def reopen_working_copy(working_path: Path) -> ReopenedHandle:
        workbook = state["workbook"]
        return ReopenedHandle(_session=_NoopSession(), _workbook=workbook, _path=working_path)

    return create_working_copy, open_working_copy, reopen_working_copy


class _NoopSession:
    def close_workbook(self, workbook, save_changes=False):
        pass

    def close(self):
        pass


def _fake_execute_tool(fail: bool = False):
    def fake(request, **_kwargs):
        if fail:
            from src.agent_contracts import ToolError

            return ToolResult.failure(request.tool, ToolError(code="simulated", message="boom"))
        return ToolResult.success(request.tool, changed=False, after_evidence={"values": [["hello"]]})

    return fake


def test_successful_execute_reaches_published_and_preserves_source(tmp_path):
    transaction_id = "run-exec-1"
    dry_run = _approved_dry_run(tmp_path, transaction_id)
    plan = _harmless_plan(transaction_id)
    source_bytes_before = dry_run.context.source_path.read_bytes()
    create_fn, open_fn, reopen_fn = _fake_ops()

    result = run_execute(
        dry_run,
        plan,
        execute_tool_fn=_fake_execute_tool(),
        create_working_copy_fn=create_fn,
        open_working_copy_fn=open_fn,
        reopen_working_copy_fn=reopen_fn,
    )

    assert result.context.state == TransactionState.PUBLISHED
    assert result.ok
    assert result.output_path is not None
    assert result.output_path.exists()
    assert dry_run.context.source_path.read_bytes() == source_bytes_before


def test_failed_step_fails_closed_before_saving(tmp_path):
    transaction_id = "run-exec-2"
    dry_run = _approved_dry_run(tmp_path, transaction_id)
    plan = _harmless_plan(transaction_id)
    create_fn, open_fn, reopen_fn = _fake_ops()

    result = run_execute(
        dry_run,
        plan,
        execute_tool_fn=_fake_execute_tool(fail=True),
        create_working_copy_fn=create_fn,
        open_working_copy_fn=open_fn,
        reopen_working_copy_fn=reopen_fn,
    )

    assert result.context.state == TransactionState.FAILED_SAFE
    assert result.output_path is None


def test_execute_refuses_to_run_when_not_approved(tmp_path):
    transaction_id = "run-exec-3"
    source = tmp_path / "book.xlsx"
    _xlsx_stub(source)
    context = TransactionContext(source, tmp_path / "artifacts", transaction_id=transaction_id)
    dry_run = DryRunResult(context)  # still RECEIVED, never approved
    plan = _harmless_plan(transaction_id)
    create_fn, open_fn, reopen_fn = _fake_ops()

    result = run_execute(
        dry_run,
        plan,
        execute_tool_fn=_fake_execute_tool(),
        create_working_copy_fn=create_fn,
        open_working_copy_fn=open_fn,
        reopen_working_copy_fn=reopen_fn,
    )

    assert result.context.state == TransactionState.FAILED_SAFE
    assert result.output_path is None


def test_source_change_since_backup_fails_closed_before_any_copy(tmp_path):
    transaction_id = "run-exec-4"
    dry_run = _approved_dry_run(tmp_path, transaction_id)
    dry_run.context.source_path.write_bytes(b"tampered")
    plan = _harmless_plan(transaction_id)
    create_fn, open_fn, reopen_fn = _fake_ops()
    calls = []

    def spy_create(*args, **kwargs):
        calls.append(args)
        return create_fn(*args, **kwargs)

    result = run_execute(
        dry_run,
        plan,
        execute_tool_fn=_fake_execute_tool(),
        create_working_copy_fn=spy_create,
        open_working_copy_fn=open_fn,
        reopen_working_copy_fn=reopen_fn,
    )

    assert result.context.state == TransactionState.FAILED_SAFE
    assert not calls


def test_preservation_violation_after_edits_fails_closed_before_save(tmp_path):
    transaction_id = "run-exec-5"
    dry_run = _approved_dry_run(tmp_path, transaction_id)
    plan = _harmless_plan(transaction_id)
    create_fn, _, reopen_fn = _fake_ops()

    # fingerprint() is called twice on the same handle (before/after edits);
    # simulate a sheet appearing mid-execution by mutating the handle's
    # workbook in place from inside the fake tool executor.
    handle_holder: dict = {}

    def open_fn(working_path: Path) -> WorkingCopyHandle:
        workbook = _FakeWorkbook(working_path, ["Sheet1"])
        engine = FakeEngine({"Sheet1": {"1,1": "hello"}})
        handle = WorkingCopyHandle(engine=engine, _session=_NoopSession(), _workbook=workbook, _path=working_path)
        handle_holder["handle"] = handle
        return handle

    def fake_execute_tool_that_mutates(request, **_kwargs):
        # Simulate an engine operation that silently added a sheet: the next
        # fingerprint() call must see the extra sheet.
        handle_holder["handle"]._workbook._sheet_names.append("UnexpectedNewSheet")
        return ToolResult.success(request.tool, changed=False, after_evidence={})

    result = run_execute(
        dry_run,
        plan,
        execute_tool_fn=fake_execute_tool_that_mutates,
        create_working_copy_fn=create_fn,
        open_working_copy_fn=open_fn,
        reopen_working_copy_fn=reopen_fn,
    )

    assert result.context.state == TransactionState.FAILED_SAFE
    assert result.output_path is None
