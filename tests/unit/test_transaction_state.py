import json
import os
from pathlib import Path

import pytest

from src.transaction_state import (
    TransactionContext,
    TransactionLockError,
    TransactionState,
)


def test_transaction_context_accepts_only_safe_ordered_transitions(tmp_path):
    context = TransactionContext(tmp_path / "book.xlsx", tmp_path / "artifacts")

    context.transition(TransactionState.PROBED)
    context.transition(TransactionState.BACKED_UP)
    context.transition(TransactionState.SNAPSHOTTED)
    context.transition(TransactionState.PLANNED)
    context.transition(TransactionState.APPROVED)

    assert context.state == TransactionState.APPROVED
    with pytest.raises(ValueError, match="Invalid transaction transition"):
        context.transition(TransactionState.PUBLISHED)


def test_transaction_failure_is_terminal_and_journal_is_reproducible(tmp_path):
    context = TransactionContext(tmp_path / "book.xlsx", tmp_path / "artifacts")
    context.fail_safe("simulated error", evidence={"phase": "backup"})

    path = context.write_journal()
    data = json.loads(path.read_text(encoding="utf-8"))

    assert context.state == TransactionState.FAILED_SAFE
    assert data["events"][-1]["event"] == "failed_safe"
    with pytest.raises(ValueError, match="Invalid transaction transition"):
        context.transition(TransactionState.PROBED)


def test_workspace_lock_blocks_a_second_transaction_on_the_same_source(tmp_path):
    source = tmp_path / "book.xlsb"
    source.write_bytes(b"stub")
    first = TransactionContext(source, tmp_path / "artifacts")
    second = TransactionContext(source, tmp_path / "artifacts")

    first.acquire_workspace_lock()
    with pytest.raises(TransactionLockError):
        second.acquire_workspace_lock()

    first.release_workspace_lock()
    second.acquire_workspace_lock()
    second.release_workspace_lock()


def test_workspace_lock_is_reentrant_for_the_same_transaction(tmp_path):
    source = tmp_path / "book.xlsb"
    source.write_bytes(b"stub")
    context = TransactionContext(source, tmp_path / "artifacts")

    context.acquire_workspace_lock()
    context.acquire_workspace_lock()
    context.release_workspace_lock()


def test_stale_workspace_lock_from_a_dead_pid_is_reclaimed(tmp_path):
    source = tmp_path / "book.xlsb"
    source.write_bytes(b"stub")
    context = TransactionContext(source, tmp_path / "artifacts")
    stale_lock_path = context._source_lock_path()
    stale_lock_path.parent.mkdir(parents=True, exist_ok=True)
    stale_lock_path.write_text(
        json.dumps({"transaction_id": "run-dead", "pid": 999_999_999}),
        encoding="utf-8",
    )

    context.acquire_workspace_lock()
    context.release_workspace_lock()


def test_hash_stages_are_recorded_and_persisted(tmp_path):
    source = tmp_path / "book.xlsb"
    source.write_bytes(b"source-bytes")
    backup = tmp_path / "backup.xlsb"
    backup.write_bytes(b"backup-bytes")
    context = TransactionContext(source, tmp_path / "artifacts")

    digest = context.record_hash("backup", backup)

    assert context.backup_sha256 == digest
    persisted = json.loads(context.journal_path.read_text(encoding="utf-8"))
    assert persisted["backup_sha256"] == digest

    with pytest.raises(ValueError):
        context.record_hash("bogus", backup)


def test_journal_is_written_automatically_without_an_explicit_call(tmp_path):
    context = TransactionContext(tmp_path / "book.xlsb", tmp_path / "artifacts")
    context.transition(TransactionState.PROBED)

    persisted = json.loads(context.journal_path.read_text(encoding="utf-8"))
    assert persisted["state"] == TransactionState.PROBED.value
    assert persisted["events"][-1]["event"] == "state_changed"


def test_resume_preserves_sequence_numbers_without_repeats(tmp_path):
    context = TransactionContext(tmp_path / "book.xlsb", tmp_path / "artifacts")
    context.transition(TransactionState.PROBED)
    context.transition(TransactionState.BACKED_UP)
    journal_path = context.journal_path
    sequences_before = [event["sequence"] for event in context.events]

    resumed = TransactionContext.resume(journal_path)
    resumed.transition(TransactionState.SNAPSHOTTED)

    sequences_after = [event["sequence"] for event in resumed.events]
    assert sequences_after[: len(sequences_before)] == sequences_before
    assert len(sequences_after) == len(set(sequences_after))
    assert resumed.state == TransactionState.SNAPSHOTTED


def test_resuming_an_unsafe_in_flight_state_fails_closed(tmp_path):
    context = TransactionContext(tmp_path / "book.xlsb", tmp_path / "artifacts")
    context.transition(TransactionState.PROBED)
    context.transition(TransactionState.BACKED_UP)
    context.transition(TransactionState.SNAPSHOTTED)
    context.transition(TransactionState.PLANNED)
    context.transition(TransactionState.APPROVED)
    context.transition(TransactionState.WORKING_COPY_READY)
    context.transition(TransactionState.EXECUTING)
    journal_path = context.journal_path

    resumed = TransactionContext.resume(journal_path)

    assert resumed.state == TransactionState.FAILED_SAFE
    assert resumed.events[-1]["event"] == "failed_safe"


def test_write_journal_retries_past_a_transient_windows_permission_error(tmp_path, monkeypatch):
    context = TransactionContext(tmp_path / "book.xlsx", tmp_path / "artifacts")
    real_replace = Path.replace
    calls = {"n": 0}

    def flaky_replace(self, target):
        calls["n"] += 1
        if calls["n"] <= 2:
            raise PermissionError("simulated transient Windows file lock")
        return real_replace(self, target)

    monkeypatch.setattr(Path, "replace", flaky_replace)
    monkeypatch.setattr("time.sleep", lambda _seconds: None)

    context.transition(TransactionState.PROBED)

    assert calls["n"] == 3
    persisted = json.loads(context.journal_path.read_text(encoding="utf-8"))
    assert persisted["state"] == TransactionState.PROBED.value


def test_interrupted_journal_write_does_not_corrupt_the_last_valid_state(tmp_path):
    context = TransactionContext(tmp_path / "book.xlsb", tmp_path / "artifacts")
    context.transition(TransactionState.PROBED)
    journal_path = context.journal_path
    good_contents = journal_path.read_text(encoding="utf-8")

    # Simulate a crash mid-write: a stray, truncated temp file is left behind.
    temp_path = journal_path.with_suffix(journal_path.suffix + ".tmp")
    temp_path.write_text('{"state": "not json-complete', encoding="utf-8")

    resumed = TransactionContext.resume(journal_path)

    assert resumed.state == TransactionState.PROBED
    assert journal_path.read_text(encoding="utf-8").startswith(good_contents[:20])
    os.remove(temp_path)

