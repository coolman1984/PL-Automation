import json

import pytest

from src.transaction_state import TransactionContext, TransactionState


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

