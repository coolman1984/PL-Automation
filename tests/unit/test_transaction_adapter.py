from types import SimpleNamespace

import pytest

from src.core import transaction_adapter


def test_working_copy_refuses_unsaved_attached_source(tmp_path, monkeypatch):
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"source")
    session = SimpleNamespace(
        source_workbook=SimpleNamespace(Saved=False),
        close=lambda: None,
    )
    monkeypatch.setattr(transaction_adapter, "connect_source_readonly", lambda *_args: session)
    calls = []
    monkeypatch.setattr(transaction_adapter, "save_working_copy", lambda *_args: calls.append(True))

    with pytest.raises(ValueError, match="unsaved"):
        transaction_adapter.create_working_copy(source, tmp_path / "working.xlsx", "attach")

    assert calls == []
