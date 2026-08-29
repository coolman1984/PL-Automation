from __future__ import annotations

import importlib.util

from src.runtime_check import _dependency_check, render_runtime_check


def test_dependency_check_handles_missing_parent_module(monkeypatch):
    def missing(_name: str):
        raise ModuleNotFoundError("missing parent")

    monkeypatch.setattr(importlib.util, "find_spec", missing)
    result = _dependency_check("missing.child", "Missing dependency")
    assert result["passed"] is False
    assert result["required"] is True


def test_render_runtime_check_includes_machine_readable_json():
    report = {
        "ready": True,
        "frozen_application": True,
        "executable": "agent.exe",
        "python_version": "3.12.0",
        "checks": [
            {"name": "example", "passed": True, "required": True, "detail": "ok"}
        ],
    }
    rendered = render_runtime_check(report)
    assert "Overall: READY" in rendered
    assert "[PASS] example: ok" in rendered
    assert '"ready": true' in rendered
