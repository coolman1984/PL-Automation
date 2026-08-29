"""Dependency-light runtime and Microsoft Excel readiness check."""

from __future__ import annotations

import importlib.util
import json
import platform
import sys
import tempfile
from pathlib import Path
from typing import Any


def _dependency_check(module: str, label: str) -> dict[str, Any]:
    try:
        found = importlib.util.find_spec(module) is not None
    except (ImportError, ModuleNotFoundError, ValueError):
        found = False
    return {
        "name": label,
        "passed": found,
        "required": True,
        "detail": "available" if found else f"missing module: {module}",
    }


def _excel_registration_check() -> dict[str, Any]:
    if platform.system() != "Windows":
        return {
            "name": "Microsoft Excel desktop",
            "passed": False,
            "required": True,
            "detail": "Windows is required for Excel COM automation",
        }
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID") as key:
            clsid, _ = winreg.QueryValueEx(key, None)
        return {
            "name": "Microsoft Excel desktop",
            "passed": bool(clsid),
            "required": True,
            "detail": f"registered COM class: {clsid}",
        }
    except OSError as exc:
        return {
            "name": "Microsoft Excel desktop",
            "passed": False,
            "required": True,
            "detail": f"Excel COM registration was not found: {exc}",
        }


def _write_check(project_root: Path) -> dict[str, Any]:
    try:
        work = project_root / "work"
        work.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=work, prefix="runtime-check-", delete=True):
            pass
        return {
            "name": "Application work folder",
            "passed": True,
            "required": True,
            "detail": str(work),
        }
    except OSError as exc:
        return {
            "name": "Application work folder",
            "passed": False,
            "required": True,
            "detail": f"not writable: {exc}",
        }


def run_runtime_check(project_root: Path) -> dict[str, Any]:
    checks = [
        {
            "name": "Windows operating system",
            "passed": platform.system() == "Windows",
            "required": True,
            "detail": platform.platform(),
        },
        {
            "name": "64-bit process",
            "passed": sys.maxsize > 2**32,
            "required": True,
            "detail": platform.architecture()[0],
        },
        _dependency_check("pythoncom", "Excel COM bridge"),
        _dependency_check("win32com.client", "Excel automation package"),
        _dependency_check("yaml", "Configuration reader"),
        _excel_registration_check(),
        _write_check(project_root),
    ]
    return {
        "ready": all(item["passed"] for item in checks if item["required"]),
        "frozen_application": bool(getattr(sys, "frozen", False)),
        "executable": sys.executable,
        "python_version": platform.python_version(),
        "checks": checks,
    }


def render_runtime_check(report: dict[str, Any]) -> str:
    lines = [
        "EXCEL AGENT SELF-CHECK",
        f"Overall: {'READY' if report['ready'] else 'NOT READY'}",
        f"Packaged application: {'YES' if report['frozen_application'] else 'NO'}",
        f"Runtime: {report['executable']}",
        f"Runtime version: {report['python_version']}",
    ]
    for item in report["checks"]:
        lines.append(
            f"[{'PASS' if item['passed'] else 'FAIL'}] {item['name']}: {item['detail']}"
        )
    lines.append("JSON:")
    lines.append(json.dumps(report, ensure_ascii=False, indent=2))
    return "\n".join(lines) + "\n"
