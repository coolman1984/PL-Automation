"""Compact onboarding/status output for a newly arrived coding agent."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .tool_registry import describe_tool, tool_catalog


def _read(root: Path, relative: str) -> str:
    path = root / relative
    try:
        return path.read_text(encoding="utf-8")
    except OSError:
        return f"[missing: {relative}]"


def project_status(root: Path) -> dict[str, Any]:
    """Return a small machine-readable status record."""
    catalog = tool_catalog()
    status_text = _read(root, "docs/PROJECT_STATUS.md")
    next_task = ""
    marker = "## Next single task"
    if marker in status_text:
        next_task = status_text.split(marker, 1)[1].strip().split("\n\n", 1)[0]
    return {
        "schema_version": "1.0",
        "project": "PL Automation",
        "milestone": "M2 — read-only Excel engine in progress",
        "available_tools": catalog["available_count"],
        "known_tools": catalog["tool_count"],
        "status_file": "docs/PROJECT_STATUS.md",
        "next_task": next_task,
    }


def render_agent_start(root: Path, *, json_output: bool = False) -> str:
    """Render only the information needed for first contact with the repo."""
    root = Path(root).resolve()
    catalog = tool_catalog()
    status = project_status(root)
    payload = {
        "project": status,
        "read_first": [
            "AGENTS.md",
            "docs/START_HERE_AGENT.md",
            "docs/PROJECT_STATUS.md",
            "docs/FILE_MAP.md",
            "docs/UNIVERSAL_EXCEL_AGENT_CODING_PLAN_V1.md",
        ],
        "available_tools": [item["name"] for item in catalog["tools"] if item["status"] == "available"],
        "locked_tools": [item["name"] for item in catalog["tools"] if item["status"] != "available"],
        "commands": [
            "python app.py --agent-start",
            "python app.py --project-status",
            "python app.py --list-tools --format json",
            "python app.py --describe-tool TOOL_NAME",
        ],
    }
    if json_output:
        return json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    lines = [
        "EXCEL AGENT START",
        f"Project: {status['project']}",
        f"Milestone: {status['milestone']}",
        f"Tools: {status['available_tools']} available / {status['known_tools']} known",
        "Read first: " + " -> ".join(payload["read_first"]),
        "Available: " + ", ".join(payload["available_tools"]),
        "Locked: " + ", ".join(payload["locked_tools"]),
        "Next: " + status["next_task"].replace("\n", " "),
        "Safety: probe -> backup -> snapshot -> plan -> execute copy -> reopen -> validate -> publish",
    ]
    return "\n".join(lines) + "\n"


def render_project_status(root: Path, *, json_output: bool = False) -> str:
    status = project_status(Path(root).resolve())
    if json_output:
        return json.dumps(status, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    return (
        "EXCEL AGENT PROJECT STATUS\n"
        f"Milestone: {status['milestone']}\n"
        f"Tools: {status['available_tools']} available / {status['known_tools']} known\n"
        f"Next: {status['next_task'].replace(chr(10), ' ')}\n"
    )


def render_tool_description(name: str, *, json_output: bool = False) -> str:
    item = describe_tool(name)
    if item is None:
        payload = {"ok": False, "error": {"code": "unknown_tool", "message": f"Tool not found: {name}"}}
    else:
        payload = {"ok": True, "tool": item}
    if json_output:
        return json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    if not payload["ok"]:
        return f"UNKNOWN TOOL: {name}\n"
    tool = payload["tool"]
    return (
        f"TOOL: {tool['name']}\n"
        f"Category: {tool['category']}\n"
        f"Status: {tool['status']}\n"
        f"Description: {tool['description']}\n"
        f"Changes workbook: {'YES' if tool['mutates_workbook'] else 'NO'}\n"
        f"Needs backup: {'YES' if tool['requires_backup'] else 'NO'}\n"
    )
