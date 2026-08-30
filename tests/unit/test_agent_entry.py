from pathlib import Path

from src.agent_entry import render_agent_start, render_project_status, render_tool_description


def test_agent_start_exposes_read_order_and_locked_tools():
    root = Path(__file__).parents[2]
    output = render_agent_start(root)

    assert "EXCEL AGENT START" in output
    assert "AGENTS.md" in output
    assert "write_range" in output
    assert "probe -> backup" in output


def test_status_and_tool_description_have_machine_mode():
    root = Path(__file__).parents[2]

    assert '"milestone"' in render_project_status(root, json_output=True)
    assert '"name": "create_backup"' in render_tool_description("create_backup", json_output=True)

