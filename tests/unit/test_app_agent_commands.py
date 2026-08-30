import json
from pathlib import Path

import app


def test_agent_commands_do_not_require_a_workbook():
    assert app.main(["--agent-start"]) == 0
    assert app.main(["--project-status", "--format", "json"]) == 0
    assert app.main(["--describe-tool", "create_backup", "--format", "json"]) == 0


def test_run_tool_accepts_a_json_request_file(tmp_path: Path):
    source = tmp_path / "not_excel.txt"
    source.write_bytes(b"not excel")
    request_path = tmp_path / "request.json"
    request_path.write_text(
        json.dumps({
            "tool": "inspect_file",
            "transaction_id": "run-test",
            "arguments": {"file": str(source)},
            "dry_run": True,
        }),
        encoding="utf-8",
    )

    assert app.main(["--run-tool", str(request_path), "--format", "json"]) == 0

