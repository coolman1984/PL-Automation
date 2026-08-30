"""Gated acceptance for the first universal read-only range capability."""

from __future__ import annotations

import os
import sys

import pytest

if not sys.platform.startswith("win"):
    pytest.skip("Excel COM integration tests require Windows desktop Excel", allow_module_level=True)
pytest.importorskip("pythoncom")

from src.agent_contracts import TargetRef, ToolRequest
from src.tool_executor import execute_tool


def test_read_range_from_authorized_excel_workbook(source_workbook_path):
    sheet = os.environ.get("PL_COM_READ_SHEET", "").strip()
    address = os.environ.get("PL_COM_READ_ADDRESS", "").strip()
    if not sheet or not address:
        pytest.skip("Set PL_COM_READ_SHEET and PL_COM_READ_ADDRESS for read-range acceptance")
    request = ToolRequest.new(
        "read_range",
        target=TargetRef("source", sheet=sheet, address=address),
        arguments={
            "file": str(source_workbook_path),
            "mode": os.environ.get("PL_COM_READ_MODE", "auto"),
            "include_formulas": True,
        },
    )

    result = execute_tool(request)

    assert result.ok, result.error.to_dict() if result.error else result.to_dict()
    assert result.changed is False
    assert result.after_evidence["values"]

