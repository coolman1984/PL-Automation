from src.agent_contracts import TargetRef, ToolRequest
from src.fake_engine import FakeEngine
from src.tool_executor import execute_tool


def _request(tool, *, target=None, arguments=None, dry_run=False):
    return ToolRequest.new(tool, target=target, arguments=arguments or {}, dry_run=dry_run)


def _cell(address="A1"):
    return TargetRef("working-copy", sheet="Data", address=address)


def test_format_range_is_bounded_and_supports_dry_run():
    engine = FakeEngine({"Data": {}})
    request = _request("format_range", target=_cell("A1:B2"), arguments={"format": {"bold": True}}, dry_run=True)
    result = execute_tool(request, engine=engine)
    assert result.ok and not result.changed
    assert engine.features["formats"] == {}

    result = execute_tool(_request("format_range", target=_cell("A1:B2"), arguments={"format": {"bold": True}}), engine=engine)
    assert result.ok and result.changed
    assert engine.features["formats"]["Data!A1:B2"] == {"bold": True}


def test_advanced_mutations_reject_the_source_workbook():
    engine = FakeEngine({"Data": {}})
    result = execute_tool(
        _request("manage_comment", target=TargetRef("source", sheet="Data", address="A1"), arguments={"operation": "set", "text": "x"}),
        engine=engine,
    )
    assert not result.ok
    assert result.error.code == "invalid_target_workbook"


def test_insert_rows_requires_exact_anchor_fingerprint():
    engine = FakeEngine({"Data": {}})
    bad = execute_tool(_request("insert_rows", target=_cell("A5"), arguments={"operation": "insert", "count": 2, "expected_anchor_row": 4}), engine=engine)
    assert not bad.ok and bad.error.code == "anchor_row_mismatch"
    good = execute_tool(_request("insert_rows", target=_cell("A5"), arguments={"operation": "insert", "count": 2, "expected_anchor_row": 5}), engine=engine)
    assert good.ok


def test_sheet_lifecycle_and_nonempty_delete_guard():
    engine = FakeEngine({"Data": {}, "Busy": {"1,1": "x"}})
    created = execute_tool(_request("manage_sheet", target=TargetRef("working-copy", sheet="New"), arguments={"operation": "create", "name": "New"}), engine=engine)
    assert created.ok and "New" in engine.values
    renamed = execute_tool(_request("manage_sheet", target=TargetRef("working-copy", sheet="New"), arguments={"operation": "rename", "new_name": "Renamed"}), engine=engine)
    assert renamed.ok and "Renamed" in engine.values
    blocked = execute_tool(_request("manage_sheet", target=TargetRef("working-copy", sheet="Busy"), arguments={"operation": "delete_empty", "expected_empty": True}), engine=engine)
    assert not blocked.ok


def test_table_filter_and_validation_operations():
    engine = FakeEngine({"Data": {}})
    table = TargetRef("working-copy", sheet="Data", address="A1:C4", object_name="Sales")
    assert execute_tool(_request("manage_table", target=table, arguments={"operation": "create", "has_headers": True}), engine=engine).ok
    assert engine.features["tables"]["Data!Sales"]["address"] == "A1:C4"
    bad_filter = execute_tool(_request("manage_filter", target=_cell("A1:C4"), arguments={"operation": "apply", "field": 4, "criteria1": "x"}), engine=engine)
    assert not bad_filter.ok and bad_filter.error.code == "filter_field_out_of_range"
    assert execute_tool(_request("manage_filter", target=_cell("A1:C4"), arguments={"operation": "apply", "field": 2, "criteria1": "x"}), engine=engine).ok
    assert execute_tool(_request("manage_validation", target=_cell("C2:C4"), arguments={"operation": "set", "validation_type": "list", "formula1": '=\"Y,N\"'}), engine=engine).ok


def test_comment_and_name_use_current_value_fingerprints():
    engine = FakeEngine({"Data": {}})
    assert execute_tool(_request("manage_comment", target=_cell(), arguments={"operation": "set", "text": "first", "expected_current_text": None}), engine=engine).ok
    mismatch = execute_tool(_request("manage_comment", target=_cell(), arguments={"operation": "set", "text": "second", "expected_current_text": "other"}), engine=engine)
    assert not mismatch.ok and mismatch.error.code == "current_fingerprint_mismatch"
    assert execute_tool(_request("manage_name", arguments={"operation": "set", "name": "TaxRate", "refers_to": "=0.15", "expected_current_refers_to": None}), engine=engine).ok
    assert engine.features["names"]["TaxRate"]["refers_to"] == "=0.15"


def test_hyperlink_rejects_unsafe_scheme_and_allows_https():
    engine = FakeEngine({"Data": {}})
    unsafe = execute_tool(_request("manage_hyperlink", target=_cell(), arguments={"operation": "set", "address": "file:///secret", "expected_current_address": None}), engine=engine)
    assert not unsafe.ok and unsafe.error.code == "unsafe_hyperlink_scheme"
    safe = execute_tool(_request("manage_hyperlink", target=_cell(), arguments={"operation": "set", "address": "https://example.com", "expected_current_address": None}), engine=engine)
    assert safe.ok


def test_chart_connection_refresh_and_calculation_are_explicit():
    engine = FakeEngine(
        {"Data": {}},
        pivot_tables={"Data!Pivot1": {"source_type": "database", "source_data": "Data!A1:B2", "shared_with": []}},
        features={"connections": {"Query1": {}}},
    )
    chart = execute_tool(_request("manage_chart", target=TargetRef("working-copy", sheet="Data", address="E2"), arguments={"operation": "create", "name": "Revenue", "expected_exists": False, "source_address": "A1:B3", "chart_type": "column"}), engine=engine)
    assert chart.ok and "Revenue" in engine.features["charts"]
    assert execute_tool(_request("manage_connection", arguments={"operation": "refresh", "name": "Query1"}), engine=engine).ok
    refreshed = execute_tool(_request("refresh_workbook", arguments={"connection_names": ["Query1"], "pivot_tables": [{"sheet": "Data", "name": "Pivot1"}]}), engine=engine)
    assert refreshed.ok and engine.pivot_tables["Data!Pivot1"]["refreshed"] is True
    assert execute_tool(_request("calculate_workbook", arguments={"full_rebuild": True}), engine=engine).ok


def test_validate_workbook_reports_pass_and_fail():
    engine = FakeEngine({"Data": {}}, features={"names": {"TaxRate": {"refers_to": "=0.15"}}})
    passed = execute_tool(_request("validate_workbook", arguments={"expected_sheet_names": ["Data"], "required_names": ["TaxRate"], "require_no_formula_errors": True}), engine=engine)
    assert passed.ok
    failed = execute_tool(_request("validate_workbook", arguments={"expected_sheet_names": ["Missing"]}), engine=engine)
    assert not failed.ok and failed.error.code == "workbook_validation_failed"
