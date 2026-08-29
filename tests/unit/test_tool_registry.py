from src.tool_registry import tool_catalog


def test_catalog_distinguishes_available_from_planned_tools():
    catalog = tool_catalog()
    by_name = {item["name"]: item for item in catalog["tools"]}
    assert by_name["create_backup"]["status"] == "available"
    assert by_name["write_range"]["status"] == "planned"
    assert by_name["write_range"]["requires_backup"] is True
    assert catalog["available_count"] < catalog["tool_count"]


def test_available_only_catalog_never_exposes_planned_tool_as_callable():
    catalog = tool_catalog(include_planned=False)
    assert all(item["status"] == "available" for item in catalog["tools"])
    assert "write_range" not in {item["name"] for item in catalog["tools"]}
