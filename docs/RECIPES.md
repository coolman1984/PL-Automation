# Recipe Examples

Recipes describe business intent and compose declared tools. They do not hold
COM calls or manipulate application state directly.

## Safe range update

```json
{
  "schema_version": "1.0",
  "transaction_id": "run-20260830-001",
  "tool": "range.write_values",
  "target": {"workbook_id": "working-copy", "sheet": "Data", "address": "A2:C4"},
  "arguments": {"values": [[1, "A", true], [2, "B", false], [3, "C", true]]},
  "preconditions": [{"type": "range_shape", "rows": 3, "columns": 3}],
  "expected_effect": {"cells_touched": 9},
  "dry_run": true
}
```

The executor must resolve the target, check the shape, return a dry-run result,
and only then allow the same request on a verified working copy.

## Safe formula change

Require the exact sheet and range, preserve formula dialect (`Formula` or
`Formula2`), capture old formulas, write in bulk, recalculate only when the
plan requests it, and validate both the new formulas and unrelated references.

## Business recipe rule

The P&L A08 recipe may decide *which* cells represent an August actual column,
but universal tools decide *how* to inspect, copy, insert, write, calculate,
validate, and publish them. This separation lets a new report reuse the safety
engine without rebuilding it.

