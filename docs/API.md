# API Reference: excel-agent-tools

**Version:** 1.0.0  
**Total Tools:** 53

---

## Quick Reference

| Category | Tools | Description |
|----------|-------|-------------|
| **Governance** (6) | clone, validate, approve_token, version_hash, lock_status, dependency_report | Safety and governance |
| **Read** (7) | get_sheet_names, get_workbook_metadata, read_range, get_defined_names, get_table_info, get_formula, get_cell_style | Data introspection |
| **Write** (4) | create_new, create_from_template, write_range, write_cell | Data creation/modification |
| **Structure** (8) | add_sheet, delete_sheet ⚠️, rename_sheet ⚠️, insert_rows, delete_rows ⚠️, insert_columns, delete_columns ⚠️, move_sheet | Structural mutation |
| **Cells** (4) | merge_cells, unmerge_cells, delete_range ⚠️, update_references | Cell operations |
| **Formulas** (6) | set_formula, recalculate, detect_errors, convert_to_values ⚠️, copy_formula_down, define_name | Formula management |
| **Objects** (5) | add_table, add_chart, add_image, add_comment, set_data_validation | Object insertion |
| **Formatting** (5) | format_range, set_column_width, freeze_panes, apply_conditional_formatting, set_number_format | Styling |
| **Macros** (5) | has_macros, inspect_macros, validate_macro_safety, remove_macros ⚠️⚠️, inject_vba_project ⚠️ | VBA safety |
| **Export** (3) | export_pdf, export_csv, export_json | Interop |

**Legend:** ⚠️ = Token required | ⚠️⚠️ = Double-token required

---

## Universal Response Schema

All tools return JSON with the following envelope:

```json
{
  "status": "success" | "error" | "warning" | "denied",
  "data": { ... },  // Tool-specific data
  "impact": {
    "cells_modified": 0,
    "formulas_updated": 0,
    "rows_inserted": 0,
    "rows_deleted": 0
  },
  "warnings": [],
  "timestamp": "2026-04-08T14:30:22Z",
  "guidance": "..."  // Present when status="denied"
}
```

## Exit Codes

| Code | Meaning | When Used |
|------|---------|-----------|
| 0 | Success | Operation completed |
| 1 | Validation Error | Bad input, impact denied |
| 2 | File Not Found | Path doesn't exist |
| 3 | Lock Contention | File locked by another process |
| 4 | Permission Denied | Invalid/missing token |
| 5 | Internal Error | Unexpected exception |

---

# Governance Tools

## xls-clone-workbook

**Purpose:** Create atomic copy of workbook to safe working directory.

**CLI:**
```bash
xls-clone-workbook --input path.xlsx [--output-dir ./work/]
```

**Input:**
- `--input` (required): Source workbook path
- `--output-dir` (optional): Destination directory (default: ./work/)

**Output:**
```json
{
  "status": "success",
  "data": {
    "clone_path": "/work/data_20260409T143022_a3f7e2d1.xlsx",
    "source_hash": "sha256:abc...",
    "clone_hash": "sha256:abc...",
    "timestamp": "20260409T143022"
  }
}
```

**Exit Codes:** 0, 2

**Agent Note:** Always work on clones. Source file never modified.

---

## xls-validate-workbook

**Purpose:** OOXML compliance check, broken reference detection, circular ref scan.

**CLI:**
```bash
xls-validate-workbook --input path.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "valid": true,
    "errors": [],
    "warnings": ["Sheet2 has 3 blank cells with formulas"],
    "circular_refs": [],
    "broken_references": 0
  }
}
```

**Exit Codes:** 0, 1, 2

---

## xls-approve-token

**Purpose:** Generate HMAC-SHA256 scoped approval token for destructive operations.

**CLI:**
```bash
xls-approve-token --scope sheet:delete --file path.xlsx [--ttl 300]
```

**Valid Scopes:**
- `sheet:delete`, `sheet:rename`
- `range:delete`
- `formula:convert`
- `macro:remove`, `macro:inject`
- `structure:modify`

**Input:**
- `--scope` (required): Token scope
- `--file` (required): Target workbook
- `--ttl` (optional): Time-to-live seconds (1-3600, default: 300)

**Output:**
```json
{
  "status": "success",
  "data": {
    "token": "eyJzY29wZSI6InNoZWV0OmRlbGV0ZSIs...",
    "scope": "sheet:delete",
    "expires_at": "2026-04-08T14:35:22Z",
    "file_hash": "sha256:abc..."
  }
}
```

**Exit Codes:** 0, 1

**Security:** Token bound to file hash. Cannot be reused on different file.

---

## xls-version-hash

**Purpose:** Compute geometry hash for concurrent modification detection.

**CLI:**
```bash
xls-version-hash --input path.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "geometry_hash": "sha256:abc...",
    "file_hash": "sha256:xyz..."
  }
}
```

**Exit Codes:** 0, 2

---

## xls-lock-status

**Purpose:** Check if workbook is currently locked by another process.

**CLI:**
```bash
xls-lock-status --input path.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "locked": false,
    "lock_file_exists": false
  }
}
```

**Exit Codes:** 0, 2

---

## xls-dependency-report

**Purpose:** Export full formula dependency graph as JSON adjacency list.

**CLI:**
```bash
xls-dependency-report --input path.xlsx [--sheet Sheet1]
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "stats": {
      "total_cells": 500,
      "total_formulas": 50,
      "total_edges": 120,
      "circular_chains": 0
    },
    "graph": {
      "Sheet1!B1": ["Sheet1!C1", "Sheet1!D1"],
      "Sheet1!A1": ["Sheet1!B1"]
    },
    "circular_refs": []
  }
}
```

**Exit Codes:** 0, 2

---

# Read Tools

## xls-get-sheet-names

**Purpose:** List all sheets with index, name, and visibility.

**CLI:**
```bash
xls-get-sheet-names --input path.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "sheets": [
      {"index": 0, "name": "Sheet1", "visibility": "visible"},
      {"index": 1, "name": "Data", "visibility": "hidden"}
    ]
  }
}
```

**Exit Codes:** 0, 2

---

## xls-get-workbook-metadata

**Purpose:** High-level workbook statistics.

**CLI:**
```bash
xls-get-workbook-metadata --input path.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "sheet_count": 3,
    "total_formulas": 47,
    "named_ranges": ["SalesData"],
    "tables": ["Table1"],
    "has_macros": false,
    "file_size_bytes": 15234
  }
}
```

**Exit Codes:** 0, 2

---

## xls-read-range

**Purpose:** Extract data as JSON array. Supports chunked streaming.

**CLI:**
```bash
xls-read-range --input path.xlsx --range A1:C10 [--sheet Sheet1] [--chunked]
```

**Input:**
- `--range` (required): A1 notation (e.g., "A1:C10", "A1", "A:A")
- `--sheet` (optional): Sheet name (default: active sheet)
- `--chunked` (flag): Emit JSONL for large datasets

**Output (normal mode):**
```json
{
  "status": "success",
  "data": {
    "values": [
      ["Name", "Value", "Doubled"],
      ["Item 1", 10, "=B2*2"]
    ],
    "range": "A1:C2",
    "sheet": "Sheet1"
  }
}
```

**Output (chunked mode - JSONL):**
```json
{"chunk": 1, "total_chunks": 5, "rows": [...]}
{"chunk": 2, "total_chunks": 5, "rows": [...]}
```

**Exit Codes:** 0, 1, 2

**Agent Note:** Use `--chunked` for >100k rows. Dates returned as ISO 8601.

---

## xls-get-defined-names

**Purpose:** List global and sheet-scoped named ranges.

**CLI:**
```bash
xls-get-defined-names --input path.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "named_ranges": [
      {"name": "SalesData", "scope": "workbook", "refers_to": "Sheet3!$A$1:$B$5"},
      {"name": "TaxRate", "scope": "Sheet1", "refers_to": "Sheet1!$E$1"}
    ]
  }
}
```

**Exit Codes:** 0, 2

---

## xls-get-table-info

**Purpose:** List Excel Table objects (ListObjects) with metadata.

**CLI:**
```bash
xls-get-table-info --input path.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "tables": [
      {
        "name": "Table1",
        "sheet": "Sheet1",
        "range": "A1:D10",
        "columns": ["ID", "Name", "Value", "Total"],
        "has_totals_row": true,
        "style": "TableStyleMedium2"
      }
    ]
  }
}
```

**Exit Codes:** 0, 2

---

## xls-get-formula

**Purpose:** Get formula string from specific cell with parsed references.

**CLI:**
```bash
xls-get-formula --input path.xlsx --cell A1 [--sheet Sheet1]
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "cell": "A1",
    "formula": "=SUM(B1:B10)",
    "references": ["B1:B10"]
  }
}
```

**Exit Codes:** 0, 1, 2

---

## xls-get-cell-style

**Purpose:** Get full style specification as JSON.

**CLI:**
```bash
xls-get-cell-style --input path.xlsx --cell A1 [--sheet Sheet1]
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "font": {"name": "Arial", "size": 11, "bold": false, "color": "FF000000"},
    "fill": {"fgColor": "FFFFFFFF", "patternType": "solid"},
    "border": {"top": null, "bottom": null, "left": null, "right": null},
    "alignment": {"horizontal": "left", "vertical": "center", "wrapText": false},
    "number_format": "General"
  }
}
```

**Exit Codes:** 0, 1, 2

---

# Write Tools

## xls-create-new

**Purpose:** Create blank workbook with specified sheets.

**CLI:**
```bash
xls-create-new --output path.xlsx [--sheets "Sheet1,Sheet2,Data"]
```

**Input:**
- `--sheets` (optional): Comma-separated sheet names (default: "Sheet1")

**Output:**
```json
{
  "status": "success",
  "data": {
    "path": "/path/to/output.xlsx",
    "sheets_created": ["Sheet1", "Sheet2", "Data"]
  }
}
```

**Exit Codes:** 0, 1

---

## xls-create-from-template

**Purpose:** Clone from .xltx/.xltm template with placeholder substitution.

**CLI:**
```bash
xls-create-from-template --template template.xltx --output path.xlsx \
  [--vars '{"company": "Acme", "year": "2026"}']
```

**Placeholder Format:** `{{variable}}`

**Output:**
```json
{
  "status": "success",
  "data": {
    "path": "/path/to/output.xlsx",
    "substitutions": 3
  }
}
```

**Exit Codes:** 0, 1, 2

---

## xls-write-range

**Purpose:** Write 2D JSON array to cell range with type inference.

**CLI:**
```bash
xls-write-range --input path.xlsx --output path.xlsx --range A1 \
  --data '[["Name", "Value"], ["Item", 42]]' [--sheet Sheet1]
```

**Type Inference:**
- Strings: `"text"`
- Numbers: `42`, `3.14`
- Booleans: `true`, `false`
- Formulas: strings starting with `=`
- Dates: ISO 8601 strings → Excel date
- Null: `null` → empty cell

**Output:**
```json
{
  "status": "success",
  "data": {"range_written": "A1:B2"},
  "impact": {"cells_modified": 4}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-write-cell

**Purpose:** Write single cell with explicit type control.

**CLI:**
```bash
xls-write-cell --input path.xlsx --output path.xlsx --cell A1 \
  --value "=SUM(B1:B10)" [--type formula] [--sheet Sheet1]
```

**Types:** `auto` (default), `string`, `number`, `formula`, `date`, `boolean`

**Output:**
```json
{
  "status": "success",
  "data": {"cell": "A1", "type": "formula"},
  "impact": {"cells_modified": 1}
}
```

**Exit Codes:** 0, 1, 2

---

# Structure Tools

## xls-add-sheet

**Purpose:** Add new sheet at specified position.

**CLI:**
```bash
xls-add-sheet --input path.xlsx --output path.xlsx --name "NewSheet" [--position 0]
```

**Output:**
```json
{
  "status": "success",
  "data": {"name": "NewSheet", "index": 0}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-delete-sheet ⚠️

**Purpose:** Delete sheet with dependency check.

**CLI:**
```bash
xls-delete-sheet --input path.xlsx --output path.xlsx --name "SheetName" \
  --token <token> [--acknowledge-impact]
```

**Required Token:** `sheet:delete`

**Output:**
```json
{
  "status": "success",
  "data": {"deleted_sheet": "SheetName"},
  "impact": {"formulas_updated": 0}
}
```

**Denial Output:**
```json
{
  "status": "denied",
  "exit_code": 1,
  "denial_reason": "Operation would break 7 formula references across 3 sheets",
  "guidance": "Run xls-update-references.py --updates '[{\"old\": \"Sheet1!A1\", \"new\": \"Sheet2!A1\"}]' before retrying",
  "impact": {"broken_references": 7, "affected_sheets": ["Sheet1", "Sheet2", "Summary"]}
}
```

**Exit Codes:** 0, 1, 2, 4

---

## xls-rename-sheet ⚠️

**Purpose:** Rename sheet with automatic cross-sheet reference update.

**CLI:**
```bash
xls-rename-sheet --input path.xlsx --output path.xlsx --old "OldName" --new "NewName" \
  --token <token>
```

**Required Token:** `sheet:rename`

**Output:**
```json
{
  "status": "success",
  "data": {"old_name": "OldName", "new_name": "NewName"},
  "impact": {"formulas_updated": 5}
}
```

**Exit Codes:** 0, 1, 2, 4

---

## xls-insert-rows

**Purpose:** Insert rows with style inheritance.

**CLI:**
```bash
xls-insert-rows --input path.xlsx --output path.xlsx --sheet Sheet1 \
  --before-row 5 --count 3 [--inherit-style]
```

**Output:**
```json
{
  "status": "success",
  "data": {"sheet": "Sheet1", "inserted_at": 5, "count": 3},
  "impact": {"formulas_updated": 12}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-delete-rows ⚠️

**Purpose:** Delete rows with impact report.

**CLI:**
```bash
xls-delete-rows --input path.xlsx --output path.xlsx --sheet Sheet1 \
  --start-row 5 --count 3 --token <token> [--acknowledge-impact]
```

**Required Token:** `range:delete`

**Output:**
```json
{
  "status": "success",
  "data": {"deleted_rows": "5:7"},
  "impact": {"cells_modified": 15, "formulas_updated": 8}
}
```

**Exit Codes:** 0, 1, 2, 4

---

## xls-insert-columns

**Purpose:** Insert columns.

**CLI:**
```bash
xls-insert-columns --input path.xlsx --output path.xlsx --sheet Sheet1 \
  --before-col C --count 2
```

**Output:**
```json
{
  "status": "success",
  "data": {"inserted_at": 3, "count": 2}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-delete-columns ⚠️

**Purpose:** Delete columns with impact report.

**CLI:**
```bash
xls-delete-columns --input path.xlsx --output path.xlsx --sheet Sheet1 \
  --start-col C --count 2 --token <token> [--acknowledge-impact]
```

**Required Token:** `range:delete`

**Exit Codes:** 0, 1, 2, 4

---

## xls-move-sheet

**Purpose:** Reorder sheet position.

**CLI:**
```bash
xls-move-sheet --input path.xlsx --output path.xlsx --name "SheetName" --position 0
```

**Output:**
```json
{
  "status": "success",
  "data": {"name": "SheetName", "new_index": 0}
}
```

**Exit Codes:** 0, 1, 2

---

# Cell Tools

## xls-merge-cells

**Purpose:** Merge cell range.

**CLI:**
```bash
xls-merge-cells --input path.xlsx --output path.xlsx --range A1:C1 [--sheet Sheet1]
```

**Output:**
```json
{
  "status": "success",
  "data": {"merged": "A1:C1"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-unmerge-cells

**Purpose:** Restore grid from merged range.

**CLI:**
```bash
xls-unmerge-cells --input path.xlsx --output path.xlsx --range A1:C1 [--sheet Sheet1]
```

**Output:**
```json
{
  "status": "success",
  "data": {"unmerged": "A1:C1"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-delete-range ⚠️

**Purpose:** Delete range with shift direction.

**CLI:**
```bash
xls-delete-range --input path.xlsx --output path.xlsx --range A1:C5 \
  --shift up|left --token <token> [--acknowledge-impact]
```

**Required Token:** `range:delete`

**Output:**
```json
{
  "status": "success",
  "data": {"deleted": "A1:C5", "shift": "up"}
}
```

**Exit Codes:** 0, 1, 2, 4

---

## xls-update-references

**Purpose:** Batch update cell references after structural changes.

**CLI:**
```bash
xls-update-references --input path.xlsx --output path.xlsx \
  --updates '[{"old": "Sheet1!A1", "new": "Sheet2!B2"}]'
```

**Output:**
```json
{
  "status": "success",
  "data": {"formulas_updated": 5}
}
```

**Exit Codes:** 0, 1, 2

---

# Formula Tools

## xls-set-formula

**Purpose:** Inject formula with syntax validation.

**CLI:**
```bash
xls-set-formula --input path.xlsx --output path.xlsx --cell A1 --formula "=SUM(B1:B10)"
```

**Output:**
```json
{
  "status": "success",
  "data": {"cell": "A1", "formula": "=SUM(B1:B10)"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-recalculate

**Purpose:** Recalculate all formulas (auto: Tier 1 → Tier 2).

**CLI:**
```bash
xls-recalculate --input path.xlsx --output path.xlsx [--tier 1|2]
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "formula_count": 47,
    "calculated_count": 47,
    "error_count": 0,
    "engine": "tier1_formulas",
    "recalc_time_ms": 45.2
  }
}
```

**Exit Codes:** 0, 1, 2, 5

---

## xls-detect-errors

**Purpose:** Scan for formula errors.

**CLI:**
```bash
xls-detect-errors --input path.xlsx [--sheet Sheet1]
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "errors": [
      {"sheet": "Sheet1", "cell": "A1", "error": "#REF!", "formula": "=Sheet2!A1"}
    ]
  }
}
```

**Exit Codes:** 0, 2

---

## xls-convert-to-values ⚠️

**Purpose:** Replace formulas with calculated values (irreversible).

**CLI:**
```bash
xls-convert-to-values --input path.xlsx --output path.xlsx --range A1:C10 \
  --token <token>
```

**Required Token:** `formula:convert`

**Output:**
```json
{
  "status": "success",
  "data": {"cells_converted": 10}
}
```

**Exit Codes:** 0, 1, 2, 4

---

## xls-copy-formula-down

**Purpose:** Auto-fill formula with reference adjustment.

**CLI:**
```bash
xls-copy-formula-down --input path.xlsx --output path.xlsx --cell A1 --count 10
```

**Output:**
```json
{
  "status": "success",
  "data": {"source": "A1", "filled": "A2:A11"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-define-name

**Purpose:** Create or update named range.

**CLI:**
```bash
xls-define-name --input path.xlsx --output path.xlsx --name "SalesData" --refers-to "Sheet1!A1:B10"
```

**Output:**
```json
{
  "status": "success",
  "data": {"name": "SalesData", "refers_to": "Sheet1!A1:B10"}
}
```

**Exit Codes:** 0, 1, 2

---

# Object Tools

## xls-add-table

**Purpose:** Convert range to Excel Table (ListObject).

**CLI:**
```bash
xls-add-table --input path.xlsx --output path.xlsx --range A1:D10 --name "Table1" \
  [--has-totals] [--style TableStyleMedium2]
```

**Output:**
```json
{
  "status": "success",
  "data": {"name": "Table1", "range": "A1:D10", "columns": ["A", "B", "C", "D"]}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-add-chart

**Purpose:** Add chart (Bar, Line, Pie, Scatter).

**CLI:**
```bash
xls-add-chart --input path.xlsx --output path.xlsx --type bar \
  --data-range "A1:B10" --position "E1" [--title "Sales Chart"]
```

**Chart Types:** `bar`, `line`, `pie`, `scatter`

**Output:**
```json
{
  "status": "success",
  "data": {"chart_id": 1, "type": "bar", "position": "E1"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-add-image

**Purpose:** Insert image with aspect preservation.

**CLI:**
```bash
xls-add-image --input path.xlsx --output path.xlsx --image logo.png --cell A1 \
  [--width 200] [--height-auto]
```

**Output:**
```json
{
  "status": "success",
  "data": {"image": "logo.png", "cell": "A1", "size": [200, 50]}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-add-comment

**Purpose:** Add threaded comment to cell.

**CLI:**
```bash
xls-add-comment --input path.xlsx --output path.xlsx --cell A1 --text "Review this value"
```

**Output:**
```json
{
  "status": "success",
  "data": {"cell": "A1", "comment_id": 1}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-set-data-validation

**Purpose:** Add dropdown lists, numeric constraints.

**CLI:**
```bash
# Dropdown list
xls-set-data-validation --input path.xlsx --output path.xlsx --range A1:A10 \
  --type list --source '["Yes", "No", "Maybe"]'

# Numeric constraint
xls-set-data-validation --input path.xlsx --output path.xlsx --range B1:B10 \
  --type numeric --min 0 --max 100
```

**Validation Types:** `list`, `numeric`, `date`, `text_length`

**Output:**
```json
{
  "status": "success",
  "data": {"range": "A1:A10", "type": "list", "source_count": 3}
}
```

**Exit Codes:** 0, 1, 2

---

# Formatting Tools

## xls-format-range

**Purpose:** Apply comprehensive formatting from JSON spec.

**CLI:**
```bash
xls-format-range --input path.xlsx --output path.xlsx --range A1:C10 \
  --spec '{"font": {"bold": true}, "fill": {"fgColor": "FFFF00"}}'
```

**Spec Schema:**
```json
{
  "font": {"name": "Arial", "size": 12, "bold": true, "italic": false, "color": "FF0000"},
  "fill": {"fgColor": "FFFF00", "patternType": "solid"},
  "border": {
    "top": {"style": "thin", "color": "000000"},
    "bottom": {"style": "medium", "color": "000000"},
    "left": {"style": "thin", "color": "000000"},
    "right": {"style": "thin", "color": "000000"}
  },
  "alignment": {"horizontal": "center", "vertical": "center", "wrapText": true}
}
```

**Output:**
```json
{
  "status": "success",
  "data": {"formatted": "A1:C10"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-set-column-width

**Purpose:** Set column width (auto-fit or fixed).

**CLI:**
```bash
# Fixed width
xls-set-column-width --input path.xlsx --output path.xlsx --column A --width 20

# Auto-fit
xls-set-column-width --input path.xlsx --output path.xlsx --column A --auto-fit
```

**Output:**
```json
{
  "status": "success",
  "data": {"column": "A", "width": 20}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-freeze-panes

**Purpose:** Freeze rows/columns for scrolling.

**CLI:**
```bash
xls-freeze-panes --input path.xlsx --output path.xlsx --row 2 [--column C]
```

**Output:**
```json
{
  "status": "success",
  "data": {"frozen_row": 2, "frozen_col": 3}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-apply-conditional-formatting

**Purpose:** Apply ColorScale, DataBar, IconSet.

**CLI:**
```bash
# Color scale
xls-apply-conditional-formatting --input path.xlsx --output path.xlsx --range A1:A10 \
  --type colorscale --colors '["FF0000", "FFFF00", "00FF00"]'

# Data bar
xls-apply-conditional-formatting --input path.xlsx --output path.xlsx --range B1:B10 \
  --type databar --color "638EC6"
```

**Types:** `colorscale`, `databar`, `iconset`

**Output:**
```json
{
  "status": "success",
  "data": {"range": "A1:A10", "type": "colorscale"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-set-number-format

**Purpose:** Apply number format codes (currency, %, date).

**CLI:**
```bash
xls-set-number-format --input path.xlsx --output path.xlsx --range A1:A10 \
  --format '$#,##0.00'
```

**Common Formats:**
- Currency: `$#,##0.00`, `€#,##0.00`
- Percentage: `0.00%`
- Date: `YYYY-MM-DD`, `MM/DD/YYYY`
- Number: `#,##0.000`, `0.00E+00`

**Output:**
```json
{
  "status": "success",
  "data": {"range": "A1:A10", "format": "$#,##0.00"}
}
```

**Exit Codes:** 0, 1, 2

---

# Macro Tools

## xls-has-macros

**Purpose:** Boolean check for VBA presence.

**CLI:**
```bash
xls-has-macros --input path.xlsm
```

**Output:**
```json
{
  "status": "success",
  "data": {"has_macros": true, "macro_count": 3}
}
```

**Exit Codes:** 0, 2

---

## xls-inspect-macros

**Purpose:** List VBA modules + signature status.

**CLI:**
```bash
xls-inspect-macros --input path.xlsm
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "modules": [
      {"name": "Module1", "type": "standard", "code_size": 2048},
      {"name": "ThisWorkbook", "type": "document", "code_size": 512}
    ],
    "has_signature": false
  }
}
```

**Exit Codes:** 0, 2

---

## xls-validate-macro-safety

**Purpose:** Risk scan for auto-exec, Shell, IOCs.

**CLI:**
```bash
xls-validate-macro-safety --input path.xlsm
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "risk_level": "high",
    "auto_exec_triggers": ["AutoOpen"],
    "suspicious_keywords": ["Shell", "CreateObject"],
    "iocs": [],
    "recommendation": "Remove or digitally sign before processing"
  }
}
```

**Exit Codes:** 0, 2

---

## xls-remove-macros ⚠️⚠️

**Purpose:** Strip VBA project (converts .xlsm → .xlsx).

**CLI:**
```bash
xls-remove-macros --input path.xlsm --output path.xlsx --token <token1> --token <token2>
```

**Required Tokens:** 2× `macro:remove` (double-token for critical operation)

**Output:**
```json
{
  "status": "success",
  "data": {"output": "path.xlsx", "macros_removed": 3}
}
```

**Exit Codes:** 0, 1, 2, 4

---

## xls-inject-vba-project ⚠️

**Purpose:** Inject pre-extracted .bin file.

**CLI:**
```bash
xls-inject-vba-project --input path.xlsx --output path.xlsm --vba-bin project.bin \
  --token <token> --scan-safety
```

**Required Token:** `macro:inject`

**Required:** `--scan-safety` flag (must pass safety scan first)

**Output:**
```json
{
  "status": "success",
  "data": {"output": "path.xlsm", "modules_injected": 2}
}
```

**Exit Codes:** 0, 1, 2, 4

---

# Export Tools

## xls-export-pdf

**Purpose:** Export via LibreOffice headless.

**CLI:**
```bash
xls-export-pdf --input path.xlsx --outfile output.pdf [--recalc] [--sheet Sheet1]
```

**Note:** Uses `--outfile` not `--output` (avoid argparse conflict).

**Output:**
```json
{
  "status": "success",
  "data": {"output": "output.pdf", "pages": 3}
}
```

**Exit Codes:** 0, 2, 5

---

## xls-export-csv

**Purpose:** Sheet to CSV with encoding control.

**CLI:**
```bash
xls-export-csv --input path.xlsx --outfile output.csv [--sheet Sheet1] [--encoding utf-8]
```

**Output:**
```json
{
  "status": "success",
  "data": {"output": "output.csv", "rows": 100, "encoding": "utf-8"}
}
```

**Exit Codes:** 0, 1, 2

---

## xls-export-json

**Purpose:** Sheet/range to structured JSON.

**CLI:**
```bash
xls-export-json --input path.xlsx --outfile output.json \
  [--format records|values|columns] [--sheet Sheet1] [--range A1:C10]
```

**Formats:**
- `records`: `[{"col1": "val", "col2": "val"}, ...]`
- `values`: `[["val", "val"], ...]`
- `columns`: `{"col1": ["val", "val"], ...}`

**Output:**
```json
{
  "status": "success",
  "data": {"output": "output.json", "records": 100, "format": "records"}
}
```

**Exit Codes:** 0, 1, 2

---

# Common Patterns

## Pattern 1: Clone-Modify-Save

```bash
# 1. Clone
CLONE=$(xls-clone-workbook --input data.xlsx --output-dir ./work/ | jq -r '.data.clone_path')

# 2. Modify
xls-write-range --input "$CLONE" --output "$CLONE" --range A1 --data '[["New", "Data"]]'

# 3. Validate
xls-validate-workbook --input "$CLONE"

# 4. Export
xls-export-csv --input "$CLONE" --outfile output.csv
```

## Pattern 2: Safe Structural Edit

```bash
# 1. Get impact report
xls-dependency-report --input data.xlsx

# 2. Get token
TOKEN=$(xls-approve-token --scope sheet:delete --file data.xlsx | jq -r '.data.token')

# 3. Attempt deletion (may be denied with guidance)
RESULT=$(xls-delete-sheet --input data.xlsx --output data.xlsx --name OldSheet --token "$TOKEN")

# 4. If denied, fix references per guidance, then retry with --acknowledge-impact
# ...parse guidance...
xls-delete-sheet --input data.xlsx --output data.xlsx --name OldSheet --token "$TOKEN" --acknowledge-impact
```

## Pattern 3: Large Dataset Processing

```bash
# Use chunked mode for >100k rows
xls-read-range --input large.xlsx --range A1:E100000 --chunked > output.jsonl

# Process in batches
xls-write-range --input large.xlsx --output large.xlsx --range A1 --data @<(chunk_processor)
```

---

**Document maintained by:** excel-agent-tools maintainers  
**License:** MIT
