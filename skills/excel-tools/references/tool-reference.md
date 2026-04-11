# Excel Tools Complete Reference

All 53 tools with full JSON examples. Updated for Phase 1 Remediation (April 11, 2026).

## Phase 1 EditSession Pattern

**NEW**: For mutations, tools now use `EditSession` instead of raw `ExcelAgent`:

```python
from excel_agent.core.edit_session import EditSession

session = EditSession.prepare(input_path, output_path)
with session:
    wb = session.workbook
    # Perform mutations
    version_hash = session.version_hash  # Capture before exit
# EditSession automatically saves ONCE (no double-save bug)
```

**Benefits**:
- Automatic copy-on-write (if input != output)
- Consistent `keep_vba=True` for macro preservation
- Single save on exit (eliminates double-save bug)
- Version hash capture before exit
- File locking integration

---

## Governance Tools (6)

### xls-clone-workbook
**Purpose**: Atomic copy to safe working directory.

**CLI**:
```bash
xls-clone-workbook --input source.xlsx [--output-dir ./work/]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "clone_path": "./work/source_20260411T143022_a3f7e2d1.xlsx",
    "source_hash": "sha256:abc...",
    "clone_hash": "sha256:abc...",
    "timestamp": "20260411T143022"
  }
}
```

---

### xls-validate-workbook
**Purpose**: OOXML compliance, broken refs, circular refs.

**CLI**:
```bash
xls-validate-workbook --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "valid": true,
    "errors": [],
    "warnings": ["Large image in B5"],
    "circular_refs": [],
    "broken_references": 0
  }
}
```

---

### xls-approve-token
**Purpose**: Generate HMAC-SHA256 scoped token.

**CLI**:
```bash
export EXCEL_AGENT_SECRET="your-256-bit-secret"
xls-approve-token --scope sheet:delete --file workbook.xlsx [--ttl 300]
```

**Scopes**: `sheet:delete`, `sheet:rename`, `range:delete`, `formula:convert`, `macro:remove`, `macro:inject`, `structure:modify`

**Important**: Set `EXCEL_AGENT_SECRET` environment variable before generating tokens. This ensures consistent secret across tool invocations (Phase 1 fix).

**Output**:
```json
{
  "status": "success",
  "data": {
    "token": "sheet:delete|sha256:abc...|nonce...|timestamp|300|signature...",
    "scope": "sheet:delete",
    "expires_at": "2026-04-11T14:35:00Z",
    "file_hash": "sha256:abc..."
  }
}
```

---

### xls-version-hash
**Purpose**: Compute geometry hash for modification detection.

**CLI**:
```bash
xls-version-hash --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "geometry_hash": "sha256:abc...",
    "file_hash": "sha256:xyz..."
  }
}
```

---

### xls-lock-status
**Purpose**: Check if workbook is locked.

**CLI**:
```bash
xls-lock-status --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "locked": false,
    "lock_file_exists": false
  }
}
```

---

### xls-dependency-report
**Purpose**: Export formula dependency graph.

**CLI**:
```bash
xls-dependency-report --input workbook.xlsx [--sheet Sheet1]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "stats": {
      "total_cells": 500,
      "total_formulas": 50,
      "total_edges": 120
    },
    "graph": {
      "Sheet1!B1": ["Sheet1!C1"]
    },
    "circular_refs": []
  }
}
```

**Phase 1 Fix**: Large ranges (e.g., full sheet) now properly expanded for impact analysis.

---

## Read Tools (7)

### xls-read-range
**Purpose**: Extract data as JSON.

**CLI**:
```bash
xls-read-range --input workbook.xlsx --range A1:C10 [--sheet Sheet1] [--chunked]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "values": [["Header", "Value"], ["A", 100]],
    "range": "A1:B2",
    "sheet": "Sheet1"
  }
}
```

---

### xls-get-sheet-names
**Purpose**: List all sheets.

**CLI**:
```bash
xls-get-sheet-names --input workbook.xlsx
```

**Output**:
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

---

### xls-get-workbook-metadata
**Purpose**: High-level statistics.

**CLI**:
```bash
xls-get-workbook-metadata --input workbook.xlsx
```

**Output**:
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

---

### xls-get-defined-names
**Purpose**: List all named ranges (global and sheet-scoped) with robust error handling.

**CLI**:
```bash
xls-get-defined-names --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "named_ranges": [
      {"name": "SalesData", "scope": "workbook", "refers_to": "Sheet3!$A$1:$B$5", "hidden": false, "is_reserved": false},
      {"name": "TaxRate", "scope": "Lists", "refers_to": "Lists!$D$2", "hidden": false, "is_reserved": false}
    ],
    "count": 2
  },
  "workbook_version": "sha256:abc..."
}
```

**Notes**:
- Handles workbooks with no named ranges (returns empty list)
- Supports both workbook-level and sheet-level named ranges
- Returns `hidden` and `is_reserved` status for each named range
- Safe for use with all openpyxl versions (Phase 16 fix)

---

### xls-get-table-info
**Purpose**: List Excel Table objects.

**CLI**:
```bash
xls-get-table-info --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "tables": [
      {
        "name": "Table1",
        "sheet": "Sheet1",
        "range": "A1:D10",
        "columns": ["ID", "Name", "Value", "Total"]
      }
    ]
  }
}
```

---

### xls-get-cell-style
**Purpose**: Get style as JSON.

**CLI**:
```bash
xls-get-cell-style --input workbook.xlsx --cell A1 [--sheet Sheet1]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "font": {"name": "Arial", "size": 11, "bold": false},
    "fill": {"fgColor": "FFFFFFFF"},
    "border": {"top": null},
    "alignment": {"horizontal": "left"},
    "number_format": "General"
  }
}
```

---

### xls-get-formula
**Purpose**: Get formula from cell.

**CLI**:
```bash
xls-get-formula --input workbook.xlsx --cell A1 [--sheet Sheet1]
```

**Output**:
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

---

## Write Tools (4) - Uses EditSession

### xls-create-new
**Purpose**: Create blank workbook.

**CLI**:
```bash
xls-create-new --output workbook.xlsx [--sheets "Sheet1,Sheet2"]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "path": "workbook.xlsx",
    "sheets_created": ["Sheet1", "Sheet2"]
  }
}
```

---

### xls-create-from-template
**Purpose**: Clone template with substitution.

**CLI**:
```bash
xls-create-from-template --template template.xltx --output workbook.xlsx \
  --vars '{"company": "Acme", "year": "2026"}'
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "path": "workbook.xlsx",
    "substitutions": 3
  }
}
```

---

### xls-write-range
**Purpose**: Write 2D array to range. Uses EditSession (Phase 1).

**CLI**:
```bash
xls-write-range --input workbook.xlsx --output workbook.xlsx --range A1 \
  --data '[["Name", "Value"], ["A", 100]]' [--sheet Sheet1]
```

**Output**:
```json
{
  "status": "success",
  "data": {"range_written": "A1:B2"},
  "impact": {"cells_modified": 4},
  "workbook_version": "sha256:abc..."
}
```

**Phase 1**: Uses EditSession - no explicit save needed

---

### xls-write-cell
**Purpose**: Write single cell. Uses EditSession (Phase 1).

**CLI**:
```bash
xls-write-cell --input workbook.xlsx --output workbook.xlsx --cell A1 \
  --value "=SUM(B1:B10)" [--type formula] [--sheet Sheet1]
```

**Types**: `auto`, `string`, `number`, `formula`, `date`, `boolean`

**Output**:
```json
{
  "status": "success",
  "data": {"cell": "A1", "type": "formula"},
  "impact": {"cells_modified": 1},
  "workbook_version": "sha256:abc..."
}
```

**Phase 1**: Uses EditSession - no explicit save needed

---

## Structure Tools (8) - ⚠️ Token Required - Uses EditSession

### xls-add-sheet
**CLI**: `xls-add-sheet --input X --output X --name "New" [--position 0]`

**Phase 1**: Uses EditSession - no explicit save needed

---

### xls-delete-sheet ⚠️
**Scope**: `sheet:delete`
**CLI**: `xls-delete-sheet --input X --output X --name "Sheet" --token T [--acknowledge-impact]`

**Denial Output**:
```json
{
  "status": "denied",
  "exit_code": 4,
  "guidance": "Run xls-update-references --updates '[...]' before retrying",
  "impact": {"broken_references": 7}
}
```

**Phase 1**: Uses EditSession, fixed audit log API, dependency tracker properly expands large ranges

---

### xls-rename-sheet ⚠️
**Scope**: `sheet:rename`
**CLI**: `xls-rename-sheet --input X --output X --old "Old" --new "New" --token T`

**Phase 1**: Uses EditSession, fixed audit log API

---

### xls-insert-rows
**CLI**: `xls-insert-rows --input X --output X --sheet S --before-row 5 --count 3`

**Phase 1**: Uses EditSession - no explicit save needed

---

### xls-delete-rows ⚠️
**Scope**: `range:delete`
**CLI**: `xls-delete-rows --input X --output X --sheet S --start-row 5 --count 3 --token T`

**Phase 1**: Uses EditSession, fixed audit log API

---

### xls-insert-columns
**CLI**: `xls-insert-columns --input X --output X --sheet S --before-col C --count 2`

**Phase 1**: Uses EditSession - no explicit save needed

---

### xls-delete-columns ⚠️
**Scope**: `range:delete`
**CLI**: `xls-delete-columns --input X --output X --sheet S --start-col C --count 2 --token T`

**Phase 1**: Uses EditSession, fixed audit log API

---

### xls-move-sheet
**CLI**: `xls-move-sheet --input X --output X --name "Sheet" --position 0`

**Phase 1**: Uses EditSession - no explicit save needed

---

## Cells Tools (4) - Uses EditSession

### xls-merge-cells
**CLI**: `xls-merge-cells --input X --output X --range A1:C1`

**Phase 1**: Uses EditSession

---

### xls-unmerge-cells
**CLI**: `xls-unmerge-cells --input X --output X --range A1:C1`

**Phase 1**: Uses EditSession

---

### xls-delete-range ⚠️
**Scope**: `range:delete`
**CLI**: `xls-delete-range --input X --output X --range A1:C5 --shift up|left --token T`

**Phase 1**: Uses EditSession

---

### xls-update-references
**CLI**: `xls-update-references --input X --output X --updates '[{"old": "...", "new": "..."}]'`

**Phase 1**: Uses EditSession, fixed audit log API

---

## Formulas Tools (6) - Uses EditSession

### xls-set-formula
**CLI**: `xls-set-formula --input X --output X --cell A1 --formula "=SUM(B1:B10)"`

**Phase 1**: Uses EditSession

---

### xls-recalculate
**CLI**: `xls-recalculate --input X --output X [--tier 1|2]`

**Output**:
```json
{
  "status": "success",
  "data": {
    "formula_count": 47,
    "calculated_count": 47,
    "error_count": 0,
    "engine": "tier1_formulas",
    "recalc_time_ms": 45.2
  },
  "workbook_version": "sha256:abc..."
}
```

**Phase 1 Fix**: Tier 1 calculator now preserves original sheet name casing after recalculation. Cross-sheet references work correctly.

---

### xls-detect-errors
**CLI**: `xls-detect-errors --input X [--sheet S]`

**Output**:
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

---

### xls-convert-to-values ⚠️
**Scope**: `formula:convert`
**CLI**: `xls-convert-to-values --input X --output X --range A1:C10 --token T`

**Phase 1**: Uses EditSession

---

### xls-copy-formula-down

**Purpose**: Copy formula from source cell down to target cells with reference adjustment.

**CLI (Preferred API)**: `xls-copy-formula-down --input X --output X --source A1 --target A1:A10`
**CLI (Legacy API)**: `xls-copy-formula-down --input X --output X --cell A1 --count 10`
**Note**: `--source`/`--target` is preferred over `--cell`/`--count` (deprecated but still supported)

**Output**:
```json
{
  "status": "success",
  "data": {
    "source": "A1",
    "target": "A1:A10",
    "filled_count": 9,
    "filled_range": "A2:A10"
  },
  "impact": {
    "cells_modified": 9,
    "formulas_added": 9
  }
}
```

**Examples**:
```bash
# Preferred: Use range syntax
xls-copy-formula-down --input data.xlsx --output data.xlsx --source A1 --target A1:A10

# Legacy: Use count
xls-copy-formula-down --input data.xlsx --output data.xlsx --cell A1 --count 9

# On specific sheet
xls-copy-formula-down --input data.xlsx --sheet Sales --source B2 --target B2:B20
```

**Phase 1 Fixes**:
1. Fixed target count calculation (excluded source row)
2. Fixed regex group indices in formula adjustment

---

### xls-define-name
**CLI**: `xls-define-name --input X --output X --name "SalesData" --refers-to "Sheet1!A1:B10"`

**Phase 1**: Uses EditSession

---

## Objects Tools (5) - Uses EditSession

### xls-add-table
**CLI**: `xls-add-table --input X --output X --range A1:D10 --name "Table1" [--has-totals]`

**Phase 1**: Uses EditSession

---

### xls-add-chart
**CLI**: `xls-add-chart --input X --output X --type bar --data-range "A1:B10" --position "E1"`

**Phase 1**: Uses EditSession

---

### xls-add-image
**CLI**: `xls-add-image --input X --output X --image logo.png --cell A1 [--width 200]`

**Phase 1**: Uses EditSession

---

### xls-add-comment
**CLI**: `xls-add-comment --input X --output X --cell A1 --text "Review"`

**Phase 1**: Uses EditSession

---

### xls-set-data-validation
**CLI**: `xls-set-data-validation --input X --output X --range A1:A10 --type list --source '["Yes", "No"]'`

**Phase 1**: Uses EditSession

---

## Formatting Tools (5) - Uses EditSession

### xls-format-range
**CLI**: `xls-format-range --input X --output X --range A1:C10 --spec '{"font": {"bold": true}}'`

**Phase 1**: Uses EditSession

---

### xls-set-column-width
**CLI**: `xls-set-column-width --input X --output X --column A [--width 20|--auto-fit]`

**Phase 1**: Uses EditSession

---

### xls-freeze-panes
**CLI**: `xls-freeze-panes --input X --output X --row 2 [--column C]`

**Phase 1**: Uses EditSession

---

### xls-apply-conditional-formatting
**CLI**: `xls-apply-conditional-formatting --input X --output X --range A1:A100 --type colorscale --colors '["FF0000", "FFFF00", "00FF00"]'`

**Phase 1**: Uses EditSession

---

### xls-set-number-format
**Purpose**: Apply number formats to cell ranges. **Important**: Format codes with `%` must be properly escaped.

**CLI**: `xls-set-number-format --input X --output X --range A1:A10 --format '$#,##0.00'`

**Common Formats**:
- Currency: `'$'#,##0.00` or `'€'#,##0.00`
- Percentage: `0.00%` (note: internally escaped as `0.00%%`)
- Date: `yyyy-mm-dd` or `mm/dd/yyyy`
- Scientific: `0.00E+00`
- Custom: Any valid Excel format code

**Output**:
```json
{
  "status": "success",
  "data": {
    "range": "A1:A10",
    "format": "$#,##0.00",
    "format_description": "Currency with dollar sign",
    "cells_formatted": 10
  },
  "impact": {"cells_formatted": 10}
}
```

**Phase 1**: Uses EditSession

---

## Macros Tools (5) - ⚠️⚠️ Double Token

### xls-has-macros
**CLI**: `xls-has-macros --input file.xlsm`

---

### xls-inspect-macros
**CLI**: `xls-inspect-macros --input file.xlsm`

---

### xls-validate-macro-safety
**CLI**: `xls-validate-macro-safety --input file.xlsm`

**Output**:
```json
{
  "status": "success",
  "data": {
    "risk_level": "high",
    "auto_exec_triggers": ["AutoOpen"],
    "suspicious_keywords": ["Shell"]
  }
}
```

---

### xls-remove-macros ⚠️⚠️
**Scope**: `macro:remove` × 2
**CLI**: `xls-remove-macros --input file.xlsm --output file.xlsx --token T1 --token T2`

---

### xls-inject-vba-project ⚠️
**Scope**: `macro:inject`
**CLI**: `xls-inject-vba-project --input file.xlsx --output file.xlsm --vba-bin project.bin --token T`

---

## Export Tools (3)

### xls-export-pdf
**CLI**: `xls-export-pdf --input workbook.xlsx --outfile output.pdf [--recalc]`

**Note**: Uses `--outfile` not `--output` (avoids argparse conflict)

**Output**:
```json
{
  "status": "success",
  "data": {"output": "output.pdf", "pages": 3}
}
```

---

### xls-export-csv
**CLI**: `xls-export-csv --input workbook.xlsx --outfile output.csv [--sheet S] [--encoding utf-8]`

**Output**:
```json
{
  "status": "success",
  "data": {"output": "output.csv", "rows": 100}
}
```

---

### xls-export-json
**CLI**: `xls-export-json --input workbook.xlsx --outfile output.json [--orient records|values|columns]`

**Output**:
```json
{
  "status": "success",
  "data": {"output": "output.json", "records": 100}
}
```

---

## Phase 1 Summary

### Tools Migrated to EditSession (4 objects/formatting + 15 others)
- All object tools (xls_add_table, xls_add_chart, xls_add_image, xls_add_comment, xls_set_data_validation)
- All formatting tools (xls_format_range, xls_set_column_width, xls_freeze_panes, xls_apply_conditional_formatting, xls_set_number_format)
- All structure tools
- All cells tools
- All formulas tools
- All write tools

### Tools Fixed
- `xls_delete_sheet.py` - Audit log API + dependency tracker
- `xls_delete_rows.py` - Audit log API
- `xls_delete_columns.py` - Audit log API
- `xls_rename_sheet.py` - Audit log API
- `xls_update_references.py` - Audit log API
- `xls_copy_formula_down.py` - Count calculation + regex
- `tier1_engine.py` - Sheet casing preservation

### Critical Requirements
1. **Set EXCEL_AGENT_SECRET** for token operations
2. **Use EditSession** for mutations (eliminates double-save)
3. **Check exit codes** before parsing JSON
4. **Use --outfile** not --output for export tools

---

**Document Version**: Phase 1 Remediation (April 11, 2026)
