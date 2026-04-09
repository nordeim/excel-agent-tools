## Phase 1: ANALYZE — Phase 10 Requirements

From Master Execution Plan and research:

**Phase 10 scope**: 5 object tools, 4 days, builds on all previous phases
- `xls_add_table` — Convert range to Excel Table (ListObject)
- `xls_add_chart` — Bar, Line, Pie, Scatter charts
- `xls_add_image` — Insert image with aspect preservation
- `xls_add_comment` — Threaded comments
- `xls_set_data_validation` — Dropdown lists, numeric constraints

**Technical foundation from research**:
- Tables: `from openpyxl.worksheet.table import Table, TableStyleInfo` — requires displayName, ref (e.g., "A1:D10"), style
- Charts: `from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart` and `Reference` for data ranges
- Images: `from openpyxl.drawing.image import Image` — supports PNG, JPEG, preserves aspect ratio
- Comments: `from openpyxl.comments import Comment` — author, text, width, height
- Data Validation: `from openpyxl.worksheet.datavalidation import DataValidation` — types: list, whole, decimal, date, time, textLength, custom

**Key constraints**:
- All objects are additive (non-destructive) — no tokens required
- Tables create structured references that affect formulas
- Charts reference cell ranges — must validate ranges exist
- Images increase file size significantly — warn on large files
- Comments are metadata — lightweight
- Data validation prevents invalid input — affects user experience

**Integration points**:
- Phase 2 DependencyTracker: Tables create structured references that should be tracked
- Phase 8 Formulas: Charts often reference formula results
- Phase 7 Cell operations: Tables expand/contract with inserted rows

## Phase 2: PLAN — Detailed Phase 10 Execution Roadmap

### Research-Validated API Details

**Tables** (openpyxl 3.1.5):
```python
from openpyxl.worksheet.table import Table, TableStyleInfo
table = Table(displayName="SalesData", ref="A1:D100")
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, 
                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table.tableStyleInfo = style
ws.add_table(table)
```
- displayName must be unique workbook-wide, no spaces, max 255 chars
- ref must be valid range with headers in first row
- 60+ built-in styles: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11

**Charts**:
```python
from openpyxl.chart import BarChart, Reference
chart = BarChart()
chart.type = "col"  # or "bar" for horizontal
chart.style = 10
chart.title = "Sales by Quarter"
data = Reference(ws, min_col=2, min_row=1, max_row=7, max_col=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=7)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "F2")
```
- Chart types: BarChart, LineChart, PieChart, ScatterChart, AreaChart, etc.
- Positioning via cell anchor (e.g., "F2")
- Data must exist before chart creation

**Images**:
```python
from openpyxl.drawing.image import Image
img = Image("logo.png")
img.width = 200  # pixels
img.height = 100
ws.add_image(img, "A1")
```
- Supports PNG, JPEG, BMP, GIF
- Anchor cell determines top-left position
- No automatic resizing — must set width/height manually

**Comments**:
```python
from openpyxl.comments import Comment
comment = Comment("Review this value", "Analyst", width=300, height=50)
ws["B2"].comment = comment
```
- Author string, text content
- Optional width/height in points
- Replaces existing comment if present

**Data Validation**:
```python
from openpyxl.worksheet.datavalidation import DataValidation
dv = DataValidation(type="list", formula1='"Option1,Option2,Option3"', allow_blank=True)
dv.add("B2:B100")
ws.add_data_validation(dv)
```
- Types: "list", "whole", "decimal", "date", "time", "textLength", "custom"
- formula1/formula2 define constraints
- Can show input message and error alert

### File Structure

| # | Path | Purpose | Est. Lines |
|---|---|---|---|
| 1 | `src/excel_agent/tools/objects/__init__.py` | Package init | 5 |
| 2 | `src/excel_agent/tools/objects/xls_add_table.py` | Create Excel Table | ~140 |
| 3 | `src/excel_agent/tools/objects/xls_add_chart.py` | Create charts | ~220 |
| 4 | `src/excel_agent/tools/objects/xls_add_image.py` | Insert images | ~120 |
| 5 | `src/excel_agent/tools/objects/xls_add_comment.py` | Add comments | ~90 |
| 6 | `src/excel_agent/tools/objects/xls_set_data_validation.py` | Data validation | ~160 |
| 7 | `tests/unit/test_table_operations.py` | Table tests | ~180 |
| 8 | `tests/unit/test_chart_operations.py` | Chart tests | ~200 |
| 9 | `tests/unit/test_object_tools.py` | Image/comment/validation tests | ~220 |
| 10 | `tests/integration/test_objects_workflow.py` | E2E tests | ~150 |

### Implementation Specifications

**Tool 1: xls_add_table**
- **Inputs**: --range (e.g., "A1:D100"), --name (displayName), --style (optional), --sheet
- **Validation**: 
  - Range must have headers (first row non-empty)
  - Name must be unique, valid Excel name (no spaces, special chars)
  - Range must exist and contain data
- **Behavior**:
  - Creates Table object with TableStyleInfo
  - Adds to worksheet via ws.add_table()
  - Returns table properties (name, ref, style, column count)
- **Edge cases**: 
  - Name collision → error with suggestion
  - Range overlaps existing table → error
  - Single row (headers only) → warning but allowed
- **JSON schema**: Validates style name against allowed list

**Tool 2: xls_add_chart**
- **Inputs**: --type (bar|line|pie|scatter), --data-range, --categories-range (optional), --title, --position, --sheet
- **Validation**:
  - Data range must exist and contain numeric data
  - Categories range (if provided) must match data dimensions
  - Position must be valid cell reference
- **Behavior**:
  - Creates appropriate chart class
  - Sets data via Reference objects
  - Positions at anchor cell
  - Applies basic styling
- **Edge cases**:
  - Pie chart with multiple series → error (pie supports 1 series only)
  - Scatter requires numeric X and Y → validate
  - Empty data range → error
- **Output**: Chart properties, data range, position

**Tool 3: xls_add_image**
- **Inputs**: --image-path, --position, --width (optional), --height (optional), --sheet
- **Validation**:
  - Image file must exist and be supported format
  - Position must be valid cell
  - Width/height positive if provided
- **Behavior**:
  - Loads image via PIL (openpyxl uses Pillow)
  - Optionally resizes maintaining aspect ratio if only one dimension provided
  - Anchors to cell
  - Warns if file >1MB (impacts workbook size)
- **Edge cases**:
  - Image not found → exit 2
  - Unsupported format → error with supported list
  - Very large image (>5MB) → warning about performance
- **Security**: Validate path to prevent traversal, restrict to allowed directories

**Tool 4: xls_add_comment**
- **Inputs**: --cell, --text, --author (optional, default "excel-agent"), --width, --height, --sheet
- **Validation**:
  - Cell must be valid reference
  - Text non-empty, max 32,767 chars (Excel limit)
- **Behavior**:
  - Creates Comment object
  - Assigns to cell.comment
  - Replaces existing comment if present (warns)
- **Edge cases**:
  - Cell has existing comment → warning, overwrite
  - Very long text → truncate with warning
- **Output**: Cell reference, author, text preview

**Tool 5: xls_set_data_validation**
- **Inputs**: --range, --type, --formula1, --formula2 (optional), --allow-blank, --show-input, --input-title, --input-message, --show-error, --error-title, --error-message, --sheet
- **Validation**:
  - Type must be valid enum
  - Formula1 required for all types
  - Formula2 required for between/notBetween
  - Range must be valid
- **Behavior**:
  - Creates DataValidation object
  - Configures all properties
  - Adds to worksheet
- **Edge cases**:
  - List validation with >255 chars → error (Excel limit)
  - Custom formula invalid → warning but allow (Excel will validate)
  - Overlapping validations → Excel allows, we warn
- **Output**: Validation properties, affected range

### Testing Strategy

**Unit tests**:
- Table: Valid creation, name validation, style application, overlap detection
- Chart: Each type (bar, line, pie, scatter), data validation, positioning
- Image: Load PNG/JPEG, resize, anchor, large file warning
- Comment: Create, replace, author default, size limits
- Validation: Each type, formula parsing, error messages

**Integration tests**:
- Create table → add data → verify structured references work
- Create chart from table data → verify updates when data changes
- Add image + comment + validation to same sheet → verify all persist
- Full workflow: table → chart → image → comment → validation

**Performance tests**:
- Table with 10,000 rows → <2 seconds
- Chart with large data range → <1 second
- Image insertion → <500ms for <1MB file

### Success Criteria

| Criterion | Validation |
|-----------|------------|
| All 5 tools implement JSON I/O | Unit tests verify build_response format |
| Tables support all 60 styles | Test with style enumeration |
| Charts support 4 types | Test each type creation |
| Images preserve aspect ratio | Test resize logic |
| Comments support author attribution | Test default and custom |
| Data validation supports all 7 types | Test each type |
| No tokens required (additive only) | Verify no TokenManager calls |
| All operations logged to audit | Check audit entries |
| File size warnings for large images | Test >1MB image |
| Name collision detection for tables | Test duplicate names |

### Key Design Decisions

1. **No governance tokens**: All 5 operations are additive and non-destructive. They don't break formulas or delete data. This aligns with Master Plan (no ⚠️ markers).

2. **Table name validation**: Excel requires unique table names workbook-wide. We check `wb.tables` before creation and return error with suggestion if collision.

3. **Chart data validation**: We verify data range contains numeric data (for value axes) before creating chart. Prevents empty or text-only charts.

4. **Image size warnings**: Images bloat .xlsx files (embedded as binary). We warn at 1MB and 5MB thresholds to help agents manage file size.

5. **Comment replacement**: Excel allows one comment per cell. We overwrite with warning rather than error, as this is expected behavior.

6. **Data validation limits**: Excel has 255-char limit for list validations. We validate and return error before attempting to add, preventing cryptic openpyxl errors.

7. **Style enumeration**: Rather than free-form style strings, we validate against known TableStyle names to prevent silent failures.

### Risk Assessment

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Table name collision | Medium | Medium | Pre-check wb.tables, suggest alternative |
| Chart with invalid data | Low | Low | Validate range contains numbers |
| Large image bloats file | Medium | High | Warn at 1MB/5MB, suggest compression |
| Data validation formula error | Low | Medium | Allow but warn, Excel validates on open |
| Memory usage with many objects | Low | Medium | Test with 100+ objects, monitor |

### Effort Breakdown (4 days)

**Day 1**: Tables and Charts (core objects)
- xls_add_table implementation + tests (4 hours)
- xls_add_chart implementation + tests (4 hours)

**Day 2**: Images and Comments
- xls_add_image implementation + tests (3 hours)
- xls_add_comment implementation + tests (3 hours)
- Integration testing (2 hours)

**Day 3**: Data Validation
- xls_set_data_validation implementation + tests (6 hours)
- Edge case testing (2 hours)

**Day 4**: Integration and Hardening
- Full workflow tests (3 hours)
- Performance testing (2 hours)
- Documentation and examples (3 hours)

## Phase 3: VALIDATE — Confirmation Required

Before I proceed to implementation, confirm these design decisions:

1. **Table structured references**: Should we automatically update formulas that reference the table range to use structured references (e.g., `SalesData[Amount]` instead of `D2:D100`), or leave that to the agent via `xls_update_references`

2. **Chart styling**: Should we support custom colors and advanced styling via JSON, or keep to basic styles (title, type, position) for v1

3. **Image storage**: Should we support images from URLs (download then embed), or only local file paths for v1

4. **Comment threading**: Excel now supports threaded comments (modern) vs. notes (legacy). Should we default to threaded comments, and should we support replies

5. **Data validation lists**: Should we support list validations that reference a range (e.g., `=$Z$1:$Z$10`) in addition to inline lists, and should we validate that the referenced range exists

Once confirmed, proceed to Phase 4 IMPLEMENT and generate all 10 files with complete, tested code following the established patterns from Phases 7-9.
