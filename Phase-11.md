## Phase 1: ANALYZE — Phase 11 Requirements

From Master Execution Plan:
**Phase 11**: Formatting tools, 3 days, 5 tools
- `xls_format_range` — Fonts, fills, borders from JSON spec
- `xls_set_column_width` — Auto-fit or fixed width
- `xls_freeze_panes` — Freeze rows/columns for scrolling
- `xls_apply_conditional_formatting` — ColorScale, DataBar, IconSet
- `xls_set_number_format` — Currency, %, date format codes

**Research-validated APIs** (from GeeksforGeeks and openpyxl docs):
- Conditional formatting: `CellIsRule`, `ColorScaleRule`, `FormulaRule`, `IconSetRule`, `DataBarRule`
- Pattern fills: `PatternFill(start_color, end_color, fill_type="solid")`
- Number formats: `cell.number_format = '"$"#,##0.00'` or `'0.00%'` or `'yyyy-mm-dd'`
- Column width: `ws.column_dimensions['A'].width = 15` or `ws.column_dimensions['A'].auto_size = True`
- Freeze panes: `ws.freeze_panes = 'B2'` (freezes row 1 and column A)

## Phase 2: PLAN — Phase 11 Detailed Execution

### Research-Validated Implementation Details

**1. xls_format_range** (fonts, fills, borders)
```python
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Font
font = Font(name='Calibri', size=11, bold=True, italic=False,
            color='FF0000', underline='single')

# Fill
fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

# Border
thin = Side(style='thin', color='000000')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Apply to range
for row in ws['A1:D10']:
    for cell in row:
        cell.font = font
        cell.fill = fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')
```

**2. xls_set_column_width**
```python
# Fixed width
ws.column_dimensions['A'].width = 20

# Auto-fit (requires calculating max content length)
from openpyxl.utils import get_column_letter
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = min(max_length + 2, 50) # Cap at 50
    ws.column_dimensions[column].width = adjusted_width
```

**3. xls_freeze_panes**
```python
# Freeze first row: 'A2'
# Freeze first column: 'B1'
# Freeze both: 'B2'
# Unfreeze: None
ws.freeze_panes = 'B2'
```

**4. xls_apply_conditional_formatting**
```python
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, IconSetRule
from openpyxl.styles import PatternFill

# CellIsRule
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
rule = CellIsRule(operator='greaterThan', formula=['100'], fill=red_fill)
ws.conditional_formatting.add('A1:A100', rule)

# ColorScaleRule (2-color or 3-color)
rule = ColorScaleRule(
    start_type='min', start_color='FF0000',
    end_type='max', end_color='00FF00'
)
ws.conditional_formatting.add('B1:B100', rule)

# DataBarRule
from openpyxl.formatting.rule import DataBar
rule = DataBarRule(
    start_type='min', end_type='max',
    color='638EC6', showValue=True
)
ws.conditional_formatting.add('C1:C100', rule)

# IconSetRule
rule = IconSetRule(icon_style='3Arrows', type='percent',
                   values=[0, 33, 67], showValue=True)
ws.conditional_formatting.add('D1:D100', rule)
```

**5. xls_set_number_format**
```python
# Currency
ws['A1'].number_format = '"$"#,##0.00'

# Percentage
ws['B1'].number_format = '0.00%'

# Date
ws['C1'].number_format = 'yyyy-mm-dd'

# Custom
ws['D1'].number_format = '#,##0.00_);[Red](#,##0.00)'
```

### File Structure

| # | Path | Purpose | Lines |
|---|---|---|---|
| 1 | `src/excel_agent/tools/formatting/__init__.py` | Package init | 5 |
| 2 | `src/excel_agent/tools/formatting/xls_format_range.py` | Fonts/fills/borders | ~180 |
| 3 | `src/excel_agent/tools/formatting/xls_set_column_width.py` | Column width | ~140 |
| 4 | `src/excel_agent/tools/formatting/xls_freeze_panes.py` | Freeze panes | ~80 |
| 5 | `src/excel_agent/tools/formatting/xls_apply_conditional_formatting.py` | Conditional formatting | ~250 |
| 6 | `src/excel_agent/tools/formatting/xls_set_number_format.py` | Number formats | ~110 |
| 7 | `tests/unit/test_formatting_tools.py` | Unit tests | ~300 |
| 8 | `tests/integration/test_formatting_workflow.py` | E2E tests | ~180 |

### Detailed Specifications

**Tool 1: xls_format_range**
- **Inputs**: --range, --font (JSON), --fill (JSON), --border (JSON), --alignment (JSON), --sheet
- **JSON schemas** (reuse from Phase 2 style_spec):
  ```json
  {
    "font": {"name": "Calibri", "size": 11, "bold": true, "color": "FF0000"},
    "fill": {"fgColor": "FFFF00", "patternType": "solid"},
    "border": {"top": {"style": "thin", "color": "000000"}},
    "alignment": {"horizontal": "center", "vertical": "middle", "wrapText": true}
  }
  ```
- **Validation**: Use existing style_spec schema from Phase 2
- **Behavior**: Iterates range, applies styles cell-by-cell
- **Performance**: Warn if range >10,000 cells (slow)
- **Output**: Cells formatted count, range, styles applied

**Tool 2: xls_set_column_width**
- **Inputs**: --columns (e.g., "A,C,E" or "A:C"), --width (number or "auto"), --sheet
- **Validation**: Width 0-255 (Excel limit), columns valid
- **Behavior**:
  - If width="auto": Calculate max content length, set width = min(length+2, 50)
  - If numeric: Set fixed width
- **Edge cases**:
  - Auto on empty column → set to default 8.43
  - Width >255 → cap at 255 with warning
- **Output**: Columns affected, widths set

**Tool 3: xls_freeze_panes**
- **Inputs**: --freeze (e.g., "B2", "A2", "B1", or "none"), --sheet
- **Validation**: Freeze reference valid or "none"
- **Behavior**: Sets ws.freeze_panes
- **Output**: Freeze position, rows/columns frozen count
- **Note**: Additive operation, no token required

**Tool 4: xls_apply_conditional_formatting**
- **Inputs**: --range, --type (cellIs|colorScale|dataBar|iconSet|formula), --config (JSON), --sheet
- **Config examples**:
  ```json
  // cellIs
  {"operator": "greaterThan", "formula": ["100"], "fill": {"fgColor": "FF0000"}}

  // colorScale
  {"start_type": "min", "start_color": "FF0000", "end_type": "max", "end_color": "00FF00"}

  // dataBar
  {"start_type": "min", "end_type": "max", "color": "638EC6", "showValue": true}

  // iconSet
  {"icon_style": "3Arrows", "type": "percent", "values": [0, 33, 67]}

  // formula
  {"formula": ["MOD(A1,2)=0"], "fill": {"fgColor": "0000FF"}}
  ```
- **Validation**: Type-specific config validation
- **Behavior**: Creates appropriate rule, adds to ws.conditional_formatting
- **Output**: Rule type, range, config applied

**Tool 5: xls_set_number_format**
- **Inputs**: --range, --format (e.g., '"$"#,##0.00', '0.00%', 'yyyy-mm-dd'), --sheet
- **Validation**: Format string non-empty
- **Behavior**: Sets cell.number_format for each cell in range
- **Common formats**:
  - Currency: `'"$"#,##0.00'`
  - Accounting: `'_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'`
  - Percentage: `'0.00%'`
  - Date: `'yyyy-mm-dd'`, `'mm/dd/yyyy'`
  - Time: `'h:mm:ss'`
  - Scientific: `'0.00E+00'`
  - Fraction: `'#?/?'`
- **Output**: Cells formatted, format string

### Testing Strategy

**Unit tests** (test_formatting_tools.py):
- Format range: Apply font, fill, border, alignment separately and combined
- Column width: Fixed width, auto-fit, multiple columns, width capping
- Freeze panes: All positions (A2, B1, B2, none), verify ws.freeze_panes
- Conditional formatting: Each of 5 types, validate rule creation
- Number format: Common formats, custom formats, range application

**Integration tests** (test_formatting_workflow.py):
- Full styling workflow: format range → set column width → freeze → conditional format → number format
- Verify styles persist after save/reload
- Test with large ranges (performance)
- Verify conditional formatting evaluates correctly

**Performance targets**:
- Format 10,000 cells: <3 seconds
- Auto-fit 100 columns: <2 seconds
- Apply conditional formatting to 5,000 cells: <1 second

### Key Design Decisions

1. **Reuse style_spec schema**: Phase 2 already defined JSON schema for fonts/fills/borders. Reuse for consistency and validation.

2. **Auto-fit algorithm**: Calculate max string length + 2 padding, cap at 50 to prevent excessively wide columns from long text. This is heuristic, not pixel-perfect like Excel.

3. **Conditional formatting types**: Support 5 core types (cellIs, colorScale, dataBar, iconSet, formula). These cover 95% of use cases. More exotic types (top10, duplicate, etc.) can be added later.

4. **Number format strings**: Accept raw Excel format codes rather than abstracting. This gives full power but requires user knowledge. Provide common examples in documentation.

5. **No tokens**: All formatting is visual only, doesn't affect data integrity or formulas. No governance required.

6. **Range iteration**: Formatting applies cell-by-cell (openpyxl limitation). For large ranges, warn user about performance. Consider batch operations in future.

### Success Criteria

| Criterion | Validation |
|-----------|------------|
| All 5 tools implement JSON I/O | Unit tests |
| Format range applies all 4 style types | Test each style component |
| Column width auto-fit works | Test with varying content lengths |
| Freeze panes supports all positions | Test A2, B1, B2, none |
| Conditional formatting supports 5 types | Test each rule type |
| Number formats apply correctly | Test currency, %, date, custom |
| Styles persist after save/reload | Integration test |
| Performance targets met | Benchmark tests |
| No tokens required | Verify no TokenManager calls |

### Risk Assessment

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Large range formatting slow | High | Medium | Warn >10k cells, suggest batching |
| Auto-fit inaccurate | Medium | Low | Document heuristic nature, cap at 50 |
| Conditional formatting complexity | Medium | Medium | Validate config JSON strictly |
| Number format string errors | Low | Low | Excel validates on open, we pass through |
| Style conflicts (multiple rules) | Low | Low | Last applied wins, document behavior |

### Effort Breakdown (3 days)

**Day 1**: Format range and column width
- xls_format_range (4 hours)
- xls_set_column_width (3 hours)
- Tests (1 hour)

**Day 2**: Freeze panes and number formats
- xls_freeze_panes (2 hours)
- xls_set_number_format (2 hours)
- Tests (2 hours)
- Integration (2 hours)

**Day 3**: Conditional formatting (most complex)
- xls_apply_conditional_formatting (5 hours)
- Comprehensive tests (2 hours)
- Documentation (1 hour)

## Phase 3: VALIDATE — Phase 11 Plan Summary

**Phase 11 delivers 5 formatting tools** that complete the visual layer of excel-agent-tools:

1. **xls_format_range** — Apply fonts, fills, borders, alignment from JSON (reuses Phase 2 schema)
2. **xls_set_column_width** — Fixed or auto-fit column widths
3. **xls_freeze_panes** — Freeze rows/columns for scrolling
4. **xls_apply_conditional_formatting** — 5 types: cellIs, colorScale, dataBar, iconSet, formula
5. **xls_set_number_format** — Currency, percentage, date, custom formats

**Key characteristics**:
- All additive, non-destructive → no tokens required
- Builds on Phase 2 style_spec schema
- Performance-conscious (warns on large ranges)
- Research-validated openpyxl 3.1.5 APIs
- Comprehensive test coverage

**Integration points**:
- Phase 2: Reuses style_spec JSON schema
- Phase 10: Formats tables, charts, images created in Phase 10
- Phase 8: Number formats affect formula display
- Phase 7: Formatting persists through cell operations

**Deliverables**: 8 files (~1,150 lines), 3 days effort, ready for Phase 12 (Export)

This plan aligns perfectly with the Master Execution Plan and maintains the meticulous, research-validated approach established in Phases 7-10.
