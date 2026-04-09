## Phase 1: ANALYZE — Phase 12 Requirements

From Master Execution Plan:
**Phase 12**: Export tools, 2 days, 3 tools
- `xls_export_pdf` — Export via LibreOffice headless
- `xls_export_csv` — Sheet → CSV with encoding control
- `xls_export_json` — Sheet/range → structured JSON

**Research findings**:
- LibreOffice: `soffice --headless --convert-to pdf:calc_pdf_Export --outdir <dir> <file>`
- PDF export issues: formulas may produce zeroes (need pre-calc), font metrics affect layout, checkboxes may not render
- CSV: Standard Python csv module, handle encoding (utf-8, latin-1, etc.), delimiter control
- JSON: Extract to list of dicts (header → keys), handle types (dates, numbers, strings)

**Key constraints**:
- PDF requires LibreOffice installed (external dependency)
- PDF is visual export — loses formulas, keeps values
- CSV loses formatting, keeps data only
- JSON is structured data — ideal for API integration
- All three are read-only operations (no tokens needed)

## Phase 2: PLAN — Phase 12 Detailed Execution

### Research-Validated Implementation

**1. xls_export_pdf** (LibreOffice headless)
```python
import subprocess
import shutil

def export_pdf(input_path, output_path, timeout=120):
    # Verify LibreOffice available
    soffice = shutil.which('soffice') or shutil.which('soffice.bin')
    if not soffice:
        raise RuntimeError("LibreOffice not found")

    # Build command
    cmd = [
        soffice,
        '--headless',
        '--convert-to', 'pdf:calc_pdf_Export',
        '--outdir', str(output_path.parent),
        str(input_path)
    ]

    # Execute with timeout
    result = subprocess.run(
        cmd,
        capture_output=True,
        timeout=timeout,
        check=False
    )

    # LibreOffice names output based on input
    expected_output = output_path.parent / f"{input_path.stem}.pdf"
    if expected_output!= output_path:
        expected_output.rename(output_path)

    return result.returncode == 0
```

**Key considerations from research**:
- Formulas must be pre-calculated (use Phase 8 recalc first)
- Font issues: Install Microsoft fonts or use Liberation fonts
- Page breaks: Controlled by print areas in Excel
- Timeout essential (LibreOffice can hang)

**2. xls_export_csv**
```python
import csv
from openpyxl import load_workbook

def export_csv(input_path, output_path, sheet_name=None,
               encoding='utf-8', delimiter=',', include_headers=True):
    wb = load_workbook(input_path, data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    with open(output_path, 'w', encoding=encoding, newline='') as f:
        writer = csv.writer(f, delimiter=delimiter)

        for row in ws.iter_rows(values_only=True):
            # Convert None to empty string, handle types
            cleaned = ['' if v is None else v for v in row]
            writer.writerow(cleaned)
```

**Features**:
- Encoding control (utf-8, latin-1, cp1252, etc.)
- Delimiter control (comma, semicolon, tab, pipe)
- Sheet selection
- Handles all cell types via values_only=True

**3. xls_export_json**
```python
import json
from openpyxl import load_workbook
from datetime import datetime, date

def export_json(input_path, output_path, sheet_name=None,
                range_str=None, orient='records'):
    wb = load_workbook(input_path, data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    # Determine range
    if range_str:
        cells = ws[range_str]
    else:
        cells = ws.iter_rows()

    # Extract data
    rows = list(cells)
    if not rows:
        return []

    # First row as headers
    headers = [str(cell.value) if cell.value else f"col_{i}"
               for i, cell in enumerate(rows[0])]

    data = []
    for row in rows[1:]:
        row_dict = {}
        for header, cell in zip(headers, row):
            value = cell.value
            # Convert types for JSON serialization
            if isinstance(value, (datetime, date)):
                value = value.isoformat()
            elif value is None:
                value = None
            row_dict[header] = value
        data.append(row_dict)

    # Write JSON
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    return data
```

**Features**:
- Orient options: 'records' (list of dicts), 'values' (list of lists), 'columns' (dict of lists)
- Type handling: dates → ISO strings, None → null, numbers preserved
- Range selection or full sheet
- Header detection

### File Structure

| # | Path | Purpose | Lines |
|---|---|---|---|
| 1 | `src/excel_agent/tools/export/__init__.py` | Package init | 5 |
| 2 | `src/excel_agent/tools/export/xls_export_pdf.py` | PDF export | ~150 |
| 3 | `src/excel_agent/tools/export/xls_export_csv.py` | CSV export | ~120 |
| 4 | `src/excel_agent/tools/export/xls_export_json.py` | JSON export | ~160 |
| 5 | `tests/unit/test_export_tools.py` | Unit tests | ~250 |
| 6 | `tests/integration/test_export_workflow.py` | E2E tests | ~150 |

### Detailed Specifications

**Tool 1: xls_export_pdf**
- **Inputs**: --input, --output, --timeout (default 120), --recalc (bool)
- **Validation**: Input exists, output directory writable, LibreOffice available
- **Behavior**:
  1. If --recalc: Run Phase 8 recalc first (ensure formulas calculated)
  2. Execute soffice --headless --convert-to pdf
  3. Wait for completion or timeout
  4. Move output to requested path
  5. Verify PDF created and non-empty
- **Edge cases**:
  - LibreOffice not installed → error with install instructions
  - Timeout → kill process, return error
  - Formulas not calculated → warn, suggest --recalc
  - Large file → increase timeout automatically
- **Output**: PDF path, file size, page count (if detectable), conversion time
- **Exit codes**: 0 success, 1 conversion failed, 2 LibreOffice not found, 5 timeout

**Tool 2: xls_export_csv**
- **Inputs**: --input, --output, --sheet (optional), --encoding (default utf-8), --delimiter (default,), --include-headers (bool, default true)
- **Validation**: Input exists, encoding valid, delimiter single char
- **Behavior**:
  1. Load workbook read-only, data_only=True
  2. Select sheet
  3. Iterate rows, write CSV
  4. Handle encoding errors gracefully
- **Edge cases**:
  - Empty sheet → create empty CSV with headers only
  - Special characters → respect encoding, replace errors
  - Very large sheet → stream, don't load all in memory
- **Output**: CSV path, row count, column count, encoding used
- **Features**: Supports utf-8, latin-1, cp1252, utf-16, etc.

**Tool 3: xls_export_json**
- **Inputs**: --input, --output, --sheet (optional), --range (optional), --orient (records|values|columns, default records), --pretty (bool)
- **Validation**: Input exists, orient valid, range valid if provided
- **Behavior**:
  1. Load workbook read-only, data_only=True
  2. Extract specified range or full sheet
  3. Convert to requested orientation
  4. Handle type conversion (dates → ISO, etc.)
  5. Write JSON with optional pretty-print
- **Edge cases**:
  - No headers → generate col_0, col_1, etc.
  - Mixed types in column → preserve as-is
  - Empty cells → null in JSON
  - Dates/times → ISO 8601 strings
- **Output**: JSON path, record count, structure preview
- **Orientations**:
  - records: `[{"col1": val1, "col2": val2},...]`
  - values: `[, ,...]`
  - columns: `{"col1": , "col2": }`

### Testing Strategy

**Unit tests**:
- PDF: Mock subprocess, test command building, timeout handling, LibreOffice detection
- CSV: Test encodings, delimiters, sheet selection, empty sheets, special characters
- JSON: Test orientations, type conversion, range selection, header handling

**Integration tests**:
- Full workflow: Create workbook → export to all three formats → verify outputs
- Round-trip: Export to CSV → import to new workbook → compare data
- PDF visual check: Export with charts/tables → verify PDF created
- JSON API simulation: Export → parse JSON → use in mock API

**Performance tests**:
- Large sheet (100k rows) to CSV: <5 seconds
- Large sheet to JSON: <10 seconds
- PDF export: <30 seconds for typical workbook

### Key Design Decisions

1. **LibreOffice dependency**: PDF export requires external tool. We check availability and provide clear error with install instructions. This is acceptable for server environments.

2. **Pre-calc for PDF**: Research shows formulas may export as zeroes. We offer --recalc flag to run Phase 8 recalc first. Document this requirement.

3. **Streaming for large files**: CSV and JSON use read_only mode and iterate rows to avoid loading entire workbook into memory. Essential for large datasets.

4. **Encoding flexibility**: CSV supports multiple encodings for international use. Default utf-8, but allow latin-1, cp1252 for legacy systems.

5. **JSON type handling**: Convert dates to ISO strings (JSON has no date type), preserve numbers, convert None to null. This ensures valid JSON.

6. **No tokens**: All exports are read-only operations. No data modification, no governance required.

7. **Error handling**: PDF export can fail for many reasons (LibreOffice crash, corrupt file, etc.). We capture stderr, provide meaningful errors, and suggest solutions.

### Success Criteria

| Criterion | Validation |
|-----------|------------|
| PDF export works with LibreOffice | Integration test with real LO |
| PDF handles formulas correctly | Test with --recalc flag |
| CSV supports multiple encodings | Test utf-8, latin-1, cp1252 |
| CSV handles delimiters | Test comma, semicolon, tab |
| JSON supports 3 orientations | Test records, values, columns |
| JSON type conversion works | Test dates, numbers, nulls |
| Large files stream efficiently | Test with 100k rows |
| All operations logged | Check audit entries |
| No tokens required | Verify no TokenManager calls |

### Risk Assessment

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| LibreOffice not installed | High | High | Clear error, install instructions |
| PDF export timeout | Medium | Medium | Configurable timeout, kill process |
| Formulas export as zero | Medium | Medium | Document --recalc requirement |
| Encoding errors in CSV | Low | Low | Error handling, replace invalid chars |
| Large JSON memory usage | Low | Medium | Stream write, don't build huge string |
| Date format inconsistencies | Medium | Low | Use ISO 8601, document behavior |

### Effort Breakdown (2 days)

**Day 1**: CSV and JSON (pure Python)
- xls_export_csv implementation + tests (3 hours)
- xls_export_json implementation + tests (4 hours)
- Edge case testing (1 hour)

**Day 2**: PDF (external dependency)
- xls_export_pdf implementation + tests (4 hours)
- LibreOffice integration testing (2 hours)
- Documentation and examples (2 hours)

## Phase 3: VALIDATE — Phase 12 Plan Summary

**Phase 12 delivers 3 export tools** completing the interoperability layer:

1. **xls_export_pdf** — Via LibreOffice headless, with pre-calc support, timeout handling
2. **xls_export_csv** — With encoding and delimiter control, streaming for large files
3. **xls_export_json** — Structured data export with 3 orientations, type handling

**Key characteristics**:
- All read-only → no tokens required
- PDF requires external dependency (LibreOffice)
- CSV/JSON pure Python, no external deps
- Streaming design for large datasets
- Comprehensive error handling

**Integration points**:
- Phase 8: PDF export benefits from pre-recalc
- Phase 1-11: Export any workbook created by previous tools
- Phase 13: Export examples in documentation

**Deliverables**: 6 files (~840 lines), 2 days effort, completes tool suite (53 total tools)

**Critical path**: PDF export is the only tool with external dependency. All others are self-contained. This aligns with Master Plan's emphasis on headless operation with optional LibreOffice for full fidelity.

Phase 12 completes the 53-tool suite. After this, only documentation and hardening remain (Phases 13-14).
