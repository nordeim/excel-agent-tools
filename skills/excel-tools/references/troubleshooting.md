# Excel Tools Troubleshooting Guide

Common issues and solutions. Updated for Phase 1 Remediation (April 11, 2026).

---

## Phase 1: Critical Issues Fixed

### Issue: Double-Save Bug (RESOLVED)

**Symptoms**: Race conditions, inconsistent saves

**Cause**: ExcelAgent saves on exit + explicit `wb.save()` = double write

**Fix**: Use EditSession abstraction (Phase 1)

```python
# BEFORE (buggy):
with ExcelAgent(path, mode="rw") as agent:
    wb = agent.workbook
    # ... mutations ...
    wb.save(str(output_path))  # Explicit save
# ExcelAgent.__exit__ also saves → DOUBLE SAVE!

# AFTER (fixed with EditSession):
session = EditSession.prepare(input_path, output_path)
with session:
    wb = session.workbook
    # ... mutations ...
    version_hash = session.version_hash
# EditSession handles save automatically ONCE
```

---

### Issue: Token Validation Fails Across Tool Invocations (RESOLVED)

**Symptoms**: Exit code 4, "Invalid token signature" when using same token across tools

**Cause**: Each tool created new ApprovalTokenManager with random secret

**Fix**: Set EXCEL_AGENT_SECRET environment variable (Phase 1)

```bash
# Set before any token operations
export EXCEL_AGENT_SECRET="your-256-bit-secret"

# Now tokens work across tool invocations
TOKEN=$(xls-approve-token --scope sheet:delete --file workbook.xlsx | jq -r '.data.token')
xls-delete-sheet --input workbook.xlsx --name "Old" --token "$TOKEN"
```

---

### Issue: Sheet Names Uppercased After Recalculate (RESOLVED)

**Symptoms**: Cross-sheet references broken after `xls-recalculate`

**Cause**: `formulas` library uppercases all sheet names

**Fix**: Tier1Calculator now restores original casing (Phase 1)

```python
# After recalculation, sheet names are automatically restored
# Cross-sheet references now work correctly
```

---

### Issue: Dependency Report Shows "Safe" for Full Sheet Deletion (RESOLVED)

**Symptoms**: Deleting sheet with references shows no broken refs

**Cause**: Large ranges not expanded properly

**Fix**: DependencyTracker now expands large ranges by iterating forward graph (Phase 1)

---

### Issue: Audit Log Not Written (RESOLVED)

**Symptoms**: No audit trail entries for structural operations

**Cause**: Tools called `audit.log_operation()` but method is `log()`

**Fix**: Updated all structural tools to use correct API (Phase 1)

---

### Issue: Copy Formula Down Off-By-One (RESOLVED)

**Symptoms**: Wrong number of cells filled

**Cause**: Target count included source cell

**Fix**: Fixed count calculation (Phase 1)

---

### Issue: PermissionDenied Returns Status "error" (RESOLVED)

**Symptoms**: SDK can't distinguish permission vs validation errors

**Cause**: Tool base returned "error" for all exceptions

**Fix**: Now returns "denied" for exit code 4 (Phase 1)

---

## Original Issues & Solutions

### Issue: File Lock Not Released

**Symptoms**: Subsequent operations fail with "File is locked"

**Cause**: Exception in context body before `__exit__`

**Solution**: 
- FileLock releases lock in `__exit__` even on exception
- If lock persists, manually delete `.{filename}.lock` file
- Implement exponential backoff retry:

```bash
for i in 0.5 1 2 4; do
  xls-read-range --input file.xlsx --range A1 && break
  sleep $i
done
```

---

### Issue: #REF! Errors After Structural Changes

**Symptoms**: Cells show `#REF!` after deleting rows/columns/sheets

**Cause**: Deleted cells were referenced by formulas

**Prevention**:
1. Run dependency report before destructive ops:
```bash
xls-dependency-report --input file.xlsx | jq '.data.graph'
```

2. Check impact before deleting:
```bash
xls-delete-sheet --input file.xlsx --name "Old" --token T
# If denied, response contains guidance with affected references
```

**Fix**:
1. Identify broken references:
```bash
xls-detect-errors --input file.xlsx | jq '.data.errors'
```

2. Update references:
```bash
xls-update-references --input file.xlsx --output file.xlsx \
  --updates '[{"old": "Sheet1!#REF!", "new": "Sheet1!A1"}]'
```

---

### Issue: Token Validation Fails (Original)

**Symptoms**: Exit code 4, "Permission denied"

**Causes**:
1. Token expired (default TTL 300s)
2. Wrong scope
3. File hash mismatch (file changed after token generation)
4. Token already used (nonce replay)
5. **Phase 1**: EXCEL_AGENT_SECRET not set

**Solution**:
1. Generate new token with correct scope:
```bash
export EXCEL_AGENT_SECRET="your-256-bit-secret"
xls-approve-token --scope sheet:delete --file workbook.xlsx --ttl 600
```

2. Ensure file hasn't changed between token generation and use

---

### Issue: Chunked Read Returns Unexpected Format

**Symptoms**: JSON parsing fails with "Extra data"

**Cause**: `--chunked` returns JSONL, not single JSON

**Solution**:
```bash
# Parse as JSONL (one JSON per line)
xls-read-range --input large.xlsx --range A1:E100000 --chunked > output.jsonl

# Parse each chunk
while IFS= read -r line; do
  chunk=$(echo "$line" | jq '.')
  # Process chunk
done < output.jsonl
```

---

### Issue: LibreOffice Not Found

**Symptoms**: PDF export or Tier 2 calculation fails

**Solution**:
```bash
# Ubuntu/Debian
sudo apt-get install -y libreoffice-calc

# macOS
brew install --cask libreoffice

# Verify
soffice --headless --version
```

---

### Issue: Export Argparse Conflict

**Symptoms**: `xls-export-pdf: error: argument --output: not allowed`

**Cause**: Export tools use `--outfile` not `--output`

**Wrong**:
```bash
xls-export-pdf --input file.xlsx --output file.pdf
```

**Correct**:
```bash
xls-export-pdf --input file.xlsx --outfile file.pdf
```

---

### Issue: Macro Safety Scan Fails

**Symptoms**: `xls-validate-macro-safety` errors on .xlsx files

**Cause**: Only .xlsm files contain VBA

**Solution**:
```bash
# Check file type first
xls-has-macros --input file.xlsm

# If false, skip macro safety scan
```

---

### Issue: Formula Not Calculating

**Symptoms**: Cells show formulas as text, not values

**Causes**:
1. Formula written as string, not formula type
2. Recalculation not performed

**Solution**:
1. Write with `--type formula`:
```bash
xls-write-cell --input file.xlsx --cell A1 --value "=SUM(B1:B10)" --type formula
```

2. Recalculate:
```bash
xls-recalculate --input file.xlsx --output file.xlsx
```

**Note**: Tier 1 calculator now preserves sheet casing after recalculation (Phase 1 fix)

---

### Issue: Date Format Incorrect

**Symptoms**: Dates appear as numbers (e.g., 45000)

**Cause**: Excel stores dates as serial numbers

**Solution**:
1. When writing, use ISO 8601 format:
```bash
xls-write-cell --input file.xlsx --cell A1 --value "2026-04-11" --type date
```

2. Set number format:
```bash
xls-set-number-format --input file.xlsx --range A1 --format "YYYY-MM-DD"
```

---

### Issue: Large File Performance

**Symptoms**: Timeout or memory errors with >100k rows

**Solution**:
1. Use chunked mode:
```bash
xls-read-range --input large.xlsx --range A1:E100000 --chunked
```

2. Process in batches:
```bash
for i in {0..9}; do
  start=$((i*10000+1))
  end=$((start+9999))
  xls-read-range --input large.xlsx --range "A${start}:E${end}"
done
```

---

### Issue: Path Not Found

**Symptoms**: Exit code 2, "File does not exist"

**Solutions**:
1. Check path exists:
```bash
test -f workbook.xlsx && echo "exists"
```

2. Use absolute paths:
```bash
xls-read-range --input "$(pwd)/workbook.xlsx" --range A1
```

3. Check permissions:
```bash
ls -la workbook.xlsx
```

---

### Issue: Validation Errors

**Symptoms**: Exit code 1, schema validation failed

**Common Causes**:
1. Malformed JSON in `--data` or `--updates`
2. Invalid range format
3. Missing required argument
4. **Phase 1**: Wrong API method for structural tools

**Solutions**:
1. Validate JSON:
```bash
echo '["data"]' | jq '.' # Check valid JSON
```

2. Check range format (A1 notation):
```bash
# Valid: A1, A1:C10, Sheet1!A1:C10
# Invalid: a1, A-1, A1..C10
```

3. Check required args with `--help`:
```bash
xls-write-range --help
```

---

### Issue: Circular Reference Errors

**Symptoms**: `#VALUE!` or calculation hangs

**Detection**:
```bash
xls-dependency-report --input file.xlsx | jq '.data.circular_refs'
```

**Fix**:
1. Break the circular chain
2. Use `xls-update-references` to fix formulas
3. Set one cell to a static value

---

## Debugging Tips

### Enable Verbose Output

Most tools don't have verbose mode, but you can:
1. Check JSON response for `warnings` array
2. Use `jq` to pretty-print:
```bash
xls-read-range --input file.xlsx --range A1 | jq '.'
```

### Check Tool Availability

```bash
which xls-read-range
xls-read-range --help | head -5
```

### Verify Installation

```bash
pip show excel-agent-tools
# Should show version 1.0.0
```

### Test with Minimal Input

```bash
# Create simple test file
xls-create-new --output /tmp/test.xlsx
xls-write-range --input /tmp/test.xlsx --range A1 --data '[["Test"]]'
xls-read-range --input /tmp/test.xlsx --range A1
```

### Check Python Version

```bash
python --version
# Requires >= 3.12
```

### Check EXCEL_AGENT_SECRET

```bash
# Verify secret is set
if [ -z "$EXCEL_AGENT_SECRET" ]; then
  echo "WARNING: EXCEL_AGENT_SECRET not set - tokens will fail"
fi
```

---

## Historical Fixes

### Phase 16: Realistic Test Plan - Known Issues & Fixes

#### Issue: Help Text Shows Format Error
**Symptoms**: Tool crashes on `--help` with `ValueError: unsupported format character`
**Affected**: `xls-set-number-format` (fixed in Phase 16)
**Solution**: Fixed by escaping `%` in help text. Update to latest version.

#### Issue: Duplicate --force Argument
**Symptoms**: `argparse.ArgumentError: conflicting option string: --force`
**Affected**: `xls-inject-vba-project` (fixed in Phase 16)
**Solution**: Fixed by removing duplicate argument. Update to latest version.

#### Issue: Named Range Reading Fails
**Symptoms**: `xls-get-defined-names` returns exit code 5
**Affected**: `xls-get-defined-names` (fixed in Phase 16)
**Solution**: Added null-safety. Tool now returns empty list for workbooks without named ranges.

#### Issue: Export Tools Don't Support --range
**Symptoms**: `unrecognized arguments: --range A1:C10`
**Cause**: Export tools export entire sheets, not ranges
**Workaround**: Use `xls-read-range` to extract data, then pipe to CSV.

---

## Getting Help

1. Check tool help: `xls-<tool> --help`
2. Review API docs: `docs/API.md`
3. Check realistic test examples: `tests/integration/test_realistic_office_workflow.py`
4. Check Phase 1 documentation: `CLAUDE.md`, `Project_Architecture_Document.md`
5. Check status codes: `src/excel_agent/utils/exit_codes.py`
6. Run with explicit paths and `jq` for debugging

---

**Document Version**: Phase 1 Remediation (April 11, 2026)
**Phase 1 Status**: All critical issues resolved, 554/554 tests passing
