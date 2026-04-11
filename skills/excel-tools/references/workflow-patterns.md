# Excel Tools Workflow Patterns

Common patterns for working with excel-agent-tools. Updated for Phase 1 Remediation (April 11, 2026).

---

## Phase 1: EditSession Pattern (NEW)

**Purpose**: Unified abstraction for safe workbook manipulation

**Benefits**:
- Automatic save on exit (no double-save bug)
- Consistent macro preservation
- Version hash capture before exit
- Copy-on-write support

**Python Usage**:
```python
from excel_agent.core.edit_session import EditSession

session = EditSession.prepare(input_path, output_path)
with session:
    wb = session.workbook
    # Perform mutations
    wb["Sheet1"]["A1"] = "New Value"
    # Capture hash before exit
    version_hash = session.version_hash
# EditSession automatically saves ONCE
```

---

## Pattern 1: Clone-Modify-Export Pipeline

**Use Case**: Make changes to an existing workbook and export results.

```bash
# 1. Clone source
CLONE=$(xls-clone-workbook --input financials.xlsx --output-dir ./work/ | jq -r '.data.clone_path')

# 2. Read current state
xls-get-workbook-metadata --input "$CLONE"

# 3. Modify
xls-write-range --input "$CLONE" --output "$CLONE" --range A1 \
  --data '[["Q1", "Q2"], [100, 200]]'

# 4. Recalculate
xls-recalculate --input "$CLONE" --output "$CLONE"

# 5. Validate
xls-validate-workbook --input "$CLONE"

# 6. Export
xls-export-pdf --input "$CLONE" --outfile ./output/report.pdf
```

**Expected JSON Output**:
```json
{
  "status": "success",
  "data": {"output": "./output/report.pdf", "pages": 3}
}
```

---

## Pattern 2: Template Population

**Use Case**: Fill placeholders in template.

```bash
# Create from template with substitution
xls-create-from-template --template invoice.xltx --output invoice_001.xlsx \
  --vars '{"company": "Acme", "amount": "$500"}'
```

---

## Pattern 3: Safe Structural Edit

**Use Case**: Delete sheet that has formula references.

```bash
# 1. Check dependencies
xls-dependency-report --input workbook.xlsx | jq '.data.graph'

# 2. Generate token (REQUIRED: Set EXCEL_AGENT_SECRET first)
export EXCEL_AGENT_SECRET="your-256-bit-secret"
TOKEN=$(xls-approve-token --scope sheet:delete --file workbook.xlsx | jq -r '.data.token')

# 3. Attempt deletion (may be denied)
xls-delete-sheet --input workbook.xlsx --output workbook.xlsx \
  --name "OldData" --token "$TOKEN"

# If denied, response contains guidance:
# {
#   "status": "denied",
#   "guidance": "Run xls-update-references --updates '[{\"old\": \"...\", \"new\": \"...\"}]'"
# }

# 4. Fix references per guidance
xls-update-references --input workbook.xlsx --output workbook.xlsx \
  --updates '[{"old": "OldData!A1", "new": "NewData!A1"}]'

# 5. Retry deletion with acknowledgment
xls-delete-sheet --input workbook.xlsx --output workbook.xlsx \
  --name "OldData" --token "$TOKEN" --acknowledge-impact
```

**Phase 1 Note**: Token now requires EXCEL_AGENT_SECRET environment variable

---

## Pattern 4: Batch Processing

**Use Case**: Process multiple files.

```bash
for file in ./data/*.xlsx; do
  # Clone
  clone=$(xls-clone-workbook --input "$file" --output-dir ./work/ | jq -r '.data.clone_path')
  
  # Process
  xls-recalculate --input "$clone" --output "$clone"
  
  # Export
  xls-export-csv --input "$clone" --outfile "./output/$(basename "$file" .xlsx).csv"
done
```

---

## Pattern 5: Large Dataset Streaming

**Use Case**: Read >100k rows efficiently.

```bash
# Chunked mode returns JSONL (one JSON per line)
xls-read-range --input large.xlsx --range A1:E100000 --chunked > output.jsonl

# Parse each chunk
while IFS= read -r line; do
  chunk=$(echo "$line" | jq '.')
  # Process chunk
done < output.jsonl
```

---

## Pattern 6: Macro Safety Audit

**Use Case**: Scan file for unsafe macros before processing.

```bash
# Check if macros exist
xls-has-macros --input report.xlsm | jq '.data.has_macros'

# If true, inspect
xls-inspect-macros --input report.xlsm

# Validate safety
SAFETY=$(xls-validate-macro-safety --input report.xlsm | jq -r '.data.risk_level')

# If high/critical risk, remove macros before processing
if [ "$SAFETY" = "high" ] || [ "$SAFETY" = "critical" ]; then
  # Set secret for tokens
  export EXCEL_AGENT_SECRET="your-secret"
  TOKEN1=$(xls-approve-token --scope macro:remove --file report.xlsm | jq -r '.data.token')
  TOKEN2=$(xls-approve-token --scope macro:remove --file report.xlsm | jq -r '.data.token')
  xls-remove-macros --input report.xlsm --output report_clean.xlsx \
    --token "$TOKEN1" --token "$TOKEN2"
fi
```

---

## Pattern 7: Formula Error Detection

**Use Case**: Find and fix broken references.

```bash
# Detect errors
ERRORS=$(xls-detect-errors --input workbook.xlsx | jq '.data.errors')

# If errors found, fix by updating references
if [ $(echo "$ERRORS" | jq 'length') -gt 0 ]; then
  xls-update-references --input workbook.xlsx --output workbook.xlsx \
    --updates '[{"old": "Sheet1!#REF!", "new": "Sheet1!A1"}]'
fi
```

---

## Pattern 8: Conditional Formatting

**Use Case**: Add visual indicators.

```bash
# Color scale (green-yellow-red)
xls-apply-conditional-formatting --input report.xlsx --range A1:A100 \
  --type colorscale --colors '["00FF00", "FFFF00", "FF0000"]'

# Data bars
xls-apply-conditional-formatting --input report.xlsx --range B1:B100 \
  --type databar --color "638EC6"
```

---

## Pattern 9: Realistic Office Workflow (Phase 16)

**Use Case**: Process expense report with structured references and named ranges.

```bash
# 1. Clone realistic office workbook
xls-clone-workbook --input OfficeOps_Expenses_KPI.xlsx --output-dir ./work/

# 2. Read expense data with structured references
xls-read-range --input ./work/OfficeOps_*.xlsx --sheet Raw_Expenses --range A1:J201

# 3. Get named ranges (Categories, Departments, TaxRate)
xls-get-defined-names --input ./work/OfficeOps_*.xlsx

# 4. Set formula for FX calculation
xls-set-formula --input ./work/OfficeOps_*.xlsx --sheet Raw_Expenses --cell G2 \
  --formula '=IF(F2="USD",1,XLOOKUP(F2,FXRates!A:A,FXRates!B:B))'

# 5. Copy formula down (preferred API)
xls-copy-formula-down --input ./work/OfficeOps_*.xlsx --sheet Raw_Expenses \
  --source G2 --target G2:G201

# 6. Recalculate
xls-recalculate --input ./work/OfficeOps_*.xlsx --output ./work/calculated.xlsx

# 7. Export to CSV (note: exports full sheet)
xls-export-csv --input ./work/calculated.xlsx --sheet Raw_Expenses --outfile expenses.csv

# 8. Validate
xls-validate-workbook --input ./work/calculated.xlsx
```

**Key Phase 16 Insights**:
- Named ranges return empty list (not error) when workbook has none
- Export tools don't support `--range` - export full sheet only
- Use `--source/--target` for `xls-copy-formula-down` (preferred API)

---

## Pattern 10: Phase 1 EditSession Workflow

**Use Case**: Python script using EditSession for mutations.

```python
from excel_agent.core.edit_session import EditSession
from excel_agent.core.version_hash import compute_file_hash
import json
import subprocess

def run_tool_with_session(tool: str, input_path: str, output_path: str, **kwargs) -> dict:
    """Run tool with EditSession pattern."""
    
    # Prepare session
    session = EditSession.prepare(input_path, output_path)
    
    with session:
        # The session.workbook is available for mutations
        # But for CLI tools, we just run the tool
        pass
    
    # Build command
    cmd = [f"xls-{tool}", "--input", input_path, "--output", output_path]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        try:
            error_data = json.loads(result.stdout)
            raise RuntimeError(f"Tool failed: {error_data.get('error')}")
        except json.JSONDecodeError:
            raise RuntimeError(f"Tool failed: {result.stdout}")
    
    return json.loads(result.stdout)

# Usage example
result = run_tool_with_session(
    "write-range",
    "./work/input.xlsx",
    "./work/output.xlsx",
    range="A1",
    data='[["Header", "Value"], ["A", 100]]',
    sheet="Data"
)
print(f"Wrote range: {result['data']['range_written']}")
print(f"Version hash: {result['workbook_version']}")
```

---

## Phase 1: Multi-Tool Workflow with Shared Secret

**Use Case**: Multiple tools using same token across invocations.

```python
import subprocess
import json
import os

# REQUIRED: Set EXCEL_AGENT_SECRET before any token operations
os.environ["EXCEL_AGENT_SECRET"] = "your-256-bit-secret"

def run_tool(tool: str, **kwargs) -> dict:
    """Run tool with Phase 1 error handling."""
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    # Phase 1: Check returncode BEFORE json.loads
    if result.returncode != 0:
        try:
            # Tools write errors to stdout
            error_data = json.loads(result.stdout)
            raise RuntimeError(f"Tool failed: {error_data.get('error', 'unknown')}")
        except json.JSONDecodeError:
            raise RuntimeError(f"Tool failed: {result.stdout or result.stderr}")
    
    return json.loads(result.stdout)

# Multi-tool workflow with shared token
work_path = "./work/data.xlsx"

# 1. Clone
clone_result = run_tool("clone-workbook", input="original.xlsx", output_dir="./work/")
clone_path = clone_result["data"]["clone_path"]

# 2. Write data
run_tool("write-range", input=clone_path, output=clone_path, range="A1", 
         data='[["New", "Data"]]')

# 3. Generate token for deletion
token_result = run_tool("approve-token", scope="sheet:delete", file=clone_path, ttl=300)
token = token_result["data"]["token"]

# 4. Delete sheet (token works because EXCEL_AGENT_SECRET is set)
delete_result = run_tool("delete-sheet", input=clone_path, output=clone_path, 
                        name="OldSheet", token=token)

print(f"Deletion status: {delete_result['status']}")
print(f"Version: {delete_result['workbook_version']}")
```

---

## Phase 1: Cross-Sheet Reference Preservation

**Use Case**: Recalculate workbook with cross-sheet references.

```python
# Before Phase 1: Cross-sheet refs might break after recalculate
# After Phase 1: Sheet casing preserved automatically

result = subprocess.run(
    ["xls-recalculate", "--input", "workbook.xlsx", "--output", "calculated.xlsx"],
    capture_output=True, text=True
)

if result.returncode == 0:
    data = json.loads(result.stdout)
    print(f"Formulas calculated: {data['data']['calculated_count']}")
    print(f"Engine used: {data['data']['engine']}")  # tier1_formulas or tier2_libreoffice
    # Cross-sheet references now preserved correctly
```

---

## Error Handling Pattern (Phase 1 Updated)

**Always check exit codes before parsing JSON**:

```python
import subprocess
import json

result = subprocess.run(
    ["xls-read-range", "--input", "data.xlsx", "--range", "A1"],
    capture_output=True,
    text=True
)

if result.returncode == 0:
    data = json.loads(result.stdout)
    values = data["data"]["values"]
elif result.returncode == 1:
    # Phase 1: Could be validation or impact denial
    data = json.loads(result.stdout)
    if data.get("status") == "denied":
        print(f"Permission denied: {data.get('guidance')}")
    else:
        print("Validation error - check input")
elif result.returncode == 2:
    print("File not found")
elif result.returncode == 3:
    print("File locked - retry with backoff")
elif result.returncode == 4:
    # Phase 1: Now returns "denied" status
    data = json.loads(result.stdout)
    print(f"Permission denied: {data.get('guidance')}")
    print("Generate new token with correct scope")
elif result.returncode == 5:
    print("Internal error - check traceback")
```

---

## Python Integration with Realistic Error Handling (Phase 1)

```python
import subprocess
import json
from pathlib import Path

# REQUIRED: Set for Phase 1 token compatibility
import os
os.environ["EXCEL_AGENT_SECRET"] = "your-secret"

def run_tool(tool: str, **kwargs) -> dict:
    """Run an excel-agent tool with Phase 1 error handling."""
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    # Phase 1: Check returncode BEFORE parsing JSON
    if result.returncode != 0:
        try:
            # Tools write errors to stdout
            error_data = json.loads(result.stdout)
            raise RuntimeError(f"Tool failed: {error_data.get('error', 'unknown error')}")
        except json.JSONDecodeError:
            raise RuntimeError(f"Tool failed: {result.stdout or result.stderr}")
    
    return json.loads(result.stdout)

# Realistic office workflow
try:
    # Clone
    result = run_tool("clone-workbook", input="OfficeOps_Expenses_KPI.xlsx", output_dir="./work/")
    clone_path = result["data"]["clone_path"]
    
    # Get named ranges
    result = run_tool("get-defined-names", input=clone_path)
    named_ranges = result["data"]["named_ranges"]
    print(f"Found {len(named_ranges)} named ranges")
    
    # Read expense data
    result = run_tool("read-range", input=clone_path, sheet="Raw_Expenses", range="A1:J10")
    values = result["data"]["values"]
    
    # Process...
    
except RuntimeError as e:
    print(f"Workflow failed: {e}")
    # Handle specific error based on message
```

---

## Error Handling Patterns (Phase 1)

### Pattern 1: Check Return Code First
```python
result = subprocess.run(cmd, capture_output=True, text=True)

# Phase 1 lesson: Check returncode BEFORE json.loads
if result.returncode != 0:
    # Parse error from stdout (not stderr)
    error = json.loads(result.stdout)
    handle_error(error)
else:
    data = json.loads(result.stdout)
    process_data(data)
```

### Pattern 2: Handle "denied" Status (Phase 1)
```python
if result.returncode == 4 or result.returncode == 1:
    data = json.loads(result.stdout)
    if data.get("status") == "denied":
        guidance = data.get("guidance")
        impact = data.get("impact", {})
        print(f"Denied: {guidance}")
        print(f"Impact: {impact}")
        # Run remediation per guidance
```

### Pattern 3: Structured Reference Handling
```python
# Check if structured references exist
result = run_tool("get-defined-names", input=workbook)
named_ranges = result["data"]["named_ranges"]

# Look for table references
for nr in named_ranges:
    if "[" in nr["refers_to"]:  # Structured reference
        print(f"Table found: {nr['name']} -> {nr['refers_to']}")
```

### Pattern 4: Dual API Support
```python
# Try preferred API first, fallback to legacy
try:
    result = run_tool("copy-formula-down", input=wb, source="A1", target="A1:A10")
except RuntimeError:
    # Fallback to legacy API
    result = run_tool("copy-formula-down", input=wb, cell="A1", count=9)
```

---

## Original Python Integration Pattern

```python
import subprocess
import json
from pathlib import Path

def run_tool(tool: str, **kwargs) -> dict:
    """Run an excel-agent tool and return parsed JSON."""
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        # Parse error from stdout (excel-agent-tools writes JSON errors to stdout)
        try:
            error_data = json.loads(result.stdout)
            raise RuntimeError(f"Tool failed: {error_data.get('error', 'unknown')}")
        except json.JSONDecodeError:
            raise RuntimeError(f"Tool failed: {result.stdout or result.stderr}")
    
    data = json.loads(result.stdout)
    
    return data

# Usage
clone = run_tool("clone-workbook", input="data.xlsx", output_dir="./work/")
clone_path = clone["data"]["clone_path"]

meta = run_tool("get-workbook-metadata", input=clone_path)
print(f"Sheets: {meta['data']['sheet_count']}")
```

---

**Document Version**: Phase 1 Remediation (April 11, 2026)
**Phase 1 Status**: ✅ All critical issues resolved, 554/554 tests passing
