# Workflow Recipes: excel-agent-tools

**Version:** 1.0.0  
**Target Audience:** AI Prompt Engineers & Orchestration Frameworks

This document provides production-ready workflow recipes with complete JSON I/O examples for AI few-shot training.

---

## Table of Contents

1. [Recipe 1: Financial Reporting Pipeline](#recipe-1-financial-reporting-pipeline)
2. [Recipe 2: Safe Structural Edit](#recipe-2-safe-structural-edit)
3. [Recipe 3: Template Population](#recipe-3-template-population)
4. [Recipe 4: Macro Security Audit](#recipe-4-macro-security-audit)
5. [Recipe 5: Large Dataset Migration](#recipe-5-large-dataset-migration)

---

## Recipe 1: Financial Reporting Pipeline

**Purpose:** Complete financial report generation with data injection, calculation, and PDF export.

**Duration:** ~10-15 seconds  
**Risk Level:** Low (non-destructive operations)  
**Token Requirements:** None

### Workflow Steps

```
Clone → Write → Recalculate → Validate → Export PDF
```

### Complete JSON I/O

#### Step 1: Clone Workbook

**Request:**
```bash
xls-clone-workbook --input financial_template.xlsx --output-dir ./work/
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "clone_path": "/work/financial_template_20260409T143022_a3f7e2d1.xlsx",
    "source_hash": "sha256:8f4a3b...",
    "clone_hash": "sha256:8f4a3b...",
    "timestamp": "20260409T143022"
  },
  "timestamp": "2026-04-09T14:30:22Z"
}
```

**Agent Action:** Extract `clone_path` for subsequent steps.

---

#### Step 2: Write Financial Data

**Request:**
```bash
xls-write-range --input /work/financial_template_20260409T143022_a3f7e2d1.xlsx \
  --output /work/financial_template_20260409T143022_a3f7e2d1.xlsx \
  --range A2 --sheet "Data" \
  --data '[["Q1 2026", 150000, 75000, 75000], ["Q2 2026", 175000, 87500, 87500]]'
```

**Response:**
```json
{
  "status": "success",
  "data": {"range_written": "A2:D3"},
  "impact": {"cells_modified": 6, "formulas_updated": 0},
  "timestamp": "2026-04-09T14:30:23Z"
}
```

**Agent Action:** Verify `cells_modified` matches expected (6 cells = 2 rows × 3 cols).

---

#### Step 3: Recalculate Formulas

**Request:**
```bash
xls-recalculate --input /work/financial_template_20260409T143022_a3f7e2d1.xlsx \
  --output /work/financial_template_20260409T143022_a3f7e2d1.xlsx
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "formula_count": 8,
    "calculated_count": 8,
    "error_count": 0,
    "engine": "tier1_formulas",
    "recalc_time_ms": 45.2
  },
  "timestamp": "2026-04-09T14:30:24Z"
}
```

**Agent Action:** Verify `error_count` is 0. Note `engine` for optimization insights.

---

#### Step 4: Validate Workbook

**Request:**
```bash
xls-validate-workbook --input /work/financial_template_20260409T143022_a3f7e2d1.xlsx
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "valid": true,
    "errors": [],
    "warnings": [],
    "circular_refs": [],
    "broken_references": 0
  },
  "timestamp": "2026-04-09T14:30:25Z"
}
```

**Agent Action:** Verify `valid: true` before export. Address any warnings.

---

#### Step 5: Export to PDF

**Request:**
```bash
xls-export-pdf --input /work/financial_template_20260409T143022_a3f7e2d1.xlsx \
  --outfile /work/financial_report_2026.pdf --recalc
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "output": "/work/financial_report_2026.pdf",
    "pages": 3,
    "file_size_bytes": 45872
  },
  "timestamp": "2026-04-09T14:30:28Z"
}
```

**Agent Action:** Report final output path to user.

---

### Complete Python Implementation

```python
import json
import subprocess

def financial_report_pipeline(template_path: str, data: list, output_dir: str) -> dict:
    """
    Generate financial report with data injection and PDF export.
    
    Args:
        template_path: Path to financial template .xlsx
        data: 2D array of financial data
        output_dir: Working directory for clones
    
    Returns:
        Result dict with output paths
    """
    # Step 1: Clone
    clone_result = subprocess.run(
        ["xls-clone-workbook", "--input", template_path, "--output-dir", output_dir],
        capture_output=True, text=True
    )
    clone_data = json.loads(clone_result.stdout)
    clone_path = clone_data["data"]["clone_path"]
    
    # Step 2: Write data
    data_json = json.dumps(data)
    subprocess.run(
        ["xls-write-range", "--input", clone_path, "--output", clone_path,
         "--range", "A2", "--sheet", "Data", "--data", data_json],
        capture_output=True, text=True
    )
    
    # Step 3: Recalculate
    subprocess.run(
        ["xls-recalculate", "--input", clone_path, "--output", clone_path],
        capture_output=True, text=True
    )
    
    # Step 4: Validate
    validate_result = subprocess.run(
        ["xls-validate-workbook", "--input", clone_path],
        capture_output=True, text=True
    )
    validate_data = json.loads(validate_result.stdout)
    assert validate_data["data"]["valid"], "Validation failed"
    
    # Step 5: Export PDF
    pdf_path = f"{output_dir}/financial_report.pdf"
    export_result = subprocess.run(
        ["xls-export-pdf", "--input", clone_path, "--outfile", pdf_path, "--recalc"],
        capture_output=True, text=True
    )
    
    return {
        "clone_path": clone_path,
        "pdf_path": pdf_path,
        "status": "success"
    }

# Usage
result = financial_report_pipeline(
    template_path="financial_template.xlsx",
    data=[["Q1 2026", 150000, 75000, 75000]],
    output_dir="./work"
)
```

---

## Recipe 2: Safe Structural Edit

**Purpose:** Remove sheet after impact analysis and reference remediation.

**Duration:** ~5-10 seconds  
**Risk Level:** High (requires governance tokens)  
**Token Requirements:** `sheet:delete`

### Workflow Steps

```
Dependency Report → Attempt Delete → Fix References → Token → Delete → Validate
```

### Complete JSON I/O

#### Step 1: Get Dependency Report

**Request:**
```bash
xls-dependency-report --input financials.xlsx
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "stats": {
      "total_cells": 1250,
      "total_formulas": 47,
      "total_edges": 89
    },
    "graph": {
      "Sheet1!A1": ["Sheet2!A1", "Summary!B5"],
      "Data!B1": ["Sheet1!C5", "Sheet1!C6"]
    },
    "circular_refs": []
  },
  "timestamp": "2026-04-09T14:30:30Z"
}
```

**Agent Action:** Identify `Data` sheet has dependents on `Sheet1`.

---

#### Step 2: Attempt Delete (Expected Denial)

**Request:**
```bash
xls-delete-sheet --input financials.xlsx --output financials.xlsx \
  --name "Data" --token "invalid-or-placeholder-token"
```

**Response:**
```json
{
  "status": "denied",
  "exit_code": 1,
  "denial_reason": "Operation would break 8 formula references across 2 sheets",
  "guidance": "Run xls-update-references.py --updates '[{\"old\": \"Data!B1\", \"new\": \"Archive!B1\"}]' before retrying",
  "impact": {
    "broken_references": 8,
    "affected_sheets": ["Sheet1", "Summary"]
  },
  "timestamp": "2026-04-09T14:30:31Z"
}
```

**Agent Action:** Parse `guidance` field for remediation instructions.

---

#### Step 3: Update References

**Request:**
```bash
xls-update-references --input financials.xlsx --output financials.xlsx \
  --updates '[{"old": "Data!B1", "new": "Archive!B1"}, {"old": "Data!B2:B10", "new": "Archive!B2:B10"}]'
```

**Response:**
```json
{
  "status": "success",
  "data": {"formulas_updated": 8},
  "impact": {"formulas_updated": 8, "cells_modified": 0},
  "timestamp": "2026-04-09T14:30:32Z"
}
```

**Agent Action:** Verify `formulas_updated` matches expected from denial.

---

#### Step 4: Generate Approval Token

**Request:**
```bash
xls-approve-token --scope sheet:delete --file financials.xlsx --ttl 300
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "token": "eyJzY29wZSI6InNoZWV0OmRlbGV0ZSIs...",
    "scope": "sheet:delete",
    "expires_at": "2026-04-09T14:35:32Z",
    "file_hash": "sha256:9e2d4f..."
  },
  "timestamp": "2026-04-09T14:30:32Z"
}
```

**Agent Action:** Extract `token` for delete request.

---

#### Step 5: Delete Sheet with Acknowledgment

**Request:**
```bash
xls-delete-sheet --input financials.xlsx --output financials.xlsx \
  --name "Data" --token "eyJzY29wZSI6InNoZWV0OmRlbGV0ZSIs..." --acknowledge-impact
```

**Response:**
```json
{
  "status": "success",
  "data": {"deleted_sheet": "Data", "sheets_remaining": ["Sheet1", "Archive", "Summary"]},
  "impact": {"formulas_updated": 0, "sheets_deleted": 1},
  "timestamp": "2026-04-09T14:30:33Z"
}
```

**Agent Action:** Verify sheet removed from `sheets_remaining`.

---

#### Step 6: Final Validation

**Request:**
```bash
xls-validate-workbook --input financials.xlsx
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "valid": true,
    "errors": [],
    "warnings": [],
    "circular_refs": [],
    "broken_references": 0
  },
  "timestamp": "2026-04-09T14:30:34Z"
}
```

**Agent Action:** Confirm `broken_references: 0`.

---

### Complete Python Implementation

```python
import json
import subprocess

def safe_structural_edit(workbook_path: str, sheet_to_delete: str) -> dict:
    """
    Safely delete sheet after impact analysis and remediation.
    
    Args:
        workbook_path: Path to workbook
        sheet_to_delete: Name of sheet to remove
    
    Returns:
        Result dict with status and impact
    """
    # Step 1: Check dependencies
    dep_result = subprocess.run(
        ["xls-dependency-report", "--input", workbook_path],
        capture_output=True, text=True
    )
    
    # Step 2: Attempt delete (will be denied with guidance)
    deny_result = subprocess.run(
        ["xls-delete-sheet", "--input", workbook_path, "--output", workbook_path,
         "--name", sheet_to_delete, "--token", "placeholder"],
        capture_output=True, text=True
    )
    
    if deny_result.returncode == 1:
        deny_data = json.loads(deny_result.stdout)
        guidance = deny_data["guidance"]
        
        # Parse guidance for updates
        # (In production, use regex to extract JSON from guidance string)
        updates = extract_updates_from_guidance(guidance)
        
        # Step 3: Update references
        subprocess.run(
            ["xls-update-references", "--input", workbook_path,
             "--output", workbook_path, "--updates", json.dumps(updates)],
            capture_output=True, text=True
        )
        
        # Step 4: Get token
        token_result = subprocess.run(
            ["xls-approve-token", "--scope", "sheet:delete",
             "--file", workbook_path, "--ttl", "300"],
            capture_output=True, text=True
        )
        token_data = json.loads(token_result.stdout)
        token = token_data["data"]["token"]
        
        # Step 5: Delete with acknowledgment
        delete_result = subprocess.run(
            ["xls-delete-sheet", "--input", workbook_path, "--output", workbook_path,
             "--name", sheet_to_delete, "--token", token, "--acknowledge-impact"],
            capture_output=True, text=True
        )
        
        return json.loads(delete_result.stdout)
    
    return {"status": "error", "message": "Unexpected response"}

def extract_updates_from_guidance(guidance: str) -> list:
    """Extract updates JSON from guidance string."""
    # Extract JSON array from guidance
    import re
    match = re.search(r'--updates\s*\'([^\']+)\'', guidance)
    if match:
        return json.loads(match.group(1))
    return []
```

---

## Recipe 3: Template Population

**Purpose:** Fill template placeholders with variables, export to multiple formats.

**Duration:** ~5-8 seconds  
**Risk Level:** Low  
**Token Requirements:** None

### Workflow Steps

```
Load Template → Substitute Variables → Export CSV/JSON
```

### Complete JSON I/O

#### Step 1: Create from Template

**Request:**
```bash
xls-create-from-template --template invoice.xltx --output invoice_001.xlsx \
  --vars '{"company": "Acme Corporation", "invoice_num": "INV-2026-001", "date": "2026-04-09", "amount": "$5,000.00"}'
```

**Template Content:**
```
Invoice: {{invoice_num}}
Date: {{date}}

Bill To:
{{company}}

Amount Due: {{amount}}
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "path": "invoice_001.xlsx",
    "substitutions": 4
  },
  "timestamp": "2026-04-09T14:30:35Z"
}
```

---

#### Step 2: Export to CSV

**Request:**
```bash
xls-export-csv --input invoice_001.xlsx --outfile invoice_001.csv --sheet "Invoice"
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "output": "invoice_001.csv",
    "rows": 15,
    "encoding": "utf-8"
  },
  "timestamp": "2026-04-09T14:30:36Z"
}
```

---

#### Step 3: Export to JSON (Records Format)

**Request:**
```bash
xls-export-json --input invoice_001.xlsx --outfile invoice_001.json \
  --format records --sheet "Invoice"
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "output": "invoice_001.json",
    "records": 15,
    "format": "records"
  },
  "timestamp": "2026-04-09T14:30:37Z"
}

# Output file content:
[
  {"Field": "Invoice", "Value": "INV-2026-001"},
  {"Field": "Date", "Value": "2026-04-09"},
  {"Field": "Bill To", "Value": "Acme Corporation"},
  {"Field": "Amount Due", "Value": "$5,000.00"}
]
```

---

### Complete Python Implementation

```python
import json
import subprocess

def populate_template(template_path: str, variables: dict, output_prefix: str) -> dict:
    """
    Populate template and export to multiple formats.
    
    Args:
        template_path: Path to .xltx template
        variables: Dict of placeholder -> value
        output_prefix: Prefix for output files
    
    Returns:
        Result dict with all output paths
    """
    vars_json = json.dumps(variables)
    
    # Step 1: Create from template
    result = subprocess.run(
        ["xls-create-from-template", "--template", template_path,
         "--output", f"{output_prefix}.xlsx", "--vars", vars_json],
        capture_output=True, text=True
    )
    
    # Step 2: Export CSV
    subprocess.run(
        ["xls-export-csv", "--input", f"{output_prefix}.xlsx",
         "--outfile", f"{output_prefix}.csv"],
        capture_output=True, text=True
    )
    
    # Step 3: Export JSON
    subprocess.run(
        ["xls-export-json", "--input", f"{output_prefix}.xlsx",
         "--outfile", f"{output_prefix}.json", "--format", "records"],
        capture_output=True, text=True
    )
    
    return {
        "xlsx": f"{output_prefix}.xlsx",
        "csv": f"{output_prefix}.csv",
        "json": f"{output_prefix}.json"
    }

# Usage
result = populate_template(
    template_path="invoice.xltx",
    variables={
        "company": "Acme Corporation",
        "invoice_num": "INV-2026-001",
        "date": "2026-04-09",
        "amount": "$5,000.00"
    },
    output_prefix="invoice_001"
)
```

---

## Recipe 4: Macro Security Audit

**Purpose:** Scan workbook for VBA macros, assess risk, remove if unsafe.

**Duration:** ~3-5 seconds  
**Risk Level:** Critical (double-token required for removal)  
**Token Requirements:** `macro:remove` × 2

### Workflow Steps

```
Check Macros → Inspect → Validate Safety → Remove if unsafe
```

### Complete JSON I/O

#### Step 1: Check for Macros

**Request:**
```bash
xls-has-macros --input report.xlsm
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "has_macros": true,
    "macro_count": 3
  },
  "timestamp": "2026-04-09T14:30:38Z"
}
```

**Agent Action:** If `has_macros: false`, workflow complete.

---

#### Step 2: Inspect Macro Details

**Request:**
```bash
xls-inspect-macros --input report.xlsm
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "modules": [
      {"name": "Module1", "type": "standard", "code_size": 2048},
      {"name": "ThisWorkbook", "type": "document", "code_size": 512},
      {"name": "Sheet1", "type": "document", "code_size": 256}
    ],
    "has_signature": false
  },
  "timestamp": "2026-04-09T14:30:39Z"
}
```

---

#### Step 3: Validate Safety

**Request:**
```bash
xls-validate-macro-safety --input report.xlsm
```

**Response (High Risk):**
```json
{
  "status": "success",
  "data": {
    "risk_level": "high",
    "auto_exec_triggers": ["AutoOpen", "AutoClose"],
    "suspicious_keywords": ["Shell", "CreateObject", "WScript.Shell"],
    "iocs": [],
    "recommendation": "Remove or digitally sign before processing"
  },
  "timestamp": "2026-04-09T14:30:40Z"
}
```

**Agent Action:** If `risk_level` is `high`, recommend removal.

---

#### Step 4: Remove Macros (Double Token Required)

**Request:**
```bash
# Generate first token
xls-approve-token --scope macro:remove --file report.xlsm

# Generate second token (separate request)
xls-approve-token --scope macro:remove --file report.xlsm

# Remove with both tokens
xls-remove-macros --input report.xlsm --output report_clean.xlsx \
  --token "TOKEN1" --token "TOKEN2"
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "output": "report_clean.xlsx",
    "macros_removed": 3,
    "original_format": ".xlsm",
    "output_format": ".xlsx"
  },
  "timestamp": "2026-04-09T14:30:42Z"
}
```

---

### Complete Python Implementation

```python
import json
import subprocess

def macro_security_audit(workbook_path: str) -> dict:
    """
    Audit and optionally remove macros from workbook.
    
    Args:
        workbook_path: Path to .xlsm file
    
    Returns:
        Result dict with audit results
    """
    # Step 1: Check for macros
    check_result = subprocess.run(
        ["xls-has-macros", "--input", workbook_path],
        capture_output=True, text=True
    )
    check_data = json.loads(check_result.stdout)
    
    if not check_data["data"]["has_macros"]:
        return {"status": "clean", "message": "No macros found"}
    
    # Step 2: Inspect
    inspect_result = subprocess.run(
        ["xls-inspect-macros", "--input", workbook_path],
        capture_output=True, text=True
    )
    inspect_data = json.loads(inspect_result.stdout)
    
    # Step 3: Validate safety
    safety_result = subprocess.run(
        ["xls-validate-macro-safety", "--input", workbook_path],
        capture_output=True, text=True
    )
    safety_data = json.loads(safety_result.stdout)
    
    risk_level = safety_data["data"]["risk_level"]
    
    if risk_level in ["high", "critical"]:
        # Get double tokens for removal
        token1 = json.loads(subprocess.run(
            ["xls-approve-token", "--scope", "macro:remove", "--file", workbook_path],
            capture_output=True, text=True
        ).stdout)["data"]["token"]
        
        token2 = json.loads(subprocess.run(
            ["xls-approve-token", "--scope", "macro:remove", "--file", workbook_path],
            capture_output=True, text=True
        ).stdout)["data"]["token"]
        
        # Remove macros
        output_path = workbook_path.replace(".xlsm", "_clean.xlsx")
        remove_result = subprocess.run(
            ["xls-remove-macros", "--input", workbook_path, "--output", output_path,
             "--token", token1, "--token", token2],
            capture_output=True, text=True
        )
        
        return {
            "status": "removed",
            "risk_level": risk_level,
            "output_path": output_path,
            "details": safety_data["data"]
        }
    
    return {
        "status": "scanned",
        "risk_level": risk_level,
        "details": safety_data["data"]
    }
```

---

## Recipe 5: Large Dataset Migration

**Purpose:** Migrate large datasets with chunked reading and batch writes.

**Duration:** ~30-60 seconds (depends on dataset size)  
**Risk Level:** Medium  
**Token Requirements:** None

### Workflow Steps

```
Chunked Read → Schema Validate → Batch Transform → Batch Write → Verify
```

### Complete JSON I/O

#### Step 1: Chunked Read

**Request:**
```bash
xls-read-range --input large_dataset.xlsx --range A1:E100000 --chunked
```

**Response (JSONL - streaming):**
```json
{"chunk": 1, "total_chunks": 10, "rows": 10000, "data": [...]}
{"chunk": 2, "total_chunks": 10, "rows": 10000, "data": [...]}
...
{"chunk": 10, "total_chunks": 10, "rows": 10000, "data": [...]}
```

---

#### Step 2: Get Workbook Metadata (for validation)

**Request:**
```bash
xls-get-workbook-metadata --input large_dataset.xlsx
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "sheet_count": 1,
    "total_formulas": 0,
    "named_ranges": [],
    "tables": [],
    "has_macros": false,
    "file_size_bytes": 5242880
  }
}
```

---

#### Step 3: Batch Write to New Workbook

**Request:**
```bash
# Create new workbook
xls-create-new --output migrated.xlsx --sheets "Data"

# Write in batches (5 chunks of 20k rows)
xls-write-range --input migrated.xlsx --output migrated.xlsx --range A1 --sheet "Data" \
  --data @chunk_1.json

xls-write-range --input migrated.xlsx --output migrated.xlsx --range A20001 --sheet "Data" \
  --data @chunk_2.json
```

**Response:**
```json
{
  "status": "success",
  "data": {"range_written": "A1:E20000"},
  "impact": {"cells_modified": 100000}
}
```

---

#### Step 4: Validate Migration

**Request:**
```bash
xls-get-workbook-metadata --input migrated.xlsx
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "sheet_count": 1,
    "total_formulas": 0,
    "file_size_bytes": 5242880
  }
}
```

**Agent Action:** Compare row counts and file sizes with source.

---

### Complete Python Implementation

```python
import json
import subprocess
from pathlib import Path

def migrate_large_dataset(source_path: str, output_path: str, chunk_size: int = 10000) -> dict:
    """
    Migrate large dataset with chunked processing.
    
    Args:
        source_path: Source large dataset
        output_path: Destination path
        chunk_size: Rows per chunk
    
    Returns:
        Migration stats
    """
    # Step 1: Create target workbook
    subprocess.run(
        ["xls-create-new", "--output", output_path, "--sheets", "Data"],
        capture_output=True, text=True
    )
    
    # Step 2: Get source metadata
    meta_result = subprocess.run(
        ["xls-get-workbook-metadata", "--input", source_path],
        capture_output=True, text=True
    )
    source_meta = json.loads(meta_result.stdout)
    
    # Step 3: Read chunked and process
    chunks = []
    read_result = subprocess.run(
        ["xls-read-range", "--input", source_path, "--range", "A1:E100000", "--chunked"],
        capture_output=True, text=True
    )
    
    # Parse JSONL response
    for line in read_result.stdout.strip().split('\n'):
        if line:
            chunks.append(json.loads(line))
    
    # Step 4: Write in batches
    total_rows = 0
    for i, chunk in enumerate(chunks):
        start_row = total_rows + 1
        range_ref = f"A{start_row}"
        
        subprocess.run(
            ["xls-write-range", "--input", output_path, "--output", output_path,
             "--range", range_ref, "--sheet", "Data", "--data", json.dumps(chunk["data"])],
            capture_output=True, text=True
        )
        
        total_rows += chunk["rows"]
    
    # Step 5: Validate
    target_result = subprocess.run(
        ["xls-get-workbook-metadata", "--input", output_path],
        capture_output=True, text=True
    )
    target_meta = json.loads(target_result.stdout)
    
    return {
        "status": "success",
        "rows_migrated": total_rows,
        "chunks_processed": len(chunks),
        "source_size": source_meta["data"]["file_size_bytes"],
        "target_size": target_meta["data"]["file_size_bytes"]
    }
```

---

## Error Handling Patterns

### Pattern: Retry on Lock Contention

```python
import time

def run_with_retry(cmd: list, max_retries: int = 3) -> dict:
    for attempt in range(max_retries):
        result = subprocess.run(cmd, capture_output=True, text=True)
        data = json.loads(result.stdout)
        
        if result.returncode == 3:  # Lock contention
            time.sleep(0.5 * (2 ** attempt))  # Exponential backoff
            continue
        
        return data
    
    raise RuntimeError("Max retries exceeded")
```

### Pattern: Parse Denial Guidance

```python
import re

def parse_guidance(denial_response: dict) -> dict:
    """Extract actionable guidance from denial."""
    guidance = denial_response.get("guidance", "")
    
    # Extract tool name
    tool_match = re.search(r'Run\s+(\S+)', guidance)
    tool = tool_match.group(1) if tool_match else None
    
    # Extract parameters
    param_match = re.search(r'--(\w+)\s+\'([^\']+)\'', guidance)
    param = {param_match.group(1): param_match.group(2)} if param_match else {}
    
    return {
        "tool": tool,
        "parameters": param,
        "raw_guidance": guidance
    }
```

---

**Document maintained by:** excel-agent-tools maintainers  
**License:** MIT
