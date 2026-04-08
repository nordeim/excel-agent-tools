### Phase 1: Analysis of PowerPoint Agent Tools (Reference Architecture)

I have meticulously reviewed `TO_REVIEW.md`. The design establishes a rigorous standard for "AI-native" file manipulation tools. Key architectural pillars include:

1.  **Hub-and-Spoke Model:** A stateful core library (`powerpoint_agent_core.py`) manages context, locking, and complex OOXML manipulation, while stateless CLI tools act as thin wrappers for AI consumption.
2.  **Governance-First Design:**
    *   **Clone-Before-Edit:** Prevents data loss by enforcing a safety copy workflow.
    *   **Approval Tokens:** HMAC-SHA256 tokens strictly enforce permissions for destructive operations (deletion, merging).
    *   **File Locking:** OS-level atomic locking prevents race conditions.
    *   **Version Hashing:** Geometry-aware hashing detects concurrent modifications or corruption.
3.  **AI-Centric Interface:**
    *   **JSON I/O:** All inputs and outputs are JSON-parsable, facilitating tool chaining.
    *   **Stdout/Stderr Hygiene:** Suppression of library noise ensures clean parsing.
    *   **Exit Codes:** Standardized error signaling (0-5) allows agents to handle retries and failures programmatically.
4.  **Validation:** XSD validation against OOXML schemas and JSON schema validation for tool payloads.

---

### Phase 2: Research & Gap Analysis for Excel Agent Tools

To replicate this level of rigor for Excel (`.xlsx`/`.xlsm`), I researched the current landscape of Python-Excel interaction, focusing on `openpyxl`, `xlwings`, and `pandas`.

**Challenges Unique to Excel:**
*   **Formula Dependencies:** Deleting a row/cell in Excel can break references (`#REF!`) in distant sheets. The tool must calculate a dependency graph before destructive mutations.
*   **Calculation Engine:** `openpyxl` reads/writes formulas but does not calculate them. We must integrate a headless calculation engine (LibreOffice) via a script, similar to `scripts/recalc.py` referenced in the documentation.
*   **Statefulness vs. Statelessness:** Excel's "grid" implies implicit relationships (merged cells, pivot caches). A stateless agent needs robust introspection tools to understand these relationships before acting.
*   **Macro Safety:** `.xlsm` files contain binary VBA projects. Modifying VBA code programmatically is high-risk and often breaks digital signatures. The design should focus on *container management* and *safe injection* rather than rewriting VBA code on the fly.
*   **Performance:** Large datasets (>100k rows) require chunked processing to avoid memory exhaustion, unlike PPTX files.

---

### Phase 3: Excel Agent Tools - Comprehensive Design Blueprint

**Project Name:** `excel-agent-tools`
**Version:** 1.0.0
**License:** MIT

#### 1. Executive Summary
`excel-agent-tools` is a governance-first orchestration layer enabling AI agents to programmatically engineer, audit, and manipulate Excel workbooks. It bridges stateless AI agents with the complex, stateful nature of spreadsheets using atomic CLI tools, strict validation, and macro-safe workflows.

#### 2. Architecture Overview

**Hub-and-Spoke with Dependency Awareness:**

```text
┌────────────────────────────────────────┐
│  AI Agent / Orchestration Layer        │
│  (Stateless, JSON-First)               │
└────────────────┬───────────────────────┘
                 │
    ┌────────────┼────────────┐
    ▼            ▼            ▼
 Tool A        Tool B      Tool C
 (40+ tools)
    │            │            │
    └────────────┼────────────┘
                 ▼
    ┌────────────────────────────────────┐
    │   excel_agent_core.py (Hub)        │
    │                                    │
    │  • ExcelAgent (context mgr)        │
    │  • Dependency Graph Analysis       │
    │  • Atomic File Locking             │
    │  • Formula/Reference Validation    │
    │  • Approval Token Validation       │
    │  • Macro Signature Check           │
    └────────────────────────────────────┘
                 │
                 ▼
    ┌────────────────────────────────────┐
    │     openpyxl + LibreOffice         │
    │  (Structure + Recalc Engine)       │
    └────────────────────────────────────┘
```

#### 3. Core Components

**A. `ExcelAgent` (Core Library)**
*   **Context Manager:** Handles `keep_vba=True` preservation and file locking.
*   **Dependency Tracker:** Before any deletion (rows, cols, sheets), the agent scans `defined_names` and `formulas` to identify impact. If a deletion causes `#REF!` errors in critical ranges, it warns or blocks based on policy.
*   **Range Normalizer:** Converts A1, R1C1, and Named Ranges into a unified internal coordinate system.
*   **Macro Auditor:** Uses `olefile` to inspect VBA project streams for signatures without breaking them.

**B. `xls_recalculate.py` (Bridge Tool)**
*   Wraps the LibreOffice headless macro execution.
*   Accepts an `.xlsx` or `.xlsm`, opens it silently, forces a full recalc, saves, and exits.
*   Returns JSON with `formula_count`, `error_count` (identifying `#REF!`, `#DIV/0!`, etc.), and `recalc_time_ms`.

#### 4. Tool Catalog (Planned 40+ Tools)

**Category 1: Creation & Structure**
*   `xls_create_new.py`: Create blank workbook.
*   `xls_create_from_template.py`: Create from `.xltx` or `.xltm`.
*   `xls_clone_workbook.py`: Atomic clone for safe editing.
*   `xls_add_sheet.py`: Add sheet with layout options (Right/Left).
*   `xls_delete_sheet.py`: ⚠️ Token-required.
*   `xls_rename_sheet.py`: Rename sheet and update cross-sheet references.

**Category 2: Cell & Range Manipulation**
*   `xls_read_range.py`: Extract data as JSON (handling dates/currencies).
*   `xls_write_range.py`: Write data with type inference.
*   `xls_insert_rows.py`: Insert rows with style copying.
*   `xls_delete_range.py`: Shift cells up/left. ⚠️ Token-required.
*   `xls_merge_cells.py`: Merge with content preservation.
*   `xls_unmerge_cells.py`: Restore grid.

**Category 3: Formulas & Data**
*   `xls_set_formula.py`: Inject formula with reference validation.
*   `xls_recalculate.py`: Trigger LibreOffice calc engine.
*   `xls_fix_formula_errors.py`: AI-guided remediation suggestions.
*   `xls_define_name.py`: Create Named Ranges.

**Category 4: Formatting & Style**
*   `xls_format_range.py`: Fonts, fills, borders, alignment.
*   `xls_apply_conditional_formatting.py`: Add rules (Icon sets, Data bars).
*   `xls_set_column_width.py`: Auto-size or fixed width.
*   `xls_freeze_panes.py`: Set view locks.

**Category 5: Tables & Objects**
*   `xls_add_table.py`: Convert range to Excel Table (ListObject).
*   `xls_add_chart.py`: Create chart from range data.
*   `xls_add_pivot_table.py**: Create PivotCache and PivotTable.
*   `xls_add_image.py`: Insert image anchored to cells.
*   `xls_add_comment.py`: Add threadable comments.

**Category 6: Macros & Security (`.xlsm`)**
*   `xls_inspect_macros.py**: List VBA modules, check for digital signatures.
*   `xls_enable_macros.py**: Ensure `keep_vba=True` is active.
*   `xls_remove_macros.py**: Strip VBA stream and convert to `.xlsx`. ⚠️ Token-required.
*   `xls_inject_macro_stub.py**: Inject a trusted "hello world" module for testing.

**Category 7: Validation & Export**
*   `xls_validate_workbook.py`: Check structural integrity, broken refs.
*   `xls_check_accessibility.py**: WCAG checks (Contrast in conditional formatting, Headers, Alt text).
*   `xls_export_pdf.py**: Export via LibreOffice.
*   `xls_export_csv.py**: Export specific sheet to CSV.

#### 5. Safety & Governance Protocols

**A. Approval Token Enforcement**
*   **Scopes:** `sheet:delete`, `range:delete`, `macro:remove`, `structure:modify`.
*   **Generation:** HMAC-SHA256 signed tokens.
*   **Validation:** Core library rejects operations if the token scope does not match the action.

**B. Clone-Before-Edit**
*   AI agents are instructed to always clone the source file to a `/work/` directory.
*   Tools verify the output path is not the source path (unless overridden by a `force` flag and token).

**C. Formula Integrity Pre-flight**
*   Before `xls_delete_range.py` or `xls_delete_sheet.py` executes, the core agent scans the target area.
*   If other formulas reference the target, the tool returns a JSON warning:
    ```json
    {
      "status": "warning",
      "impact": {
        "broken_references": 12,
        "affected_sheets": ["Summary", "Dashboard"],
        "sample_errors": ["#REF! in 'Summary'!B4"]
      },
      "suggestion": "Update references using xls_update_references.py before deletion."
    }
    ```

#### 6. Data Structures Reference

**Range Definition (Input):**
*   **A1 Style:** `"A1:C10"`
*   **Table Style:** `"Table1[Column1]"`
*   **Named Range:** `"SalesData"`
*   **Coordinate:** `{"start_row": 1, "start_col": 1, "end_row": 10, "end_col": 3}`

**Output Schema:**
*   All tools output:
    *   `status`: "success" | "error" | "warning"
    *   `workbook_version`: Hash of structure + formulas (excludes values for speed).
    *   `data`: Result payload.
    *   `impact`: Changes made (rows added, cells modified).

#### 7. Implementation Roadmap

1.  **Foundation:** Setup `excel_agent_core.py` with `openpyxl` integration, file locking, and context management.
2.  **Validation Framework:** Integrate JSON schema validation for inputs. Build the `DependencyTracker` to analyze formulas.
3.  **Core Tools:** Implement Read/Write/Create/Clone tools first.
4.  **Structure Tools:** Implement Sheet/Row/Col manipulation with dependency checks.
5.  **Macro & Advanced:** Implement `.xlsm` inspection and LibreOffice recalculation bridge.
6.  **Safety Layer:** Add HMAC token enforcement for destructive tools.
7.  **E2E Testing:** Test workflows (Clone -> Modify -> Recalculate -> Validate).

#### 8. Technical Stack

*   **Language:** Python 3.9+ (Strict typing).
*   **Core Library:** `openpyxl` (for structure/formulas), `defusedxml` (for XML parsing safety).
*   **Macro Inspection:** `olefile` (to read VBA streams safely).
*   **Calculation:** LibreOffice Headless (via `scripts/recalc.py` wrapper).
*   **CLI:** `argparse` + `json` output.
*   **Testing:** `pytest` with fixtures for `.xlsx` and `.xlsm` files.

This blueprint ensures that the Excel Agent Tools match the "Meticulous" and "Anti-Generic" standards of the PowerPoint tools while addressing the specific complexities of spreadsheet data, formulas, and macros.

# https://chat.qwen.ai/s/6bd169ac-86f4-4c10-b636-9fd4035a3ea9?fev=0.2.35 
