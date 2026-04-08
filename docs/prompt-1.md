Please meticulously review the draft design below, then meticulously research with extensive searches for an equivalent Excel tool to help an AI agent to manipulate Microsoft .xlsx spreadsheet with macros. Then meticulously plan to create a comprehensive design blueprint with architecture for such an Excel tool.

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

---

**Executive Summary**

I have conducted a meticulous review of the provided draft design for `excel-agent-tools`. The draft demonstrates a strong understanding of the governance-first, AI-native architecture established by the PowerPoint agent tools. It correctly identifies critical challenges unique to Excel: formula dependency resolution, macro safety, and the need for an external calculation engine.

Building upon this foundation, I have performed extensive research into the Python-Excel ecosystem, macro manipulation libraries, and headless calculation strategies. The following response presents an **enhanced and comprehensive design blueprint**. This blueprint refines the draft with deeper technical specifications, addresses identified gaps (e.g., VBA project integrity, large-file streaming, and cross-platform consistency), and aligns fully with the **Meticulous Approach** and **Anti-Generic** philosophy.

**Key Enhancements Over the Draft:**
- **Granular Formula Dependency Graph:** Explicit implementation details using `openpyxl` tokenization and topological sorting.
- **Macro Safety Protocol:** A tiered approach for VBA inspection, signature preservation, and safe `.xlsm` handling.
- **Cross-Platform Recalculation:** Definitive integration of LibreOffice headless with fallback strategies.
- **Performance Optimization:** Chunked streaming for reading/writing large ranges (>100k rows) to maintain sub-second CLI response times.
- **Tool Taxonomy Refinement:** Categorization of 40+ tools with explicit safety flags and token scopes.

---

### Detailed Plan: Enhanced Excel Agent Tools Blueprint

**Project Codename:** `excel-agent-tools`  
**Version:** 1.0.0  
**Design Standard:** Meticulous / Governance-First / AI-Native

#### 1. Research Findings and Technology Selection

Following the **Meticulous Approach Phase 1 (Analyze)**, I have conducted a multi-dimensional analysis of the Python-Excel ecosystem. The following table synthesizes the research and justifies the chosen technology stack.

| Library / Tool | Domain | Suitability for AI Agent Tools | Verdict & Rationale |
| :--- | :--- | :--- | :--- |
| **`openpyxl`** | Core Read/Write | Excellent for structural manipulation, formula reading/writing, styling. Lacks calculation engine. | **Selected.** It is the de facto standard for `.xlsx`/`.xlsm` without Excel installed. Its object model aligns with the OOXML specification. |
| **`xlwings`** | Automation (Windows/macOS) | Requires Excel application instance. Unsuitable for headless server environments where agents operate. | **Rejected** for core library. It introduces stateful Excel application dependencies and platform lock-in. |
| **`pandas`** | Data Analysis | Fast for data import/export, but strips formulas and macros. Destructive to workbook structure. | **Rejected** as primary manipulator. Will be used **internally** for high-performance data array transformations *only* when formulas are not present. |
| **`oletools` / `olefile`** | VBA/Macro Inspection | `olefile` provides safe, read-only access to OLE streams containing VBA projects. | **Selected.** Essential for `xls_inspect_macros.py` without modifying or breaking the binary VBA signature. |
| **LibreOffice Headless** | Formula Calculation | The only reliable, cross-platform, open-source method to force recalculation of formulas in `.xlsx`/`.xlsm` files. | **Selected.** Will be wrapped in `scripts/recalc.py` with strict timeout and process isolation. |
| **`defusedxml`** | Security | Protects against XML External Entity (XXE) attacks when parsing OOXML parts. | **Selected.** Mandatory for any tool processing user-supplied or AI-generated `.xlsx` files. |

#### 2. Enhanced Architecture: The Dependency-Aware Core

The draft architecture is solid. The enhancement lies in the **Dependency Tracker** component within `excel_agent_core.py`.

```text
excel_agent_core.py (Hub)
│
├── ExcelAgent (Context Manager)
│   ├── _acquire_lock()           # fcntl / msvcrt atomic lock
│   ├── _load_workbook()          # openpyxl load with keep_vba=True, data_only=False
│   └── _save_workbook()          # Version hash update, lock release
│
├── FormulaDependencyGraph (NEW)   # Critical for safe mutations
│   ├── build_graph(sheet)        # Parses tokens from formula strings
│   ├── find_dependents(cell)     # Returns list of cells that reference this cell
│   └── validate_deletion(target) # Pre-flight check before xls_delete_*
│
├── MacroHandler (NEW)             # Safe .xlsm interaction
│   ├── has_vba_project(path)     # olefile check for "_VBA_PROJECT_CUR" stream
│   ├── get_vba_modules(path)     # Extracts module names (no code decompilation)
│   └── preserve_vba_flag()       # Ensures keep_vba=True during save
│
└── RangeSerializer                # Converts A1/Table/Name to (min_col, min_row, max_col, max_row)
```

**Decision Rationale:**
- **Graph vs. Full Parse:** Building a full dependency graph from all formula tokens is computationally feasible for files up to ~50MB. For larger files, we implement a lazy graph: we only scan formulas in sheets that share references with the target of a mutation.
- **Macro Safety:** We will **never** attempt to parse or modify VBA source code. The tool will only detect presence, report module names, and ensure the binary stream is preserved intact during save operations. This prevents accidental corruption of digitally signed macros.

#### 3. Tool Catalog: 42 Meticulous Tools

Each tool adheres to the **Hub-and-Spoke Model** with strict JSON I/O and standardized exit codes. Tools marked with ⚠️ require an **Approval Token** scoped to the specific destructive action.

**Category 1: Governance & Environment (5 Tools)**
| Tool Name | Description | Safety Flags |
| :--- | :--- | :--- |
| `xls_clone_workbook.py` | Creates a safe copy for editing. Verifies source hash. | Read-only on source. |
| `xls_validate_workbook.py` | Checks OOXML compliance, broken refs, and circular refs. | Read-only. |
| `xls_approve_token.py` | Utility to generate HMAC tokens for AI orchestrator. | Internal Use. |
| `xls_version_hash.py` | Calculates geometry hash (sheet structure + formulas). | Read-only. |
| `xls_lock_status.py` | Checks if a file is currently locked by another process. | Read-only. |

**Category 2: Core Read & Write (7 Tools)**
| Tool Name | Performance Consideration |
| :--- | :--- |
| `xls_read_range.py` | **Chunked streaming** for >100k rows. Returns JSON Lines. |
| `xls_write_range.py` | Accepts JSON Lines input. Type inference (date, bool, float). |
| `xls_get_sheet_names.py` | Returns index, name, visibility state. |
| `xls_get_defined_names.py` | Returns global and local named ranges. |
| `xls_get_table_info.py` | Returns ListObject schema (columns, totals, style). |
| `xls_get_cell_style.py` | Returns font, fill, border, number format for a cell. |
| `xls_get_formula.py` | Returns formula string or `None`. |

**Category 3: Structural Mutation (8 Tools) ⚠️ Token Required**
| Tool Name | Dependency Check | Action |
| :--- | :--- | :--- |
| `xls_delete_sheet.py` | Scans all formulas for cross-sheet refs. | Deletes sheet. |
| `xls_rename_sheet.py` | Updates all defined names and formula refs. | Renames sheet. |
| `xls_insert_rows.py` | Updates formula offsets (e.g., `A1` -> `A2`). | Inserts rows. |
| `xls_delete_rows.py` | **Pre-flight impact report**. Blocks if refs break. | Deletes rows. |
| `xls_insert_columns.py` | Updates formula offsets. | Inserts columns. |
| `xls_delete_columns.py` | **Pre-flight impact report**. | Deletes columns. |
| `xls_merge_cells.py` | Warns if merging overlaps data in hidden cells. | Merges range. |
| `xls_unmerge_cells.py` | Restores original grid. | Unmerges range. |

**Category 4: Formulas & Calculation (5 Tools)**
| Tool Name | Description |
| :--- | :--- |
| `xls_set_formula.py` | Validates formula syntax before injection. |
| `xls_recalculate.py` | **Bridges to LibreOffice.** Returns recalc time and error count. |
| `xls_detect_errors.py` | Scans for `#REF!`, `#VALUE!`, etc. Returns coordinates. |
| `xls_convert_to_values.py` | Replaces formulas with their calculated values. ⚠️ Token. |
| `xls_copy_formula_down.py` | Auto-fill behavior for a column. |

**Category 5: Objects & Visualization (7 Tools)**
| Tool Name | Description |
| :--- | :--- |
| `xls_add_table.py` | Converts range to Excel Table. |
| `xls_add_chart.py` | Supports Bar, Line, Pie. Anchors to cells. |
| `xls_add_pivot_table.py` | Creates PivotCache and layout. |
| `xls_add_image.py` | Inserts image with aspect ratio lock. |
| `xls_add_sparkline.py` | In-cell miniature chart. |
| `xls_add_comment.py` | Threaded comment compatible with modern Excel. |
| `xls_set_data_validation.py` | Adds dropdown lists or numeric constraints. |

**Category 6: Formatting & Styling (5 Tools)**
| Tool Name | Description |
| :--- | :--- |
| `xls_format_range.py` | Applies `openpyxl.styles` objects from JSON definition. |
| `xls_set_column_width.py` | Auto-fit based on content sample or fixed width. |
| `xls_set_row_height.py` | Fixed height or auto-fit. |
| `xls_freeze_panes.py` | Sets view anchor point. |
| `xls_apply_conditional_formatting.py` | Supports ColorScale, DataBar, IconSet. |

**Category 7: Macros & Security (.xlsm) (5 Tools)**
| Tool Name | Description | Safety Protocol |
| :--- | :--- | :--- |
| `xls_inspect_macros.py` | Lists VBA modules, signature status. | **Read-only OLE access.** |
| `xls_has_macros.py` | Boolean check for VBA stream presence. | Read-only. |
| `xls_enable_macros.py` | Ensures `keep_vba=True` flag is active for current session. | Internal State. |
| `xls_remove_macros.py` | Strips VBA stream, converts to `.xlsx`. | ⚠️ Token + **Double Confirmation**. |
| `xls_validate_macro_safety.py` | Scans for known high-risk VBA functions (Shell, CreateObject). | Read-only heuristic check. |

#### 4. Implementation Phases (Meticulous Roadmap)

This plan adheres to the **Phase 4: IMPLEMENT** structure, breaking work into modular, testable units.

| Phase | Duration | Deliverables | Success Criteria |
| :--- | :--- | :--- | :--- |
| **Phase 1: Core Foundation** | 1 Week | `excel_agent_core.py`, `RangeSerializer`, File Locking, JSON Schema Validation. | Unit tests pass for lock acquisition, A1 parsing, and loading/saving `.xlsx`. |
| **Phase 2: Dependency Graph** | 1 Week | `FormulaDependencyGraph` implementation. Pre-flight checks for deletion tools. | Integration tests confirm `#REF!` warnings are generated when deleting a cell referenced elsewhere. |
| **Phase 3: Read/Write Tools** | 1 Week | `xls_read_range.py`, `xls_write_range.py`, `xls_get_*` tools. | Large file (100k rows) read completes in <2s. Data types preserved correctly. |
| **Phase 4: Structural Tools** | 2 Weeks | All `xls_delete_*`, `xls_insert_*`, and `xls_rename_sheet.py`. | Mutation tests verify formula offsets update correctly and cross-sheet references remain intact. |
| **Phase 5: Recalc & Macros** | 1 Week | `xls_recalculate.py` (LibreOffice bridge), `xls_inspect_macros.py`. | Recalc bridge successfully forces calculation and returns error counts on test file. Macro inspection correctly identifies VBA modules. |
| **Phase 6: Objects & Visuals** | 1 Week | Charting, Tables, Pivot Tables, Images. | Generated files open in Excel without repair prompts. |
| **Phase 7: Security & Tokens** | 3 Days | HMAC token validation integrated into all ⚠️ tools. | Unauthorized destructive action returns error code 4 (Permission Denied). |
| **Phase 8: Documentation & QA** | 3 Days | Full CLI help, Meticulous Approach runbook, Anti-Generic design justification. | Documentation allows a new AI agent to perform a complex multi-step task (clone -> analyze -> modify chart -> recalc) using only the CLI tools. |

#### 5. Validation Strategy (Phase 5: VERIFY)

- **Unit Tests:** Cover `RangeSerializer` edge cases (A1, R1C1, Table[#All]).
- **Integration Tests:** Simulate AI agent workflows using `subprocess` calls.
- **Stress Tests:** 500k row file manipulation to ensure memory stability.
- **Cross-Platform Tests:** Validate LibreOffice recalc path on Linux (Ubuntu) and Windows (WSL/PS).

#### 6. Next Steps (Phase 3: VALIDATE)

**Request for Confirmation:**

The enhanced design blueprint is complete. Before proceeding to **Phase 4: IMPLEMENT** (coding the core foundation), I require explicit validation on the following points:

1.  **Macro Handling:** Do you approve the strict read-only/container management approach for VBA? (We will **not** attempt to generate or modify VBA code).
2.  **Recalc Engine:** Do you accept the dependency on LibreOffice headless being installed on the host environment?
3.  **Token Scope:** Are the proposed scopes (`sheet:delete`, `range:delete`, `macro:remove`) sufficient for your governance model?
