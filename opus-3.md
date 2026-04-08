I'll conduct extensive research across multiple dimensions to validate, challenge, and enhance this blueprint. Let me search for the current state of the art in Python-Excel manipulation, macro handling, AI agent tool frameworks, and related areas.---

# Comprehensive Design Blueprint Review & Enhanced Architecture for `excel-agent-tools`

## I. Research Synthesis & Critical Findings

My extensive research across the Python-Excel ecosystem, AI agent frameworks, formula calculation engines, and macro security tools has validated core assumptions in your draft while revealing several critical enhancements and one significant architectural correction. Here is my full synthesis:

### A. Core Library Validation: `openpyxl` is Correct

openpyxl is an efficient Python library that enables you to read and write Excel files in the modern XML-based formats (.xlsx, .xlsm) introduced in Excel 2007, with full support for advanced features from Excel 2010 onward. One major advantage of openpyxl is that it doesn't require Excel to be installed on your computer, making it ideal for server environments, automated reporting systems, and data processing pipelines. This confirms it as the correct choice for headless AI agent environments. The latest version is 3.1.5 and the package is stable and actively maintained.

**Critical security finding:** By default openpyxl does not guard against quadratic blowup or billion laughs xml attacks. To guard against these attacks install defusedxml. This validates the draft's inclusion of `defusedxml` as **mandatory**, not optional.

### B. Formula Dependency Graph: Major Enhancement Discovered

The draft's `FormulaDependencyGraph` concept is correct but underspecified. Research uncovered **four** independent Python libraries that can inform or replace the planned implementation:

1. **`pycel`:** Pycel is a small python library that can translate an Excel spreadsheet into executable python code which can be run independently of Excel. The python code is based on a graph and uses caching & lazy evaluation to ensure (relatively) fast execution. Critically, it has been tested extensively on spreadsheets with 10 sheets & more than 10000 formulae. In that case calculation of the equations takes about 50ms and agrees with Excel up to 5 decimal places. However, the code currently uses a tokenizer of similar origin from the openpyxl library, meaning it shares the same parsing foundation.

2. **`xlcalculator`:** xlcalculator is a Python library that reads MS Excel files and, to the extent of supported functions, can translate the Excel functions into Python code and subsequently evaluate the generated Python code. Essentially doing the Excel calculations without the need for Excel. xlcalculator is a modernization of the koala2 library.

3. **`formulas`:** formulas implements an interpreter for Excel formulas, which parses and compile Excel formulas expressions. Moreover, it compiles Excel workbooks to python and executes without using the Excel COM server. Hence, Excel is not needed. This library uniquely offers a command-line interface that works with spreadsheet models and accepts .xlsx, .ods, and .json inputs. A typical workflow starts by calculating a workbook. You can override input values directly from the command line and request specific cells to be rendered in the output.

4. **`xlcalcmodel`:** xlcalcmodel is a high-performance, Python-based Excel calculation engine that transforms Excel workbooks into tokenized, JSON-based models for rapid, standalone computation. It achieves this by converting spreadsheet formulas into refined abstract syntax trees (ASTs) and executing them in a rules-engine fashion that faithfully mimics Microsoft Excel. It provides configurable iterative calculation support that mirrors Excel's "Enable Iterative Calculation" option, allowing for convergence in workbooks with circular references.

**Architecture Decision:** This is a critical upgrade over the draft. The design should adopt a **tiered calculation strategy**:
- **Tier 1 (Fast/In-process):** Use `formulas` or `pycel` for dependency graph construction and lightweight recalculation of common functions.
- **Tier 2 (Full fidelity):** Use LibreOffice headless as the authoritative recalculation engine for full Excel compatibility (VLOOKUP, complex array formulas, etc.).
- This eliminates the draft's hard dependency on LibreOffice for *every* recalculation while preserving it as a fallback.

### C. Macro Safety: `oletools` is Superior to `olefile` Alone

The draft specifies `olefile` for macro inspection. Research confirms this should be upgraded to `oletools` (which includes `olefile`):

olevba is a script to parse OLE and OpenXML files such as MS Office documents (e.g. Word, Excel), to detect VBA Macros, extract their source code in clear text, and detect security-related patterns such as auto-executable macros, suspicious VBA keywords used by malware, anti-sandboxing and anti-virtualization techniques, and potential IOCs. XLM/Excel 4 Macros are also supported in Excel and SLK files. It can be used either as a command-line tool, or as a python module from your own applications.

Crucially for the `xls_validate_macro_safety.py` tool: The function detect_autoexec checks if VBA macro code contains specific macro names that will be triggered when the document/workbook is opened, closed, changed, etc. It returns a list of tuples containing two strings, the detected keyword, and the description of the trigger. And the function detect_suspicious checks if VBA macro code contains specific keywords often used by malware to act on the system (create files, run commands or applications, write to the registry, etc).

oletools can detect, extract and analyse VBA macros, OLE objects, Excel 4 macros (XLM) and DDE links. oletools are used by a number of projects and online malware analysis services, including Viper, REMnux, FAME, Hybrid-analysis.com, Joe Sandbox, Deepviz, Laika BOSS, Cuckoo Sandbox, and probably VirusTotal.

**One caveat:** Further analysis of the maintenance status of oletools determined that its maintenance is Inactive. An important project maintenance signal to consider is that it hasn't seen any new versions released to PyPI in the past 12 months. This means the tool should isolate `oletools` behind an abstraction layer so it can be swapped for a maintained fork if necessary.

### D. VBA Binary Container: The Draft's "Never Modify VBA" Approach is Correct

The vbaProject.bin file is a binary OLE COM container. This was the format used in older xls versions of Excel prior to Excel 2007. Unlike all of the other components of an xlsx/xlsm file the data isn't stored in XML format. Instead the functions and macros as stored as a pre-parsed binary format. As such it wouldn't be feasible to define macros and create a vbaProject.bin file from scratch.

For safe macro injection, the pattern should follow XlsxWriter's approach: The vba_extract.py utility is used to extract the vbaProject.bin binary from an Excel 2007+ xlsm file. If the VBA project is signed, vba_extract.py also extracts the vbaProjectSignature.bin file from the xlsm file. For signature handling: The vba_extract.py utility can be used to extract the vbaProject.bin and vbaProjectSignature.bin files from an existing xlsm file with signed macros. To add these files to the XlsxWriter workbook using the add_signed_vba_project() method.

### E. Competitive Landscape: What Existing AI Agent Tools Lack

The research reveals a clear gap in the market that this project fills:

- **LlamaIndex Spreadsheet Agent:** Spreadsheets represent one of the most challenging document types for AI systems. Unlike clean CSV files or structured databases, spreadsheets are designed for human consumption, not machine readability. Their approach uses RL to parse the general structure of the spreadsheet — fundamentally different from our governance-first tool-based architecture.

- **AI Agent for Excel (Microsoft Marketplace):** While most AI tools in Excel provide guidance in the form of text, AI Agent for Excel suggests cell edits—both data and formulas—that users can apply with a single click. It is therefore more automated and enables complex tasks. However, the add-in is built with Excel JavaScript API 1.2 — it requires Excel to be running. Our tool is headless.

- **Shortcut AI:** Verifiably accurate outputs and mathematical precision are prioritized above all else. Outputs are formula-driven to ensure they update dynamically with your data—not hard-coded values that break when inputs change. This validates our design principle of formula integrity, but Shortcut is a closed SaaS product, not an open toolset for agent orchestration.

**None of these competitors provide:**
- Governance tokens for destructive operations
- Clone-before-edit safety workflows
- Formula dependency pre-flight checks before mutations
- OS-level file locking for concurrent access
- Headless, JSON-first CLI tools for agent tool chaining

This confirms the draft's unique positioning.

### F. Formula Dependency Tracing: Validated by Independent Research

The TreeNode class forms the foundation of the solution. It represents each formula and its associated cell in the workbook as a node in a tree structure. This is how we model the relationships between formulas and their referenced cells. To retrieve formulas from Excel cells, we use the openpyxl library. This ensures that all dependencies are captured, even across multiple sheets.

Additionally, there is an existing open-source project for figuring out dependencies between cells in excel on python, which can serve as reference implementation material.

---

## II. Enhanced & Comprehensive Design Blueprint

**Project Codename:** `excel-agent-tools`
**Version:** 1.0.0
**License:** MIT
**Design Standard:** Meticulous / Governance-First / AI-Native

### 1. Revised Technology Stack

| Component | Draft Choice | **Revised Choice** | Rationale |
|:---|:---|:---|:---|
| Core I/O | `openpyxl` | **`openpyxl` (confirmed)** | Stable, headless, full .xlsx/.xlsm support |
| XML Security | `defusedxml` | **`defusedxml` (mandatory)** | Prevents XXE, billion laughs attacks |
| Macro Inspection | `olefile` | **`oletools` (upgraded)** | `olevba` provides VBA extraction, auto-exec detection, suspicious keyword scanning, XLM/Excel 4 macro support |
| Formula Graph | Custom | **`formulas` + custom `DependencyTracker`** | `formulas` provides AST parsing, CLI, dependency resolution; custom layer adds deletion impact analysis |
| Calculation Tier 1 | — (new) | **`formulas` or `pycel`** | In-process Python recalc for common functions (~50ms for 10k formulas) |
| Calculation Tier 2 | LibreOffice | **LibreOffice Headless (confirmed)** | Full-fidelity recalc for complex/unsupported functions |
| Data Arrays | — | **`pandas` (internal only)** | For chunked I/O of large ranges (>100k rows); never exposed to agent directly |
| CLI | `argparse` | **`argparse` + strict JSON schema validation** | All inputs validated against JSON schemas before execution |
| Testing | `pytest` | **`pytest` + `hypothesis` (property-based)** | Catch edge cases in range parsing, formula offset calculation |

### 2. Enhanced Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                    AI Agent / Orchestration Layer                    │
│              (Stateless, JSON-First, Tool-Chaining)                 │
└──────────────────────────────┬──────────────────────────────────────┘
                               │ JSON stdin/stdout
              ┌────────────────┼────────────────┐
              ▼                ▼                ▼
          Tool A           Tool B           Tool C
       (42+ CLI tools, each with --help, JSON output, exit codes 0-5)
              │                │                │
              └────────────────┼────────────────┘
                               ▼
┌──────────────────────────────────────────────────────────────────────┐
│                     excel_agent_core.py (Hub)                        │
│                                                                      │
│  ┌──────────────────┐  ┌───────────────────────┐  ┌──────────────┐  │
│  │  ExcelAgent       │  │  DependencyTracker     │  │ MacroHandler │  │
│  │  (Context Mgr)    │  │  (Formula Graph)       │  │ (oletools)   │  │
│  │                   │  │                        │  │              │  │
│  │ • _acquire_lock() │  │ • build_graph(wb)      │  │ • detect()   │  │
│  │ • _load(keep_vba) │  │ • find_dependents(ref) │  │ • inspect()  │  │
│  │ • _save()         │  │ • impact_report(target)│  │ • scan_risk()│  │
│  │ • _version_hash() │  │ • circular_ref_check() │  │ • preserve() │  │
│  └──────────────────┘  └───────────────────────┘  └──────────────┘  │
│                                                                      │
│  ┌──────────────────┐  ┌───────────────────────┐  ┌──────────────┐  │
│  │ RangeSerializer   │  │  ApprovalTokenMgr      │  │ ChunkedIO    │  │
│  │ (A1/R1C1/Name/   │  │  (HMAC-SHA256)         │  │ (pandas      │  │
│  │  Table/Coord)    │  │                        │  │  streaming)  │  │
│  └──────────────────┘  └───────────────────────┘  └──────────────┘  │
└──────────────────────────────┬───────────────────────────────────────┘
                               │
              ┌────────────────┼────────────────┐
              ▼                ▼                ▼
    ┌──────────────┐  ┌──────────────┐  ┌──────────────────────┐
    │  openpyxl     │  │  formulas/   │  │  LibreOffice Headless │
    │  (Structure)  │  │  pycel       │  │  (Full Recalc)        │
    │              │  │  (Tier 1     │  │  (Tier 2 Calc)        │
    │              │  │   Calc)      │  │                       │
    └──────────────┘  └──────────────┘  └──────────────────────┘
```

### 3. Core Component Specifications

#### A. `ExcelAgent` Context Manager

```python
class ExcelAgent:
    """Stateful context manager for safe workbook manipulation."""

    def __init__(self, path: Path, *, mode: str = "rw", keep_vba: bool = True):
        self._path = path
        self._mode = mode
        self._keep_vba = keep_vba
        self._lock_fd: Optional[int] = None
        self._wb: Optional[Workbook] = None
        self._entry_hash: Optional[str] = None

    def __enter__(self) -> 'ExcelAgent':
        self._acquire_lock()          # fcntl.flock / msvcrt.locking
        self._wb = load_workbook(
            str(self._path),
            keep_vba=self._keep_vba,
            data_only=False            # ALWAYS preserve formulas
        )
        self._entry_hash = self._compute_version_hash()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None and self._mode == "rw":
            self._verify_no_concurrent_modification()
            self._wb.save(str(self._path))
        self._release_lock()

    def _compute_version_hash(self) -> str:
        """Geometry-aware hash: sheet names + cell coordinates + formulas.
        Excludes values for speed (formulas are the structural contract)."""
        ...

    def _verify_no_concurrent_modification(self):
        """Re-reads file hash; raises ConcurrentModificationError if changed."""
        ...
```

#### B. `DependencyTracker` (Formula Graph Engine)

This is the most critical component for safe mutations. It integrates with `formulas` library for AST parsing:

```python
class DependencyTracker:
    """Builds and queries the workbook's formula dependency graph."""

    def __init__(self, workbook: Workbook):
        self._wb = workbook
        self._graph: Dict[str, Set[str]] = {}  # cell -> {cells that reference it}
        self._reverse: Dict[str, Set[str]] = {} # cell -> {cells it references}

    def build_graph(self, sheets: Optional[List[str]] = None):
        """Parse all formulas into AST, extract cell references,
        build directed acyclic graph of dependencies.
        Uses openpyxl Tokenizer for formula parsing."""
        ...

    def find_dependents(self, target: str) -> Set[str]:
        """Returns all cells that would break if 'target' is deleted.
        Performs transitive closure (A->B->C: deleting A affects C)."""
        ...

    def impact_report(self, target_range: str, action: str) -> dict:
        """Pre-flight check. Returns JSON-serializable impact report:
        {
            "status": "warning" | "safe",
            "broken_references": int,
            "affected_sheets": [...],
            "sample_errors": ["#REF! in 'Summary'!B4", ...],
            "suggestion": "Run xls_update_references.py first"
        }"""
        ...

    def detect_circular_references(self) -> List[List[str]]:
        """Topological sort; returns cycles if they exist."""
        ...
```

#### C. `MacroHandler` (Upgraded with `oletools`)

```python
class MacroHandler:
    """Safe, read-only inspection of VBA projects in .xlsm files.
    Uses oletools.olevba for deep analysis. Never modifies VBA binary."""

    def has_vba_project(self, path: Path) -> bool:
        """Uses olevba.detect_vba_macros() for definitive check."""
        ...

    def get_vba_modules(self, path: Path) -> List[dict]:
        """Returns module names, types (Standard/Class/Form), sizes.
        Uses olevba.extract_macros() for enumeration only."""
        ...

    def scan_risk(self, path: Path) -> dict:
        """Returns security risk assessment:
        - auto_executable: list of auto-exec triggers (AutoOpen, etc.)
        - suspicious_keywords: list of dangerous functions (Shell, CreateObject)
        - iocs: extracted IP addresses, URLs, filenames
        Uses olevba.detect_autoexec() and olevba.detect_suspicious()"""
        ...

    def has_digital_signature(self, path: Path) -> bool:
        """Checks for vbaProjectSignature.bin in OLE structure."""
        ...

    def preserve_vba_on_save(self, wb: Workbook) -> None:
        """Ensures keep_vba=True flag is maintained throughout pipeline."""
        ...
```

### 4. Revised Tool Catalog (43 Tools)

All tools produce standardized JSON output:
```json
{
    "status": "success | error | warning",
    "exit_code": 0,
    "workbook_version": "sha256:abc123...",
    "data": { },
    "impact": { "cells_modified": 0, "formulas_updated": 0 },
    "warnings": []
}
```

Exit codes: `0`=success, `1`=validation error, `2`=file not found, `3`=lock contention, `4`=permission denied (bad token), `5`=internal error.

#### Category 1: Governance & Environment (6 Tools)
| # | Tool | Description | Safety |
|---|:---|:---|:---|
| 1 | `xls_clone_workbook.py` | Atomic copy to `/work/` directory. Computes entry hash. | Read-only source |
| 2 | `xls_validate_workbook.py` | OOXML compliance, broken refs, circular ref detection | Read-only |
| 3 | `xls_approve_token.py` | Generate scoped HMAC-SHA256 approval tokens | Internal |
| 4 | `xls_version_hash.py` | Geometry hash (structure + formulas, excludes values) | Read-only |
| 5 | `xls_lock_status.py` | Check OS-level file lock state | Read-only |
| 6 | `xls_dependency_report.py` | **NEW.** Full dependency graph export as JSON adjacency list | Read-only |

#### Category 2: Core Read (7 Tools)
| # | Tool | Description | Performance |
|---|:---|:---|:---|
| 7 | `xls_read_range.py` | Extract data as JSON; handles dates/currencies/booleans | Chunked for >100k rows |
| 8 | `xls_get_sheet_names.py` | Returns index, name, visibility, cell count | Instant |
| 9 | `xls_get_defined_names.py` | Global and sheet-scoped named ranges | Instant |
| 10 | `xls_get_table_info.py` | ListObject schema, columns, totals row, style name | Instant |
| 11 | `xls_get_cell_style.py` | Font, fill, border, number_format, alignment as JSON | Instant |
| 12 | `xls_get_formula.py` | Formula string or `null`; includes parsed references | Instant |
| 13 | `xls_get_workbook_metadata.py` | **NEW.** Sheet count, total formulas, named ranges, table count | Instant |

#### Category 3: Core Write (4 Tools)
| # | Tool | Description | Safety |
|---|:---|:---|:---|
| 14 | `xls_create_new.py` | Create blank workbook with optional sheet names | — |
| 15 | `xls_create_from_template.py` | Clone from .xltx/.xltm with variable substitution | — |
| 16 | `xls_write_range.py` | Write data with type inference (date, bool, float, str) | Validates target exists |
| 17 | `xls_write_cell.py` | **NEW.** Single-cell write with explicit type coercion | Validates target exists |

#### Category 4: Structural Mutation (8 Tools) — ⚠️ Token Required
| # | Tool | Description | Dependency Check |
|---|:---|:---|:---|
| 18 | `xls_add_sheet.py` | Add sheet at position (before/after) | — |
| 19 | `xls_delete_sheet.py` ⚠️ | Delete sheet | Full cross-sheet formula scan |
| 20 | `xls_rename_sheet.py` ⚠️ | Rename + update all cross-sheet refs | Auto-updates formulas |
| 21 | `xls_insert_rows.py` | Insert rows with style inheritance | Updates formula offsets |
| 22 | `xls_delete_rows.py` ⚠️ | Delete rows | Pre-flight impact report |
| 23 | `xls_insert_columns.py` | Insert columns | Updates formula offsets |
| 24 | `xls_delete_columns.py` ⚠️ | Delete columns | Pre-flight impact report |
| 25 | `xls_move_sheet.py` | **NEW.** Reorder sheet position | — |

#### Category 5: Cell Operations (4 Tools)
| # | Tool | Description | Safety |
|---|:---|:---|:---|
| 26 | `xls_merge_cells.py` | Merge range; warns if data in non-anchor cells | Pre-check for hidden data |
| 27 | `xls_unmerge_cells.py` | Restore grid from merged range | — |
| 28 | `xls_delete_range.py` ⚠️ | Shift cells up/left after clearing | Impact report |
| 29 | `xls_update_references.py` | **NEW.** Batch-update cell references after structural changes | — |

#### Category 6: Formulas & Calculation (6 Tools)
| # | Tool | Description | Engine |
|---|:---|:---|:---|
| 30 | `xls_set_formula.py` | Inject formula with syntax validation | openpyxl tokenizer |
| 31 | `xls_recalculate.py` | Force recalculation | Tier 1: `formulas`; Tier 2: LibreOffice |
| 32 | `xls_detect_errors.py` | Scan for `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?` | Read-only |
| 33 | `xls_convert_to_values.py` ⚠️ | Replace formulas with calculated values | Token required |
| 34 | `xls_copy_formula_down.py` | Auto-fill formula along column/row | Reference adjustment |
| 35 | `xls_define_name.py` | Create/update named ranges | — |

#### Category 7: Objects & Visualization (5 Tools)
| # | Tool | Description |
|---|:---|:---|
| 36 | `xls_add_table.py` | Convert range to Excel Table (ListObject) |
| 37 | `xls_add_chart.py` | Bar, Line, Pie, Scatter; anchored to cell |
| 38 | `xls_add_image.py` | Insert image with aspect ratio preservation |
| 39 | `xls_add_comment.py` | Threaded comment (modern Excel format) |
| 40 | `xls_set_data_validation.py` | Dropdown lists, numeric constraints |

#### Category 8: Formatting & Style (5 Tools)
| # | Tool | Description |
|---|:---|:---|
| 41 | `xls_format_range.py` | Fonts, fills, borders, alignment from JSON spec |
| 42 | `xls_set_column_width.py` | Auto-fit (sample-based) or fixed width |
| 43 | `xls_freeze_panes.py` | Set freeze point for scroll lock |
| 44 | `xls_apply_conditional_formatting.py` | ColorScale, DataBar, IconSet rules |
| 45 | `xls_set_number_format.py` | **NEW.** Apply number format codes (currency, %, date) |

#### Category 9: Macros & Security — `.xlsm` (5 Tools)
| # | Tool | Description | Engine |
|---|:---|:---|:---|
| 46 | `xls_has_macros.py` | Boolean check for VBA project presence | oletools |
| 47 | `xls_inspect_macros.py` | List VBA modules, types, signature status | oletools `olevba` |
| 48 | `xls_validate_macro_safety.py` | Risk scan: auto-exec, Shell, CreateObject, IOCs | oletools `detect_suspicious` |
| 49 | `xls_remove_macros.py` ⚠️⚠️ | Strip VBA → convert to .xlsx | Double-token + confirmation |
| 50 | `xls_inject_vba_project.py` ⚠️ | Inject pre-extracted .bin with optional signature | Token required |

#### Category 10: Export & Interop (3 Tools)
| # | Tool | Description |
|---|:---|:---|
| 51 | `xls_export_pdf.py` | Export via LibreOffice headless |
| 52 | `xls_export_csv.py` | Export specific sheet to CSV with encoding control |
| 53 | `xls_export_json.py` | **NEW.** Export sheet/range as structured JSON |

### 5. Safety & Governance Protocol (Enhanced)

#### A. Approval Token Scopes (Expanded)

| Scope | Required By | Risk Level |
|:---|:---|:---|
| `sheet:delete` | `xls_delete_sheet.py` | High |
| `sheet:rename` | `xls_rename_sheet.py` | Medium (formula side-effects) |
| `range:delete` | `xls_delete_rows/columns/range.py` | High |
| `formula:convert` | `xls_convert_to_values.py` | High (irreversible data loss) |
| `macro:remove` | `xls_remove_macros.py` | Critical (double-token) |
| `macro:inject` | `xls_inject_vba_project.py` | High |
| `structure:modify` | Any structural mutation with `--force` | Medium |

#### B. Formula Integrity Pre-flight Protocol

Before **every** destructive operation, the tool **must**:
1. Invoke `DependencyTracker.impact_report(target)`
2. If `broken_references > 0`: return `exit_code=1` with the impact JSON
3. The agent can then call `xls_update_references.py` to resolve, or pass `--acknowledge-impact` flag with a valid token to force execution
4. All impact reports are logged to an audit trail file (`.excel_agent_audit.jsonl`)

#### C. Clone-Before-Edit Enforcement

```
Source file: /data/financials.xlsx
  ↓ xls_clone_workbook.py
Working copy: /work/financials_20260408T143022_abc123.xlsx
  ↓ All mutations happen here
  ↓ xls_validate_workbook.py (final check)
Output: /output/financials_final.xlsx
```

Tools verify `output_path != source_path` unless `--force --token <scope:structure:modify>` is provided.

### 6. Implementation Roadmap (Revised)

| Phase | Duration | Deliverables | Gate Criteria |
|:---|:---|:---|:---|
| **Phase 1: Core Foundation** | 5 days | `ExcelAgent`, `RangeSerializer`, file locking, JSON schema validation, exit code framework | Lock/unlock cycle test passes; A1/R1C1/Name parsing 100% coverage |
| **Phase 2: Dependency Engine** | 5 days | `DependencyTracker` using `formulas` library AST; `xls_dependency_report.py` | Correctly identifies all cross-sheet dependencies in 5-sheet test workbook |
| **Phase 3: Read Tools** | 4 days | All Category 2 tools; chunked streaming for large files | 500k-row read in <3s; all data types preserved round-trip |
| **Phase 4: Write + Structure** | 8 days | Categories 3, 4, 5; formula offset updating; impact reports | Mutation tests: insert/delete rows correctly adjusts formulas in adjacent sheets |
| **Phase 5: Calculation Engine** | 4 days | Tier 1 (`formulas` in-process) + Tier 2 (LibreOffice wrapper); `xls_recalculate.py` | Recalc agrees with Excel on test workbook to 5 decimal places |
| **Phase 6: Macro Safety** | 3 days | Category 9 tools using `oletools`; risk scanning | Correctly detects AutoOpen, Shell, CreateObject in test .xlsm |
| **Phase 7: Objects + Formatting** | 5 days | Categories 7, 8, 10 | Generated files open in Excel without repair prompts |
| **Phase 8: Token System** | 2 days | HMAC-SHA256 token generation, scoped validation, audit trail | Unauthorized destructive action returns exit code 4 |
| **Phase 9: E2E + Docs** | 3 days | Integration test suite (agent workflow simulation), CLI `--help`, design doc | New agent can execute: clone → analyze → modify → recalc → validate using only CLI |

**Total Estimated Duration: ~39 working days (8 weeks)**

### 7. Key Differences From the Draft

| Aspect | Draft | **Enhanced Blueprint** |
|:---|:---|:---|
| Macro analysis | `olefile` (basic OLE) | **`oletools/olevba`** (auto-exec detection, suspicious keyword scanning, XLM macro support) |
| Calculation | LibreOffice only | **Two-tier: `formulas`/`pycel` (fast, in-process) + LibreOffice (full fidelity fallback)** |
| Dependency graph | Custom, underspecified | **Built on `formulas` AST parsing with transitive closure and circular reference detection** |
| Tool count | 40 (some duplicated) | **53 tools**, de-duplicated, with new tools: `xls_dependency_report`, `xls_update_references`, `xls_get_workbook_metadata`, `xls_move_sheet`, `xls_write_cell`, `xls_set_number_format`, `xls_export_json` |
| Token scopes | 3 (`sheet:delete`, `range:delete`, `macro:remove`) | **7 scopes** including `sheet:rename`, `formula:convert`, `macro:inject`, `structure:modify` |
| Large file handling | Mentioned but unspecified | **Explicit chunked I/O via pandas streaming for >100k rows** |
| Audit trail | Not mentioned | **`.excel_agent_audit.jsonl`** — every destructive operation logged |
| `oletools` maintenance risk | Not addressed | **Abstraction layer** to allow swap if `oletools` becomes unmaintained |

---

## III. Validation Questions (Expanded)

Before proceeding to implementation, I recommend confirming:

1. **Macro Handling:** The strict read-only/container management approach for VBA is validated by research. The only modification we perform is injecting pre-extracted `.bin` files (following the XlsxWriter pattern). Do you approve this, including the `xls_inject_vba_project.py` tool for trusted macro templates?

2. **Two-Tier Calculation:** Do you accept the tiered approach? Tier 1 (`formulas`/`pycel`, in-process, ~50ms for 10k formulas) handles common cases without LibreOffice. Tier 2 (LibreOffice headless) is invoked only when Tier 1 encounters unsupported functions or the user explicitly requests full-fidelity recalc.

3. **Token Scopes:** The expanded 7-scope model adds `sheet:rename` (because formula side-effects), `formula:convert` (irreversible), and `macro:inject` (security-sensitive). Are these sufficient, or should `chart:delete` and `table:delete` also require tokens?

4. **`oletools` Risk Mitigation:** Given the maintenance concern, do you approve wrapping `oletools` behind a `MacroAnalyzer` interface so we can swap in `oledump.py` or a community fork without changing tool APIs?

5. **Audit Trail:** Is the `.excel_agent_audit.jsonl` file (appended on every destructive operation) acceptable, or should auditing be pluggable (file, database, webhook)?
