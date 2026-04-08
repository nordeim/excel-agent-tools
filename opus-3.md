Please conduct extensive research across multiple dimensions to validate, challenge, and enhance this blueprint. Let us search for the current state of the art in Python-Excel manipulation, macro handling, AI agent tool frameworks, and related areas.---

# Comprehensive Design Blueprint Review & Enhanced Architecture for `excel-agent-tools`

## I. Research Synthesis & Critical Findings

Extensive research across the Python-Excel ecosystem, AI agent frameworks, formula calculation engines, and macro security tools has validated core assumptions in the draft while revealing several critical enhancements and one significant architectural correction. 

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

---

# Phase 1: ANALYZE — Deep Multi-Dimensional Research & Validation

I'll conduct rigorous, independent research across every critical axis of this blueprint to validate claims, surface risks, and identify enhancements the draft may have missed. Let me begin.---

# Comprehensive Blueprint Validation Report

## Executive Summary

I have conducted independent research across all critical axes of your enhanced design blueprint. Below is a structured, multi-dimensional validation — confirming what is sound, challenging what needs refinement, and proposing targeted enhancements you haven't yet considered. My findings span **6 research axes** with **14 actionable findings**.

---

## I. Research Axis 1: Core Library — `openpyxl`

### ✅ VALIDATED

The latest available version on PyPI is openpyxl 3.1.5. The blueprint's choice is confirmed as correct.

openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files. It was born from lack of existing library to read/write natively from Python the Office Open XML format. This confirms full headless operation without Excel installed — essential for your AI-agent environment.

### ⚠️ CRITICAL SECURITY CONFIRMATION

By default openpyxl does not guard against quadratic blowup or billion laughs xml attacks. To guard against these attacks install defusedxml. Your blueprint correctly flags `defusedxml` as **mandatory**. This is validated.

### 🔶 NEW CONCERN: Python Version Support

The official docs state support for Python 3.6, 3.7, 3.8 and 3.9. However, Python 3.9 is already EOL (October 2025), and the recommendation is to plan migration to 3.12+ (or 3.14 for latest features). openpyxl's documentation appears stale on this front, while the package itself still works on modern Python. **Recommendation:** Pin your project to **Python ≥3.11** to align with Python 3.14.0, which was released with free-threaded builds moving from experimental to officially supported, and features including t-strings (PEP 750), deferred annotations (PEP 649), and sub-interpreters (PEP 734). This gives you access to modern features while maintaining compatibility with openpyxl. Use `3.12` as the floor.

---

## II. Research Axis 2: Macro Safety — `oletools`

### ✅ VALIDATED: Upgrade from `olefile` to `oletools` is Correct

oletools is a package of python tools to analyze Microsoft OLE2 files, mainly for malware analysis, forensics and debugging. It is based on the olefile parser. It also provides tools to analyze RTF files and files based on the OpenXML format such as MS Office 2007+ documents. oletools can detect, extract and analyse VBA macros, OLE objects, Excel 4 macros (XLM) and DDE links.

The toolset includes exactly what the blueprint's `MacroHandler` needs: olevba extracts and analyzes VBA Macro source code from MS Office documents (OLE and OpenXML).

### ✅ VALIDATED: Industry Adoption Confirms Maturity

oletools are used by a number of projects and online malware analysis services, including ACE, ADAPT, CAPE, CinCan, Cortex XSOAR (Palo Alto), Cuckoo Sandbox, FAME, Hybrid-analysis.com, InQuest Labs, IntelOwl, Joe Sandbox, Laika BOSS, MalwareBazaar, REMnux, Splunk add-on for MS O365 Email, Strelka, TheHive/Cortex, Viper, and probably VirusTotal. This is an extraordinary level of real-world validation for security tooling.

### ⚠️ CONFIRMED RISK: Maintenance Status

Further analysis of the maintenance status of oletools based on released PyPI versions cadence determined that its maintenance is Inactive. An important project maintenance signal is that it hasn't seen any new versions released to PyPI in the past 12 months, and could be considered as a discontinued project.

However: The python package oletools was scanned for known vulnerabilities and missing license, and no issues were found. Thus the package was deemed as safe to use.

The latest PyPI release is oletools version 0.60.2. The recommended Python version to run oletools is the latest Python 3.x (3.12 for now).

**Verdict:** The blueprint's decision to wrap `oletools` behind a `MacroAnalyzer` abstraction interface is **validated and essential**. The library is functionally sound and deeply battle-tested, but the maintenance gap creates bus-factor risk. The abstraction layer is the correct mitigation.

---

## III. Research Axis 3: Formula Calculation Engine — Tiered Architecture

### ✅ VALIDATED: `formulas` Library is the Strongest Tier 1 Candidate

formulas implements an interpreter for Excel formulas, which parses and compile Excel formulas expressions. Moreover, it compiles Excel workbooks to python and executes without using the Excel COM server. Hence, Excel is not needed.

Critically, it is **actively maintained**: The latest version is 1.3.4, released Mar 11, 2026, with the previous version 1.3.3 released Nov 4, 2025. This is a significant advantage over `pycel`.

The library offers exactly what the blueprint needs for dependency graph construction: It can plot the dependency graph that depicts relationships between Excel cells. And it supports circular references when you add `circular=True` to the finish method.

The CLI capabilities are a bonus for agent tool chaining: The formulas command-line interface works with spreadsheet models and accepts .xlsx, .ods, and .json inputs. A typical workflow starts by calculating a workbook. You can override input values directly from the command line and request specific cells to be rendered in the output.

Additionally, spreadsheet models can also be converted into a portable JSON representation, which is useful when the model needs to be versioned, inspected, or executed without the original workbook. This aligns perfectly with the blueprint's JSON-first philosophy.

### 🔶 IMPORTANT NUANCE: `pycel` vs `formulas`

Pycel can translate an Excel spreadsheet into executable python code which can be run independently of Excel. The python code is based on a graph and uses caching & lazy evaluation to ensure (relatively) fast execution.

Performance claim validated: Tested extensively on spreadsheets with 10 sheets & more than 10000 formulae. In that case calculation of the equations takes about 50ms and agrees with Excel up to 5 decimal places.

However, pycel has a critical limitation for headless environments. The *original* pycel fork requires an instance of Excel for initial compilation (i.e., the compilation needs to be run on Windows). The `stephenrauch` fork modernized this with a tokenizer of similar origin from the openpyxl library, removing the COM dependency. But the latest PyPI release is pycel-1.0b30 — still a beta.

**Architecture Decision Reinforcement:** The blueprint should specify:
- **Tier 1 PRIMARY: `formulas` 1.3.4** — Production-stable, actively maintained (release within last month), dependency graph built-in, CLI, JSON model export, circular reference support.
- **Tier 1 SECONDARY (optional): `pycel` via `stephenrauch` fork** — Useful for fast graph-based caching, but beta-quality. Wrap behind the same calculation interface.
- **Tier 2: LibreOffice Headless** — Full fidelity fallback, unchanged from draft.

### 🔶 CONSIDERATION: `xlcalculator` as Reference

xlcalculator is a Python library that reads MS Excel files and can translate the Excel functions into Python code and subsequently evaluate them. Essentially doing the Excel calculations without the need for Excel. xlcalculator is a modernization of the koala2 library.

It has a notable modern AST approach: Reimplemented evaluation engine to not generate Python code anymore, but build a proper AST from the AST nodes. Each AST node supports an eval() function that knows how to compute a result. This removes a lot of complexities around determining the evaluation context at code creation time.

However, it acknowledges a fundamental precision limitation: Further work will be required to keep numbers in-line with Excel throughout different transformations. This requires a low-level implementation of a numeric datatype (C or C++, Cython?) to replicate its behaviour. Python built-in numeric types don't replicate behaviours appropriately.

**Verdict:** Keep `xlcalculator` as reference material for AST design patterns, but do not add it as a runtime dependency. `formulas` is superior in maturity, maintenance, and feature coverage.

---

## IV. Research Axis 4: Governance Token Architecture — HMAC-SHA256

### ✅ VALIDATED: HMAC-SHA256 is Industry Standard for Agent Tool Gating

While new authentication standards like OAuth 2.1 and JWT dominate headlines, HMAC (Hash-based Message Authentication Code) remains the most practical, efficient, and reliable way to verify message integrity between trusted systems.

HMAC is a cryptographic method that uses a secret key and a hashing algorithm (like SHA256) to verify that the message hasn't been changed in transit, and the message came from a trusted sender that knows the shared secret.

The blueprint's choice is validated by current agent framework implementations. Production AI agent platforms are using exactly this pattern: AES-256-GCM at rest, HMAC-hashed access tokens with per-agent and per-fleet permissioning — each agent or group gets its own token with scoped access.

### 🔶 ENHANCEMENT: Add TTL, Nonce, and Timestamp

Current agent approval frameworks implement stronger patterns than the blueprint specifies. Challenge/response (nonce + HMAC token + request hash) + short TTL is the standard. The phantom token pattern implements: Each proxied request gets a UUID v4 request ID and an HMAC-SHA256 signature computed over a canonical string. The server verifies signatures on ingestion with constant-time comparison, rejects stale timestamps (more than 5 minutes in the future, more than 24 hours in the past).

Best practices per current security guidance: Include a timestamp in signed messages. This prevents replay attacks — attackers can't reuse old requests. And critically: Use constant-time comparison. In Python, use hmac.compare_digest.

**Enhancement to the `ApprovalTokenMgr` spec:**

```python
@dataclass
class ApprovalToken:
    scope: str              # e.g., "sheet:delete"
    target_file_hash: str   # SHA-256 of target workbook
    nonce: str              # UUID4, one-time use
    issued_at: float        # Unix timestamp
    ttl_seconds: int        # Default: 300 (5 min)
    signature: str          # HMAC-SHA256(secret, scope|hash|nonce|issued_at|ttl)
```

The token MUST be:
1. **Scoped** to a specific file hash (prevents reuse across workbooks)
2. **Time-limited** with TTL (prevents replay attacks)
3. **Single-use** via nonce tracking (prevents re-submission)
4. **Validated with `hmac.compare_digest()`** (prevents timing attacks)

This is a significant upgrade over the blueprint's current unspecified token internals.

---

## V. Research Axis 5: Python Ecosystem Context (2026)

### 🔶 NEW FINDING: Supply Chain Security

Python 3.14 ships with Sigstore-only signing (PEP 761) replacing PGP and includes SBOMs. Python ecosystem modernizing supply-chain security. PyPI continuing malware detection improvements.

The recent PyPI attack is relevant context: PyPI cloud token malware campaign disclosed. 20 packages masquerading as time utilities stole AWS, Alibaba Cloud, and Tencent Cloud credentials. 14,100+ downloads before removal.

**Recommendation:** The `excel-agent-tools` project should:
1. Pin all dependency versions with hashes in `requirements.txt`
2. Publish with Sigstore attestations
3. Include SBOM generation in the build pipeline

### 🔶 NEW FINDING: Python 3.14 Template Strings (PEP 750)

Steering Council approves PEP 750 (template strings). New t"..." syntax enables custom processing of interpolated strings for safer SQL, HTML, and DSL generation.

This could be relevant for the formula string construction tools (`xls_set_formula.py`), where injection of malicious formula content is a real risk. Template strings would allow **structured formula building** that prevents injection. Mark this as a Phase 10 enhancement opportunity.

---

## VI. Research Axis 6: Agent Approval Pattern Validation

### ✅ VALIDATED: Denial-with-guidance Pattern

Modern agent approval frameworks implement a pattern directly relevant to the blueprint's impact report system: When an async exec approval is denied, the platform prevents the agent from reusing output from any earlier run of the same command in the session. The denial reason is passed with explicit guidance that no command output is available, which stops the agent from claiming there is new output or repeating the denied command.

**Recommendation:** When `xls_delete_rows.py` returns `exit_code=1` due to broken references, the JSON response should include **prescriptive guidance** that the agent cannot ignore:

```json
{
    "status": "denied",
    "exit_code": 1,
    "denial_reason": "Operation would break 7 formula references across 3 sheets",
    "guidance": "Run xls_update_references.py --target='Sheet1!A5:A10' before retrying",
    "stale_output_warning": "Do not proceed with cached data from prior reads of affected cells"
}
```

---

## VII. Consolidated Findings Matrix

| # | Finding | Category | Severity | Blueprint Status | Action Required |
|---|:---|:---|:---|:---|:---|
| 1 | openpyxl 3.1.5 confirmed correct | Core I/O | — | ✅ Validated | None |
| 2 | `defusedxml` mandatory | Security | Critical | ✅ Validated | None |
| 3 | Python version floor should be ≥3.12 | Platform | Medium | 🔶 Missing | Add to blueprint |
| 4 | `oletools` upgrade from `olefile` correct | Macro Safety | High | ✅ Validated | None |
| 5 | `oletools` maintenance risk confirmed | Dependency | Medium | ✅ Mitigated by abstraction | Proceed as designed |
| 6 | `formulas` 1.3.4 is best Tier 1 engine | Calculation | High | ✅ Validated | Promote over `pycel` |
| 7 | `pycel` is beta (1.0b30), use cautiously | Calculation | Medium | 🔶 Needs clarification | Demote to optional |
| 8 | `formulas` JSON model export aligns with JSON-first | Architecture | Medium | ✅ Bonus capability | Leverage in tools |
| 9 | `xlcalculator` AST pattern useful as reference only | Calculation | Low | ✅ Correctly excluded | No runtime dependency |
| 10 | HMAC-SHA256 validated for agent tool gating | Security | High | ✅ Validated | None |
| 11 | Token needs TTL + nonce + timestamp + constant-time compare | Security | **Critical** | 🔶 Underspecified | Enhance token spec |
| 12 | Supply chain security (Sigstore, SBOM, pinned hashes) | Security | Medium | 🔶 Missing | Add to Phase 9 |
| 13 | PEP 750 template strings for formula injection prevention | Security | Low | 🔶 Future opportunity | Phase 10 enhancement |
| 14 | Denial-with-prescriptive-guidance in impact reports | Agent UX | Medium | 🔶 Underspecified | Enhance JSON schema |

---

## VIII. Answers to Your 5 Validation Questions

### Q1: Macro Handling — Read-only/container management + `xls_inject_vba_project.py`?

**Approved with one constraint.** The approach is validated by research — VBA binary containers are opaque OLE structures that cannot be feasibly constructed from scratch. The XlsxWriter pattern (extract `.bin`, re-inject) is the only safe approach. oletools can detect, extract and analyse VBA macros, OLE objects, Excel 4 macros (XLM) and DDE links — but note that injection MUST be preceded by a `scan_risk()` call on the `.bin` being injected. **Never inject an unscanned macro binary.** Add this as a hard pre-condition in the tool's execution flow.

### Q2: Two-Tier Calculation — Accept the tiered approach?

**Approved with a hierarchy clarification.** Based on research, the tier should be:
- **Tier 1 PRIMARY: `formulas` 1.3.4** — Released March 2026, production-stable, with workbook compilation to Python and execution without Excel COM server. Excel is not needed.
- **Tier 1 SECONDARY (optional): `pycel`** — Still at beta version 1.0b30. Useful for graph caching but should be an opt-in dependency behind a feature flag.
- **Tier 2: LibreOffice Headless** — Unchanged. Invoked when `formulas` encounters an unsupported function or when the user explicitly requests full-fidelity recalc.

### Q3: Token Scopes — Are 7 sufficient, or should `chart:delete` and `table:delete` also require tokens?

**7 is sufficient for v1.0.0, but enhance the token specification.** `chart:delete` and `table:delete` are non-destructive to formula integrity (they don't break cell references), so they don't warrant governance tokens in v1.0. However, the token itself needs the enhanced specification from Finding #11: TTL, nonce, target file hash binding, and `hmac.compare_digest()` validation — include a timestamp to prevent replay attacks, and use constant-time comparison — in Python, use hmac.compare_digest.

### Q4: `oletools` Risk Mitigation — Approve `MacroAnalyzer` abstraction?

**Strongly approved.** The maintenance status is Inactive and it hasn't seen any new versions in the past 12 months. Yet it is used by Cortex XSOAR (Palo Alto), Cuckoo Sandbox, Joe Sandbox, REMnux, Splunk, TheHive/Cortex, and many others. This creates a paradox: battle-tested but dormant. The abstraction layer is the correct architectural hedge. Define the interface as:

```python
class MacroAnalyzer(Protocol):
    def has_macros(self, path: Path) -> bool: ...
    def extract_modules(self, path: Path) -> list[MacroModule]: ...
    def detect_auto_exec(self, path: Path) -> list[AutoExecTrigger]: ...
    def detect_suspicious(self, path: Path) -> list[SuspiciousKeyword]: ...
    def scan_risk(self, path: Path) -> RiskReport: ...
```

The `OletoolsMacroAnalyzer` implements this today. If a maintained fork appears, swap the implementation — zero tool API changes.

### Q5: Audit Trail — `.excel_agent_audit.jsonl` acceptable, or should auditing be pluggable?

**Make it pluggable from day one, with `.jsonl` as the default backend.** The research shows that production agent frameworks consistently implement auditing where every secret access, every API call, every auth attempt is logged with actor, timestamp, and IP. The phantom token pattern also emphasizes every individual API call being independently auditable — if you need to investigate what happened at 14:23:07 UTC, you have a signed record.

Design the audit layer as:

```python
class AuditBackend(Protocol):
    def log_event(self, event: AuditEvent) -> None: ...

class JsonlAuditBackend(AuditBackend):
    """Default: append to .excel_agent_audit.jsonl"""

class WebhookAuditBackend(AuditBackend):
    """Optional: POST to configurable endpoint"""

class CompositeAuditBackend(AuditBackend):
    """Fan-out to multiple backends simultaneously"""
```

This adds ~2 hours to Phase 8, but prevents painful retrofitting when enterprise users need webhook/SIEM integration.

---

## IX. Final Recommendation

**The blueprint is architecturally sound and research-validated.** The 14 findings above are refinements, not rewrites. The most critical actions before proceeding to implementation are:

1. **Pin Python ≥3.12** as minimum version
2. **Enhance the `ApprovalToken` specification** with TTL + nonce + file hash binding + constant-time comparison
3. **Promote `formulas` to sole Tier 1 engine**, demote `pycel` to optional
4. **Add denial-with-prescriptive-guidance** to the JSON response schema for all governance-gated tools
5. **Make the audit backend pluggable** from day one

**I recommend proceeding to Phase 3 (VALIDATE) — explicit confirmation from you on these 5 adjustments — then moving directly to implementation.**
