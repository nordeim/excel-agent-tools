# Master Execution Plan for `excel-agent-tools`

## Executive Summary

This Master Execution Plan provides a meticulous, phase-by-phase implementation roadmap for the `excel-agent-tools` project. Each phase includes detailed file specifications, feature requirements, interface definitions, and verification checklists. The plan follows a dependency-aware progression, ensuring foundational components are battle-tested before dependent tools are built.

**Total Phases:** 10  
**Total Files:** 78 (53 tools + 25 support files)  
**Estimated Duration:** 8 weeks  
**Quality Gates:** Each phase requires 100% checklist completion before proceeding

---

## Phase Structure Template

Each phase follows this standardized structure:
- **Phase Objective:** High-level goal
- **Dependencies:** Prerequisites from previous phases
- **Duration:** Estimated working days
- **Files:** Detailed specifications
- **Gate Criteria:** Must-pass requirements before next phase
- **Risk Mitigations:** Identified risks and countermeasures

---

## PHASE 1: Foundation & Infrastructure

**Objective:** Establish core library, project structure, and development environment  
**Dependencies:** None (genesis phase)  
**Duration:** 5 days  
**Gate Criteria:** Core library loads/saves workbooks, file locking works cross-platform

### Files to Create

#### 1.1 `pyproject.toml`
**Description:** Project metadata, dependencies, and build configuration  
**Features & Interfaces:**
- Python 3.9+ requirement
- Dependencies: `openpyxl>=3.1.5`, `defusedxml>=0.7.1`, `click>=8.1.0`
- Optional dependencies for each tier: `[formulas]`, `[oletools]`, `[pandas]`
- Entry points for all 53 CLI tools
- Development dependencies: `pytest>=7.0`, `hypothesis>=6.0`, `black`, `mypy`, `ruff`

**Checklist:**
- [ ] Project name: `excel-agent-tools`
- [ ] Version: `1.0.0`
- [ ] License: MIT
- [ ] Python requirement: `>=3.9,<4.0`
- [ ] Core dependencies pinned with minimum versions
- [ ] Optional dependency groups defined
- [ ] Build backend: `setuptools>=61.0`
- [ ] Package discovery configured for `src/` layout
- [ ] CLI entry points mapped for future tools

#### 1.2 `src/excel_agent_tools/__init__.py`
**Description:** Package initialization and version export  
**Features & Interfaces:**
```python
__version__ = "1.0.0"
__all__ = ["ExcelAgent", "DependencyTracker", "MacroHandler", "RangeSerializer"]

from .core import ExcelAgent
from .dependencies import DependencyTracker
from .macros import MacroHandler
from .serializers import RangeSerializer
```

**Checklist:**
- [ ] Version string matches pyproject.toml
- [ ] `__all__` exports public API
- [ ] Lazy imports for heavy dependencies
- [ ] Package docstring with project description

#### 1.3 `src/excel_agent_tools/core.py`
**Description:** Core `ExcelAgent` context manager and file operations  
**Features & Interfaces:**
```python
class ExcelAgent:
    def __init__(self, path: Path, *, mode: str = "r", keep_vba: bool = True)
    def __enter__(self) -> 'ExcelAgent'
    def __exit__(self, exc_type, exc_val, exc_tb)
    def _acquire_lock(self) -> None  # OS-specific file locking
    def _release_lock(self) -> None
    def _compute_version_hash(self) -> str  # SHA256 of structure+formulas
    def _verify_no_concurrent_modification(self) -> None
    @property
    def workbook(self) -> Workbook
    def save(self, path: Optional[Path] = None) -> None
```

**Checklist:**
- [ ] Cross-platform file locking (fcntl on Unix, msvcrt on Windows)
- [ ] Atomic lock acquisition with timeout (default 5s)
- [ ] Context manager properly handles exceptions
- [ ] Version hash excludes values for performance
- [ ] `keep_vba=True` preservation verified
- [ ] Concurrent modification detection implemented
- [ ] Comprehensive docstrings with usage examples
- [ ] Type hints for all methods
- [ ] 100% unit test coverage

#### 1.4 `src/excel_agent_tools/serializers.py`
**Description:** Range notation converters (A1, R1C1, Table, Named Range)  
**Features & Interfaces:**
```python
class RangeSerializer:
    @staticmethod
    def parse_a1(notation: str) -> Tuple[int, int, int, int]
    @staticmethod
    def parse_r1c1(notation: str) -> Tuple[int, int, int, int]
    @staticmethod
    def parse_table(notation: str, workbook: Workbook) -> Tuple[int, int, int, int]
    @staticmethod
    def parse_named_range(name: str, workbook: Workbook) -> Tuple[int, int, int, int]
    @staticmethod
    def normalize(notation: str, workbook: Optional[Workbook] = None) -> dict
    @staticmethod
    def to_a1(min_col: int, min_row: int, max_col: int, max_row: int) -> str
```

**Checklist:**
- [ ] A1 notation: handles single cells ("A1") and ranges ("A1:C10")
- [ ] R1C1 notation: supports relative references ("R[-1]C[2]")
- [ ] Table references: parses "Table1[Column]", "Table1[#Headers]", etc.
- [ ] Named ranges: resolves both workbook and sheet-scoped names
- [ ] Coordinate output: `{"min_row": 1, "min_col": 1, "max_row": 10, "max_col": 3}`
- [ ] Sheet qualifiers: handles "'Sheet Name'!A1" with proper escaping
- [ ] Error handling: raises `RangeParseError` with clear messages
- [ ] Edge cases: full column ("A:A"), full row ("1:1"), entire sheet ("*")
- [ ] Performance: <1ms for typical ranges

#### 1.5 `src/excel_agent_tools/exceptions.py`
**Description:** Custom exception hierarchy  
**Features & Interfaces:**
```python
class ExcelAgentError(Exception): ...
class FileLockTimeout(ExcelAgentError): ...
class ConcurrentModificationError(ExcelAgentError): ...
class RangeParseError(ExcelAgentError): ...
class TokenValidationError(ExcelAgentError): ...
class DependencyError(ExcelAgentError): ...
class MacroSecurityError(ExcelAgentError): ...
```

**Checklist:**
- [ ] Base exception class with structured error data
- [ ] Each exception includes error code for CLI exit mapping
- [ ] JSON-serializable error details
- [ ] Stack trace suppression for user-facing errors

#### 1.6 `src/excel_agent_tools/constants.py`
**Description:** Project-wide constants and configuration  
**Features & Interfaces:**
```python
# Exit codes
EXIT_SUCCESS = 0
EXIT_VALIDATION_ERROR = 1
EXIT_FILE_NOT_FOUND = 2
EXIT_LOCK_CONTENTION = 3
EXIT_PERMISSION_DENIED = 4
EXIT_INTERNAL_ERROR = 5

# Limits
MAX_ROWS = 1_048_576  # Excel 2007+ limit
MAX_COLS = 16_384     # XFD column
CHUNK_SIZE = 10_000   # Rows per chunk for streaming

# Paths
DEFAULT_WORK_DIR = Path("/work")
AUDIT_LOG = Path(".excel_agent_audit.jsonl")
```

**Checklist:**
- [ ] Exit codes documented with scenarios
- [ ] Excel limits match official specifications
- [ ] Configurable paths via environment variables
- [ ] Token expiry duration (default 1 hour)
- [ ] Logging levels defined

#### 1.7 `tests/conftest.py`
**Description:** pytest fixtures and test utilities  
**Features & Interfaces:**
```python
@pytest.fixture
def sample_workbook() -> Path: ...

@pytest.fixture
def workbook_with_formulas() -> Path: ...

@pytest.fixture
def workbook_with_macros() -> Path: ...

@pytest.fixture
def temp_work_dir(tmp_path) -> Path: ...
```

**Checklist:**
- [ ] Sample .xlsx files in `tests/fixtures/`
- [ ] Sample .xlsm with benign macro
- [ ] Temporary directory cleanup
- [ ] Mock LibreOffice for CI environments
- [ ] Hypothesis strategies for range generation

#### 1.8 `tests/test_core.py`
**Description:** Unit tests for ExcelAgent  
**Checklist:**
- [ ] Lock acquisition and release
- [ ] Concurrent access blocking
- [ ] Version hash stability
- [ ] VBA preservation through save cycle
- [ ] Exception handling in context manager
- [ ] Cross-platform compatibility

#### 1.9 `tests/test_serializers.py`
**Description:** Unit tests for RangeSerializer  
**Checklist:**
- [ ] All notation formats with valid inputs
- [ ] Invalid notation error handling
- [ ] Sheet name escaping edge cases
- [ ] Performance benchmark (<1ms)
- [ ] Round-trip conversion (A1 → coords → A1)

#### 1.10 `.github/workflows/ci.yml`
**Description:** GitHub Actions CI pipeline  
**Features & Interfaces:**
```yaml
name: CI
on: [push, pull_request]
jobs:
  test:
    strategy:
      matrix:
        os: [ubuntu-latest, windows-latest, macos-latest]
        python: ["3.9", "3.10", "3.11", "3.12"]
```

**Checklist:**
- [ ] Multi-OS testing matrix
- [ ] Python version matrix (3.9-3.12)
- [ ] LibreOffice installation step
- [ ] pytest with coverage reporting
- [ ] mypy type checking
- [ ] ruff linting
- [ ] Security scanning with bandit
- [ ] Artifact upload for test files

---

## PHASE 2: Dependency Tracking Engine

**Objective:** Implement formula dependency graph for safe mutations  
**Dependencies:** Phase 1 completed  
**Duration:** 5 days  
**Gate Criteria:** Dependency graph correctly identifies impact of cell deletion

### Files to Create

#### 2.1 `src/excel_agent_tools/dependencies.py`
**Description:** Formula dependency tracker with graph algorithms  
**Features & Interfaces:**
```python
class DependencyTracker:
    def __init__(self, workbook: Workbook)
    def build_graph(self, sheets: Optional[List[str]] = None) -> None
    def find_dependents(self, target: str) -> Set[str]
    def find_precedents(self, target: str) -> Set[str]
    def impact_report(self, target_range: str, action: str) -> dict
    def detect_circular_references(self) -> List[List[str]]
    def topological_sort(self) -> List[str]
    def _parse_formula(self, formula: str) -> List[str]  # Extract cell refs
    def _resolve_reference(self, ref: str, source_sheet: str) -> str
```

**Checklist:**
- [ ] Integration with `formulas` library for AST parsing
- [ ] Fallback to openpyxl.formula.Tokenizer if formulas unavailable
- [ ] Cross-sheet reference resolution
- [ ] Named range expansion
- [ ] Table reference resolution
- [ ] Circular reference detection via DFS
- [ ] Transitive dependency calculation
- [ ] Impact report includes severity levels
- [ ] Performance: <100ms for 10k formulas
- [ ] Memory efficient for large workbooks

#### 2.2 `src/excel_agent_tools/formula_parser.py`
**Description:** Formula tokenization and reference extraction  
**Features & Interfaces:**
```python
class FormulaParser:
    def __init__(self, formula: str)
    def get_cell_references(self) -> List[CellReference]
    def get_range_references(self) -> List[RangeReference]
    def get_named_references(self) -> List[str]
    def get_function_calls(self) -> List[FunctionCall]
    def is_array_formula(self) -> bool
    def has_volatile_functions(self) -> bool  # NOW(), RAND(), etc.
```

**Checklist:**
- [ ] Handles all Excel operators (+, -, *, /, ^, &, =, <>, etc.)
- [ ] Supports 400+ Excel functions
- [ ] Array formula detection ({=...})
- [ ] Structured references (Table1[@Column])
- [ ] 3D references (Sheet1:Sheet3!A1)
- [ ] External workbook references ([Book1.xlsx]Sheet1!A1)
- [ ] R1C1 notation support
- [ ] Error handling for malformed formulas

#### 2.3 `src/excel_agent_tools/graph_algorithms.py`
**Description:** Graph utilities for dependency analysis  
**Features & Interfaces:**
```python
def build_adjacency_list(edges: List[Tuple[str, str]]) -> Dict[str, Set[str]]
def find_strongly_connected_components(graph: dict) -> List[Set[str]]
def find_cycles(graph: dict) -> List[List[str]]
def transitive_closure(graph: dict, node: str) -> Set[str]
def topological_sort(graph: dict) -> Optional[List[str]]
```

**Checklist:**
- [ ] Tarjan's algorithm for SCC
- [ ] Cycle detection with path reconstruction
- [ ] BFS/DFS implementations
- [ ] Transitive closure optimization
- [ ] Handles graphs with 100k+ nodes

#### 2.4 `tests/test_dependencies.py`
**Description:** Unit tests for dependency tracking  
**Checklist:**
- [ ] Simple formula dependencies (A1 = B1 + C1)
- [ ] Cross-sheet dependencies
- [ ] Circular reference detection
- [ ] Named range dependencies
- [ ] Table reference dependencies
- [ ] Impact report accuracy
- [ ] Performance benchmarks
- [ ] Memory usage tests

#### 2.5 `tools/xls_dependency_report.py`
**Description:** CLI tool to export dependency graph  
**Features & Interfaces:**
```python
# CLI: xls_dependency_report.py workbook.xlsx --output-format json
# Output: {"dependencies": {"Sheet1!A1": ["Sheet1!B1", "Sheet1!C1"], ...}}
```

**Checklist:**
- [ ] Argument parsing with argparse
- [ ] JSON output format (default)
- [ ] DOT format for Graphviz visualization
- [ ] CSV format for analysis
- [ ] Filter by sheet option
- [ ] Include/exclude external references
- [ ] Performance metrics in output
- [ ] Proper exit codes

---

## PHASE 3: Token System & Governance

**Objective:** Implement HMAC approval tokens and audit logging  
**Dependencies:** Phase 1 completed  
**Duration:** 2 days  
**Gate Criteria:** Destructive operations fail without valid token

### Files to Create

#### 3.1 `src/excel_agent_tools/tokens.py`
**Description:** HMAC-SHA256 token generation and validation  
**Features & Interfaces:**
```python
class ApprovalToken:
    def __init__(self, secret_key: Optional[str] = None)
    def generate(self, scope: str, resource: str, expiry_minutes: int = 60) -> str
    def validate(self, token: str, required_scope: str, resource: str) -> bool
    def decode(self, token: str) -> dict
    
class TokenScope(Enum):
    SHEET_DELETE = "sheet:delete"
    SHEET_RENAME = "sheet:rename"
    RANGE_DELETE = "range:delete"
    FORMULA_CONVERT = "formula:convert"
    MACRO_REMOVE = "macro:remove"
    MACRO_INJECT = "macro:inject"
    STRUCTURE_MODIFY = "structure:modify"
```

**Checklist:**
- [ ] HMAC-SHA256 signing with secret key
- [ ] Base64URL encoding for tokens
- [ ] Expiry timestamp validation
- [ ] Scope hierarchy (structure:modify implies sheet:delete)
- [ ] Resource-specific tokens (file path hash)
- [ ] Constant-time comparison to prevent timing attacks
- [ ] Token version field for future upgrades
- [ ] JSON payload with claims

#### 3.2 `src/excel_agent_tools/audit.py`
**Description:** Audit trail logging for all operations  
**Features & Interfaces:**
```python
class AuditLogger:
    def __init__(self, log_path: Path = AUDIT_LOG)
    def log_operation(self, operation: str, file: Path, user: str, token: str, impact: dict)
    def log_error(self, operation: str, error: Exception)
    def get_recent_operations(self, limit: int = 100) -> List[dict]
    def export_audit_report(self, start_date: datetime, end_date: datetime) -> dict
```

**Checklist:**
- [ ] JSON Lines format (.jsonl)
- [ ] Atomic append operations
- [ ] Timestamp with microsecond precision
- [ ] User identification (from env or token)
- [ ] Operation classification
- [ ] Impact metrics recording
- [ ] Log rotation at 100MB
- [ ] Compression of old logs
- [ ] Query interface for analysis

#### 3.3 `tools/xls_approve_token.py`
**Description:** CLI tool to generate approval tokens  
**Features & Interfaces:**
```bash
xls_approve_token.py --scope sheet:delete --resource workbook.xlsx --expiry 60
# Output: {"token": "eyJ0eXAiOiJKV1QiLCJhbGc...", "expires_at": "2024-01-01T12:00:00Z"}
```

**Checklist:**
- [ ] Secret key from environment variable
- [ ] All scope types supported
- [ ] Resource path validation
- [ ] Expiry in minutes (default 60)
- [ ] JSON output with metadata
- [ ] Option to save to file
- [ ] Verbose mode with token details

#### 3.4 `tests/test_tokens.py`
**Description:** Unit tests for token system  
**Checklist:**
- [ ] Token generation and validation
- [ ] Expiry enforcement
- [ ] Scope validation
- [ ] Invalid token rejection
- [ ] Tampered token detection
- [ ] Resource binding verification

---

## PHASE 4: Core Read Operations

**Objective:** Implement all data extraction tools  
**Dependencies:** Phases 1-3 completed  
**Duration:** 4 days  
**Gate Criteria:** Can read 500k rows in <3 seconds

### Files to Create

#### 4.1 `src/excel_agent_tools/readers.py`
**Description:** High-performance data reading with chunking  
**Features & Interfaces:**
```python
class WorkbookReader:
    def __init__(self, workbook: Workbook)
    def read_range(self, range_spec: str, data_only: bool = True) -> List[List[Any]]
    def read_range_chunked(self, range_spec: str, chunk_size: int = 10000) -> Iterator[List[List[Any]]]
    def get_sheet_dimensions(self, sheet_name: str) -> dict
    def get_all_formulas(self, sheet_name: Optional[str] = None) -> Dict[str, str]
    def get_all_values(self, sheet_name: Optional[str] = None) -> Dict[str, Any]
    def get_cell_metadata(self, cell_ref: str) -> dict
```

**Checklist:**
- [ ] Type inference for values (date, time, boolean, number, string)
- [ ] Formula preservation option
- [ ] Memory-efficient chunking for large ranges
- [ ] Merged cell handling
- [ ] Hidden row/column detection
- [ ] Style extraction optional
- [ ] Performance: 100k rows/second minimum

#### 4.2 `tools/xls_read_range.py`
**Description:** Extract data from specified range  
**Features & Interfaces:**
```bash
xls_read_range.py workbook.xlsx "Sheet1!A1:C10" --format json
# Output: {"data": [[val1, val2, val3], ...], "dimensions": {"rows": 10, "cols": 3}}
```

**Checklist:**
- [ ] Range notation parsing (A1, Table, Named)
- [ ] JSON output (default)
- [ ] CSV output option
- [ ] JSON Lines for large data
- [ ] Include formulas option
- [ ] Include styles option
- [ ] Null handling configuration
- [ ] Date format specification

#### 4.3 `tools/xls_get_sheet_names.py`
**Description:** List all sheets with metadata  
**Features & Interfaces:**
```bash
xls_get_sheet_names.py workbook.xlsx
# Output: {"sheets": [{"index": 0, "name": "Sheet1", "visible": true, "type": "worksheet"}]}
```

**Checklist:**
- [ ] Sheet index (0-based)
- [ ] Sheet name with quotes if needed
- [ ] Visibility state (visible/hidden/veryHidden)
- [ ] Sheet type (worksheet/chart/dialog/macro)
- [ ] Cell count estimation
- [ ] Protection status

#### 4.4 `tools/xls_get_defined_names.py`
**Description:** Extract all named ranges  
**Checklist:**
- [ ] Global names
- [ ] Sheet-scoped names
- [ ] Name, reference, scope, comment
- [ ] Built-in names (_xlnm.*)
- [ ] Dynamic names detection

#### 4.5 `tools/xls_get_table_info.py`
**Description:** List Excel Tables (ListObjects)  
**Checklist:**
- [ ] Table name and display name
- [ ] Range reference
- [ ] Column names and data types
- [ ] Total row presence
- [ ] Style name
- [ ] Auto-filter state

#### 4.6 `tools/xls_get_cell_style.py`
**Description:** Extract cell formatting  
**Checklist:**
- [ ] Font (name, size, bold, italic, color)
- [ ] Fill (pattern, fgColor, bgColor)
- [ ] Border (style, color for each side)
- [ ] Alignment (horizontal, vertical, wrap, indent)
- [ ] Number format code
- [ ] Protection settings

#### 4.7 `tools/xls_get_formula.py`
**Description:** Get formula from specific cell  
**Checklist:**
- [ ] Formula string or null
- [ ] Array formula detection
- [ ] Parsed references
- [ ] R1C1 conversion option

#### 4.8 `tools/xls_get_workbook_metadata.py`
**Description:** Workbook-level statistics  
**Checklist:**
- [ ] Sheet count
- [ ] Total formulas
- [ ] Total named ranges
- [ ] Table count
- [ ] Chart count
- [ ] VBA presence
- [ ] File size
- [ ] Creation/modification dates

#### 4.9 `tests/test_readers.py`
**Description:** Unit tests for all read operations  
**Checklist:**
- [ ] Small file reads
- [ ] Large file streaming
- [ ] Type inference accuracy
- [ ] Formula extraction
- [ ] Performance benchmarks

---

## PHASE 5: Core Write Operations

**Objective:** Implement data writing and workbook creation tools  
**Dependencies:** Phase 4 completed  
**Duration:** 3 days  
**Gate Criteria:** Round-trip data integrity verified

### Files to Create

#### 5.1 `src/excel_agent_tools/writers.py`
**Description:** Data writing with type preservation  
**Features & Interfaces:**
```python
class WorkbookWriter:
    def __init__(self, workbook: Workbook)
    def write_range(self, range_spec: str, data: List[List[Any]], preserve_formulas: bool = False)
    def write_cell(self, cell_ref: str, value: Any, data_type: Optional[str] = None)
    def set_formula(self, cell_ref: str, formula: str, validate: bool = True)
    def clear_range(self, range_spec: str, clear_formats: bool = False)
    def copy_range(self, source: str, target: str, include_formats: bool = True)
```

**Checklist:**
- [ ] Type inference from Python types
- [ ] Explicit type override option
- [ ] Formula syntax validation
- [ ] Date/time handling with timezone
- [ ] Bulk write optimization
- [ ] Transaction-like behavior (all-or-nothing)
- [ ] Memory efficiency for large writes

#### 5.2 `tools/xls_create_new.py`
**Description:** Create blank workbook  
**Features & Interfaces:**
```bash
xls_create_new.py output.xlsx --sheets "Data" "Summary" "Charts"
# Output: {"status": "success", "file": "output.xlsx", "sheets": 3}
```

**Checklist:**
- [ ] Custom sheet names
- [ ] Sheet count option
- [ ] Template selection
- [ ] Default styles
- [ ] Locale settings

#### 5.3 `tools/xls_create_from_template.py`
**Description:** Create from template with substitutions  
**Checklist:**
- [ ] Template validation
- [ ] Variable substitution {{var}}
- [ ] Named range preservation
- [ ] Macro preservation for .xltm
- [ ] Style inheritance

#### 5.4 `tools/xls_write_range.py`
**Description:** Write data to range  
**Checklist:**
- [ ] JSON input from stdin or file
- [ ] CSV input support
- [ ] Type hints in input
- [ ] Append mode
- [ ] Overwrite confirmation
- [ ] Formula preservation option

#### 5.5 `tools/xls_write_cell.py`
**Description:** Single cell write  
**Checklist:**
- [ ] Value from command line
- [ ] Type specification
- [ ] Formula input
- [ ] Style preservation

#### 5.6 `tools/xls_clone_workbook.py`
**Description:** Safe copy for editing  
**Checklist:**
- [ ] Source hash verification
- [ ] Unique output name generation
- [ ] Timestamp in filename
- [ ] Work directory validation
- [ ] VBA preservation
- [ ] External link handling

---

## PHASE 6: Structural Operations

**Objective:** Implement sheet/row/column manipulation with dependency awareness  
**Dependencies:** Phases 1-5 completed  
**Duration:** 8 days  
**Gate Criteria:** Formula references update correctly after structural changes

### Files to Create

#### 6.1 `src/excel_agent_tools/structural.py`
**Description:** Sheet and range structure modifications  
**Features & Interfaces:**
```python
class StructuralEditor:
    def __init__(self, workbook: Workbook, dependency_tracker: DependencyTracker)
    def add_sheet(self, name: str, position: Optional[int] = None) -> Worksheet
    def delete_sheet(self, name: str, token: str) -> dict  # Returns impact
    def rename_sheet(self, old_name: str, new_name: str, token: str) -> dict
    def insert_rows(self, sheet: str, before_row: int, count: int) -> dict
    def delete_rows(self, sheet: str, start_row: int, count: int, token: str) -> dict
    def insert_columns(self, sheet: str, before_col: int, count: int) -> dict
    def delete_columns(self, sheet: str, start_col: int, count: int, token: str) -> dict
    def move_sheet(self, sheet_name: str, new_position: int) -> None
    def _update_formula_references(self, modifications: List[RefUpdate]) -> None
```

**Checklist:**
- [ ] Token validation for destructive operations
- [ ] Pre-flight impact analysis
- [ ] Formula reference updating
- [ ] Named range adjustment
- [ ] Table range expansion/contraction
- [ ] Conditional formatting range updates
- [ ] Chart data range updates
- [ ] Validation range updates
- [ ] Audit logging for all changes

#### 6.2 `tools/xls_add_sheet.py`
**Description:** Add new worksheet  
**Checklist:**
- [ ] Name validation (no special chars)
- [ ] Position specification (before/after)
- [ ] Copy from existing option
- [ ] Tab color setting

#### 6.3 `tools/xls_delete_sheet.py` ⚠️
**Description:** Delete worksheet with dependency check  
**Checklist:**
- [ ] Token validation
- [ ] Impact report generation
- [ ] Cross-sheet reference scan
- [ ] Force flag with double-token
- [ ] Backup before deletion

#### 6.4 `tools/xls_rename_sheet.py` ⚠️
**Description:** Rename sheet and update references  
**Checklist:**
- [ ] Token validation
- [ ] Name collision check
- [ ] Formula reference updates
- [ ] Named range updates
- [ ] External link updates

#### 6.5-6.10: Row/Column operations
**Tools:** `xls_insert_rows.py`, `xls_delete_rows.py` ⚠️, `xls_insert_columns.py`, `xls_delete_columns.py` ⚠️, `xls_move_sheet.py`

**Shared Checklist:**
- [ ] Range validation
- [ ] Token check for deletions
- [ ] Formula offset calculation
- [ ] Style copying for insertions
- [ ] Merged cell handling
- [ ] Hidden row/column preservation
- [ ] Freeze pane adjustment

---

## PHASE 7: Formula & Calculation Engine

**Objective:** Implement formula manipulation and calculation tools  
**Dependencies:** Phases 1-6 completed  
**Duration:** 4 days  
**Gate Criteria:** Two-tier calculation produces Excel-compatible results

### Files to Create

#### 7.1 `src/excel_agent_tools/calculation.py`
**Description:** Two-tier calculation engine wrapper  
**Features & Interfaces:**
```python
class CalculationEngine:
    def __init__(self, workbook: Workbook)
    def recalculate_tier1(self) -> dict  # Using formulas/pycel
    def recalculate_tier2(self, workbook_path: Path) -> dict  # LibreOffice
    def detect_errors(self) -> List[CellError]
    def get_calculated_value(self, cell_ref: str) -> Any
    def set_calculation_mode(self, mode: str) -> None  # automatic/manual
    def evaluate_formula(self, formula: str, context: dict) -> Any
```

**Checklist:**
- [ ] Tier 1: formulas library integration
- [ ] Tier 1: Common functions (SUM, AVERAGE, IF, VLOOKUP)
- [ ] Tier 2: LibreOffice process management
- [ ] Tier 2: Timeout handling (default 30s)
- [ ] Error cell detection (#REF!, #VALUE!, etc.)
- [ ] Circular reference handling
- [ ] Volatile function marking
- [ ] Calculation dependency order

#### 7.2 `scripts/recalc.py`
**Description:** LibreOffice headless wrapper script  
**Features & Interfaces:**
```python
#!/usr/bin/env python3
# Usage: recalc.py input.xlsx output.xlsx --timeout 30
# Spawns: soffice --headless --convert-to xlsx --outdir /tmp
```

**Checklist:**
- [ ] LibreOffice installation detection
- [ ] Process isolation
- [ ] Timeout enforcement
- [ ] Error stream capture
- [ ] Memory limit (1GB default)
- [ ] Temporary file cleanup
- [ ] Platform-specific paths

#### 7.3-7.8: Formula tools
**Tools:** 
- `xls_set_formula.py` - Set cell formula
- `xls_recalculate.py` - Force recalculation
- `xls_detect_errors.py` - Find error cells
- `xls_convert_to_values.py` ⚠️ - Replace formulas
- `xls_copy_formula_down.py` - Auto-fill
- `xls_define_name.py` - Create named range

**Shared Checklist:**
- [ ] Formula syntax validation
- [ ] Reference validity check
- [ ] Calculation mode respect
- [ ] Array formula handling
- [ ] Performance metrics output

---

## PHASE 8: Macro & Security Tools

**Objective:** Implement VBA inspection and manipulation tools  
**Dependencies:** Phases 1-7 completed  
**Duration:** 3 days  
**Gate Criteria:** Correctly identifies macro risks without corrupting signatures

### Files to Create

#### 8.1 `src/excel_agent_tools/macros.py`
**Description:** Macro inspection and manipulation  
**Features & Interfaces:**
```python
class MacroHandler:
    def __init__(self, workbook_path: Path)
    def has_vba_project(self) -> bool
    def get_vba_modules(self) -> List[VBAModule]
    def scan_risks(self) -> MacroRiskReport
    def extract_iocs(self) -> List[str]  # IPs, URLs, file paths
    def has_digital_signature(self) -> bool
    def remove_vba_project(self, token: str) -> None
    def inject_vba_project(self, vba_bin_path: Path, token: str) -> None
```

**Checklist:**
- [ ] oletools integration (with abstraction layer)
- [ ] olevba for code extraction
- [ ] Auto-exec detection (AutoOpen, Document_Open)
- [ ] Suspicious keyword scanning
- [ ] XLM/Excel 4 macro detection
- [ ] Digital signature preservation
- [ ] Binary stream manipulation without corruption
- [ ] Audit logging for all macro operations

#### 8.2-8.6: Macro tools
**Tools:**
- `xls_has_macros.py` - Boolean macro check
- `xls_inspect_macros.py` - List VBA modules
- `xls_validate_macro_safety.py` - Security risk scan
- `xls_remove_macros.py` ⚠️⚠️ - Strip VBA project
- `xls_inject_vba_project.py` ⚠️ - Add VBA from .bin file

**Shared Checklist:**
- [ ] .xlsm file handling
- [ ] Token validation for modifications
- [ ] Signature status reporting
- [ ] Risk scoring (0-100)
- [ ] IOC extraction

---

## PHASE 9: Objects, Formatting & Export

**Objective:** Implement visualization, formatting, and export tools  
**Dependencies:** Phases 1-8 completed  
**Duration:** 5 days  
**Gate Criteria:** Generated files open in Excel without repair prompts

### Files to Create

#### 9.1 `src/excel_agent_tools/objects.py`
**Description:** Charts, tables, images, and other objects  
**Features & Interfaces:**
```python
class ObjectManager:
    def add_table(self, range_spec: str, name: str, style: str) -> Table
    def add_chart(self, data_range: str, chart_type: str, position: str) -> Chart
    def add_image(self, image_path: Path, anchor: str, width: int, height: int) -> Image
    def add_comment(self, cell_ref: str, text: str, author: str) -> Comment
    def add_sparkline(self, data_range: str, location: str) -> Sparkline
    def set_data_validation(self, range_spec: str, validation_type: str, **kwargs) -> None
```

**Checklist:**
- [ ] Table creation with headers
- [ ] Chart types: Bar, Line, Pie, Scatter, Area
- [ ] Image format support: PNG, JPEG, SVG
- [ ] Comment threading support
- [ ] Sparkline types: Line, Column, Win/Loss
- [ ] Validation types: List, Number, Date, TextLength

#### 9.2 `src/excel_agent_tools/formatting.py`
**Description:** Cell and range formatting  
**Features & Interfaces:**
```python
class Formatter:
    def format_range(self, range_spec: str, style_dict: dict) -> None
    def set_column_width(self, columns: str, width: float) -> None
    def set_row_height(self, rows: str, height: float) -> None
    def freeze_panes(self, cell_ref: str) -> None
    def apply_conditional_formatting(self, range_spec: str, rule_type: str, **kwargs) -> None
    def set_number_format(self, range_spec: str, format_code: str) -> None
```

**Checklist:**
- [ ] Style inheritance
- [ ] Auto-fit algorithms
- [ ] Conditional format types: ColorScale, DataBar, IconSet, Expression
- [ ] Number format validation
- [ ] Theme color support

#### 9.3 `src/excel_agent_tools/exporters.py`
**Description:** Export to various formats  
**Features & Interfaces:**
```python
class Exporter:
    def export_pdf(self, workbook_path: Path, output_path: Path, **options) -> None
    def export_csv(self, sheet_name: str, output_path: Path, **options) -> None
    def export_json(self, range_spec: str, output_path: Path, **options) -> None
    def export_html(self, sheet_name: str, output_path: Path, **options) -> None
```

**Checklist:**
- [ ] PDF via LibreOffice
- [ ] CSV encoding options
- [ ] JSON structure options
- [ ] HTML with styles

#### 9.4-9.18: Object & formatting tools (15 tools)
Including: `xls_add_table.py`, `xls_add_chart.py`, `xls_add_image.py`, `xls_format_range.py`, `xls_export_pdf.py`, etc.

---

## PHASE 10: Integration, Documentation & Release

**Objective:** End-to-end testing, documentation, and release preparation  
**Dependencies:** All previous phases completed  
**Duration:** 3 days  
**Gate Criteria:** New user can complete complex workflow using only documentation

### Files to Create

#### 10.1 `README.md`
**Description:** Comprehensive project documentation  
**Sections:**
```markdown
# Excel Agent Tools
## Quick Start
## Installation
## Architecture
## Tool Catalog (53 tools with examples)
## Security Model
## API Reference
## Contributing
## License
```

**Checklist:**
- [ ] Installation instructions for all platforms
- [ ] Architecture diagram
- [ ] Tool categorization
- [ ] Usage examples for each tool
- [ ] Security best practices
- [ ] Troubleshooting guide
- [ ] Performance tuning tips

#### 10.2 `docs/METICULOUS_APPROACH.md`
**Description:** Design philosophy documentation  
**Checklist:**
- [ ] Six-phase methodology
- [ ] Anti-generic principles
- [ ] Governance-first design
- [ ] Decision rationale
- [ ] Trade-off analysis

#### 10.3 `docs/API.md`
**Description:** Python API documentation  
**Checklist:**
- [ ] Class references
- [ ] Method signatures
- [ ] Code examples
- [ ] Type annotations
- [ ] Exception catalog

#### 10.4 `tests/integration/test_workflows.py`
**Description:** End-to-end workflow tests  
**Test Scenarios:**
```python
def test_clone_modify_recalc_workflow(): ...
def test_dependency_aware_deletion(): ...
def test_macro_safety_workflow(): ...
def test_large_file_processing(): ...
def test_multi_tool_pipeline(): ...
```

**Checklist:**
- [ ] 20+ real-world workflows
- [ ] Error recovery scenarios
- [ ] Performance benchmarks
- [ ] Memory leak detection
- [ ] Concurrent access tests

#### 10.5 `tests/integration/test_ai_agent_simulation.py`
**Description:** Simulate AI agent using tools  
**Checklist:**
- [ ] JSON communication
- [ ] Tool chaining
- [ ] Error handling
- [ ] Token generation and validation
- [ ] Audit trail verification

#### 10.6 `scripts/install_libreoffice.sh`
**Description:** LibreOffice installation helper  
**Checklist:**
- [ ] Ubuntu/Debian support
- [ ] RHEL/CentOS support
- [ ] macOS support
- [ ] Windows WSL support
- [ ] Version verification

#### 10.7 `examples/`
**Description:** Example workflows and scripts  
**Files:**
- `examples/financial_report_automation.py`
- `examples/dependency_visualization.py`
- `examples/bulk_formula_update.py`
- `examples/macro_risk_assessment.py`
- `examples/large_file_processing.py`

**Checklist:**
- [ ] Real-world scenarios
- [ ] Well-commented code
- [ ] Error handling
- [ ] Performance optimization

#### 10.8 `CHANGELOG.md`
**Description:** Version history  
**Checklist:**
- [ ] Version 1.0.0 features
- [ ] Breaking changes
- [ ] Migration guides
- [ ] Known issues

#### 10.9 `SECURITY.md`
**Description:** Security policies  
**Checklist:**
- [ ] Vulnerability reporting
- [ ] Security model
- [ ] Token best practices
- [ ] Audit configuration

#### 10.10 `.pre-commit-config.yaml`
**Description:** Pre-commit hooks configuration  
**Checklist:**
- [ ] Black formatting
- [ ] isort imports
- [ ] Flake8 linting
- [ ] mypy type checking
- [ ] Security scanning

---

## Master Timeline

| Week | Phases | Key Deliverables | Milestone |
|:-----|:-------|:-----------------|:----------|
| 1 | Phase 1 | Core library, serializers, file locking | Foundation complete |
| 2 | Phase 2-3 | Dependency engine, token system | Governance active |
| 3 | Phase 4-5 | Read/write tools (17 tools) | Data I/O complete |
| 4-5 | Phase 6 | Structural operations (10 tools) | Structure manipulation complete |
| 5-6 | Phase 7 | Formula & calculation (6 tools) | Calculation engine active |
| 6-7 | Phase 8-9 | Macros, objects, formatting (20 tools) | Full feature set |
| 8 | Phase 10 | Documentation, integration tests, release | Production ready |

---

## Risk Mitigation Matrix

| Risk | Probability | Impact | Mitigation |
|:-----|:-----------|:-------|:-----------|
| LibreOffice unavailable | Medium | High | Fallback to formulas library for basic calc |
| oletools unmaintained | High | Medium | Abstraction layer for easy replacement |
| Large file memory issues | Medium | High | Streaming I/O with pandas chunks |
| Formula parser bugs | Medium | Medium | Extensive test suite with edge cases |
| VBA corruption | Low | Critical | Never modify binary, only inject pre-validated |
| Token bypass attempts | Low | High | Constant-time comparison, audit all attempts |
| Cross-platform issues | Medium | Medium | CI matrix testing on all platforms |

---

## Quality Gates

### Phase Gate Criteria

Each phase must meet these criteria before proceeding:

1. **Code Coverage:** >90% unit test coverage
2. **Type Coverage:** 100% mypy strict mode
3. **Documentation:** All public APIs documented
4. **Performance:** Meets benchmark targets
5. **Security:** Passes security scan
6. **Integration:** E2E test passes

### Final Release Criteria

1. All 53 tools implemented and tested
2. Documentation allows new user success in <1 hour
3. Performance: 100k rows in <2s, 1M rows in <20s
4. Zero critical security vulnerabilities
5. Cross-platform tests pass (Linux, macOS, Windows)
6. AI agent simulation completes 10 workflows successfully

---

## Appendix: File Count Summary

| Category | Count | Phase |
|:---------|:------|:------|
| Core library modules | 11 | 1-3 |
| CLI tools | 53 | 4-9 |
| Test modules | 12 | 1-10 |
| Documentation | 6 | 10 |
| Configuration | 3 | 1, 10 |
| Scripts | 2 | 7, 10 |
| **Total Files** | **87** | — |

This Master Execution Plan provides the meticulous, phase-by-phase blueprint for implementing `excel-agent-tools` with the same rigor and governance-first philosophy as the PowerPoint agent tools reference architecture.
