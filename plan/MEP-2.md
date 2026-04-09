# Master Execution Plan: `excel-agent-tools`

## Project Structure Overview

```
excel-agent-tools/
├── README.md
├── LICENSE
├── pyproject.toml
├── setup.py
├── requirements.txt
├── requirements-dev.txt
├── .gitignore
├── .github/
│   └── workflows/
│       └── ci.yml
├── src/
│   └── excel_agent/
│       ├── __init__.py
│       ├── core/
│       │   ├── __init__.py
│       │   ├── agent.py
│       │   ├── locking.py
│       │   ├── serializers.py
│       │   ├── dependency.py
│       │   ├── macro_handler.py
│       │   ├── chunked_io.py
│       │   └── version_hash.py
│       ├── governance/
│       │   ├── __init__.py
│       │   ├── token_manager.py
│       │   ├── audit_trail.py
│       │   └── schemas/
│       │       ├── __init__.py
│       │       └── *.schema.json
│       ├── calculation/
│       │   ├── __init__.py
│       │   ├── tier1_engine.py
│       │   ├── tier2_libreoffice.py
│       │   └── error_detector.py
│       └── utils/
│           ├── __init__.py
│           ├── exit_codes.py
│           ├── json_io.py
│           └── cli_helpers.py
├── tools/
│   ├── governance/
│   │   ├── xls_clone_workbook.py
│   │   ├── xls_validate_workbook.py
│   │   ├── xls_approve_token.py
│   │   ├── xls_version_hash.py
│   │   ├── xls_lock_status.py
│   │   └── xls_dependency_report.py
│   ├── read/
│   │   ├── xls_read_range.py
│   │   ├── xls_get_sheet_names.py
│   │   ├── xls_get_defined_names.py
│   │   ├── xls_get_table_info.py
│   │   ├── xls_get_cell_style.py
│   │   ├── xls_get_formula.py
│   │   └── xls_get_workbook_metadata.py
│   ├── write/
│   │   ├── xls_create_new.py
│   │   ├── xls_create_from_template.py
│   │   ├── xls_write_range.py
│   │   └── xls_write_cell.py
│   ├── structure/
│   │   ├── xls_add_sheet.py
│   │   ├── xls_delete_sheet.py
│   │   ├── xls_rename_sheet.py
│   │   ├── xls_insert_rows.py
│   │   ├── xls_delete_rows.py
│   │   ├── xls_insert_columns.py
│   │   ├── xls_delete_columns.py
│   │   └── xls_move_sheet.py
│   ├── cells/
│   │   ├── xls_merge_cells.py
│   │   ├── xls_unmerge_cells.py
│   │   ├── xls_delete_range.py
│   │   └── xls_update_references.py
│   ├── formulas/
│   │   ├── xls_set_formula.py
│   │   ├── xls_recalculate.py
│   │   ├── xls_detect_errors.py
│   │   ├── xls_convert_to_values.py
│   │   ├── xls_copy_formula_down.py
│   │   └── xls_define_name.py
│   ├── objects/
│   │   ├── xls_add_table.py
│   │   ├── xls_add_chart.py
│   │   ├── xls_add_image.py
│   │   ├── xls_add_comment.py
│   │   └── xls_set_data_validation.py
│   ├── formatting/
│   │   ├── xls_format_range.py
│   │   ├── xls_set_column_width.py
│   │   ├── xls_freeze_panes.py
│   │   ├── xls_apply_conditional_formatting.py
│   │   └── xls_set_number_format.py
│   ├── macros/
│   │   ├── xls_has_macros.py
│   │   ├── xls_inspect_macros.py
│   │   ├── xls_validate_macro_safety.py
│   │   ├── xls_remove_macros.py
│   │   └── xls_inject_vba_project.py
│   └── export/
│       ├── xls_export_pdf.py
│       ├── xls_export_csv.py
│       └── xls_export_json.py
├── scripts/
│   ├── recalc.py                    # LibreOffice headless wrapper
│   ├── install_libreoffice.sh       # Setup script for CI
│   └── generate_test_files.py       # Creates test .xlsx/.xlsm files
├── tests/
│   ├── __init__.py
│   ├── conftest.py
│   ├── fixtures/
│   │   ├── sample.xlsx
│   │   ├── sample_with_macros.xlsm
│   │   ├── large_dataset.xlsx
│   │   └── complex_formulas.xlsx
│   ├── unit/
│   │   ├── test_agent.py
│   │   ├── test_locking.py
│   │   ├── test_dependency.py
│   │   ├── test_token_manager.py
│   │   └── ...
│   ├── integration/
│   │   ├── test_clone_modify_workflow.py
│   │   ├── test_formula_dependency_workflow.py
│   │   └── ...
│   └── property/
│       └── test_range_serializer.py  # Hypothesis property tests
└── docs/
    ├── DESIGN.md                      # This blueprint
    ├── API.md                         # CLI interface reference
    ├── WORKFLOWS.md                   # Common agent workflows
    ├── GOVERNANCE.md                  # Token & safety protocols
    └── DEVELOPMENT.md                 # Contributing guide
```

---

# Phase 0: Project Scaffolding & Infrastructure
**Duration:** 2 days  
**Goal:** Establish project structure, tooling, CI/CD, and development environment.

## Files to Create

### 1. `README.md`
**Purpose:** Project overview, installation, quick start guide.

**Features:**
- Project description and philosophy (Governance-First, AI-Native)
- Installation instructions (pip, from source)
- Quick start example (clone → modify → recalc workflow)
- Link to full documentation
- Badge for CI status, coverage, PyPI version

**Checklist:**
- [ ] Project tagline and description
- [ ] Installation instructions for Python 3.9+
- [ ] LibreOffice headless installation guide (Linux, macOS, Windows)
- [ ] Quick start code example with 3-step workflow
- [ ] Links to `docs/` for detailed documentation
- [ ] License badge and contribution guidelines link
- [ ] Requirements: Python 3.9+, openpyxl 3.1.5+, oletools, formulas

---

### 2. `LICENSE`
**Purpose:** MIT License file.

**Checklist:**
- [ ] MIT License text
- [ ] Copyright year: 2026
- [ ] Copyright holder name

---

### 3. `pyproject.toml`
**Purpose:** Modern Python project metadata (PEP 518/621).

**Features:**
- Build system configuration (setuptools)
- Project metadata (name, version, authors, description)
- Dependencies with version constraints
- Optional dependencies for dev/test
- Entry points for CLI tools (all 53 tools)

**Checklist:**
- [ ] `[build-system]` with setuptools backend
- [ ] `[project]` metadata: name="excel-agent-tools", version="1.0.0"
- [ ] `dependencies`: openpyxl>=3.1.5, defusedxml, oletools, formulas, pandas
- [ ] `[project.optional-dependencies]` dev: pytest, pytest-cov, hypothesis, black, mypy
- [ ] `[project.scripts]` with all 53 tool entry points
- [ ] Python version constraint: requires-python = ">=3.9"
- [ ] `[tool.black]` configuration
- [ ] `[tool.mypy]` strict typing configuration
- [ ] `[tool.pytest.ini_options]` test discovery settings

---

### 4. `requirements.txt`
**Purpose:** Core runtime dependencies.

**Checklist:**
- [ ] openpyxl>=3.1.5
- [ ] defusedxml>=0.7.0
- [ ] oletools>=0.60
- [ ] formulas>=1.2.0
- [ ] pandas>=2.0.0
- [ ] Pinned versions for reproducibility

---

### 5. `requirements-dev.txt`
**Purpose:** Development and testing dependencies.

**Checklist:**
- [ ] pytest>=7.0.0
- [ ] pytest-cov>=4.0.0
- [ ] hypothesis>=6.0.0
- [ ] black>=23.0.0
- [ ] mypy>=1.0.0
- [ ] flake8>=6.0.0
- [ ] pre-commit>=3.0.0

---

### 6. `.gitignore`
**Purpose:** Exclude generated files from version control.

**Checklist:**
- [ ] Python cache: `__pycache__/`, `*.pyc`, `*.pyo`, `.pytest_cache/`
- [ ] Virtual environments: `venv/`, `.venv/`, `env/`
- [ ] IDEs: `.vscode/`, `.idea/`, `*.swp`
- [ ] Build artifacts: `build/`, `dist/`, `*.egg-info/`
- [ ] Test outputs: `.coverage`, `htmlcov/`, `.tox/`
- [ ] Excel working files: `/work/`, `*.tmp.xlsx`, `.~lock.*`
- [ ] Audit trail (unless explicitly committed): `.excel_agent_audit.jsonl`

---

### 7. `.github/workflows/ci.yml`
**Purpose:** GitHub Actions CI/CD pipeline.

**Features:**
- Test on Python 3.9, 3.10, 3.11, 3.12
- Install LibreOffice headless for Tier 2 calculation tests
- Run unit, integration, and property tests
- Coverage reporting (codecov)
- Linting (black, mypy, flake8)

**Checklist:**
- [ ] Matrix strategy: Python 3.9-3.12, OS: ubuntu-latest
- [ ] Install LibreOffice: `sudo apt-get install -y libreoffice-calc`
- [ ] Install dependencies: `pip install -r requirements.txt -r requirements-dev.txt`
- [ ] Run linters: `black --check .`, `mypy src/`, `flake8 src/`
- [ ] Run tests: `pytest --cov=excel_agent --cov-report=xml`
- [ ] Upload coverage to codecov
- [ ] Fail build if coverage <90%

---

### 8. `setup.py` (Legacy Compatibility)
**Purpose:** Fallback for older pip versions.

**Checklist:**
- [ ] Minimal shim that defers to `pyproject.toml`
- [ ] `from setuptools import setup; setup()`

---

### 9. `src/excel_agent/__init__.py`
**Purpose:** Package initialization and version export.

**Checklist:**
- [ ] `__version__ = "1.0.0"`
- [ ] Import key classes for convenience: `ExcelAgent`, `DependencyTracker`, `ApprovalTokenManager`
- [ ] `__all__` list for public API

---

### 10. `src/excel_agent/utils/exit_codes.py`
**Purpose:** Standardized exit code constants.

**Features:**
- Enum or constants for all exit codes
- Consistent with PowerPoint agent tools

**Interface:**
```python
class ExitCode(IntEnum):
    SUCCESS = 0
    VALIDATION_ERROR = 1
    FILE_NOT_FOUND = 2
    LOCK_CONTENTION = 3
    PERMISSION_DENIED = 4
    INTERNAL_ERROR = 5
```

**Checklist:**
- [ ] Define `ExitCode` enum
- [ ] Add docstrings for each code
- [ ] Export in `__init__.py`

---

### 11. `src/excel_agent/utils/json_io.py`
**Purpose:** Standardized JSON output formatting.

**Features:**
- Output schema builder
- Date/datetime serialization
- Pretty-print option for debugging

**Interface:**
```python
def build_response(
    status: str,
    data: Any,
    workbook_version: str,
    impact: Optional[dict] = None,
    warnings: Optional[list] = None,
    exit_code: int = 0
) -> dict:
    """Builds standardized JSON response."""
    ...

def print_json(data: dict) -> None:
    """Prints JSON to stdout with indent=2."""
    ...
```

**Checklist:**
- [ ] Implement `build_response()` with schema validation
- [ ] Custom JSON encoder for `datetime`, `Path` objects
- [ ] `print_json()` outputs to stdout only (no stderr pollution)
- [ ] Unit tests for edge cases (None values, nested dicts)

---

### 12. `src/excel_agent/utils/cli_helpers.py`
**Purpose:** Reusable CLI argument parsing and validation.

**Features:**
- Common arguments (input path, output path, token, force flag)
- JSON input parsing from stdin or file
- Schema validation against JSON schemas

**Interface:**
```python
def add_common_args(parser: argparse.ArgumentParser) -> None:
    """Adds --input, --output, --token, --force to parser."""
    ...

def validate_json_input(schema_name: str, data: dict) -> None:
    """Validates data against schema. Raises ValidationError."""
    ...

def load_json_input() -> dict:
    """Reads JSON from stdin."""
    ...
```

**Checklist:**
- [ ] `add_common_args()` with standardized flag names
- [ ] `validate_json_input()` using `jsonschema` library
- [ ] `load_json_input()` with error handling for malformed JSON
- [ ] Helper for path validation (exists, readable, writable)

---

**Phase 0 Exit Criteria:**
- [ ] All 12 files created and pass linting (black, mypy, flake8)
- [ ] CI pipeline runs successfully on empty test suite
- [ ] Virtual environment created with all dependencies installed
- [ ] LibreOffice headless installed and `soffice --headless --version` works
- [ ] README renders correctly on GitHub
- [ ] Project can be installed with `pip install -e .`

---

# Phase 1: Core Foundation
**Duration:** 5 days  
**Goal:** Implement the hub (ExcelAgent, RangeSerializer, file locking, version hashing).

## Files to Create

### 13. `src/excel_agent/core/locking.py`
**Purpose:** Cross-platform atomic file locking.

**Features:**
- OS-level file locking (fcntl on Linux/macOS, msvcrt on Windows)
- Context manager for automatic release
- Timeout and retry logic
- Lock contention detection

**Interface:**
```python
class FileLock:
    def __init__(self, path: Path, timeout: float = 30.0):
        """Creates lock for given file path."""
        ...
    
    def __enter__(self) -> 'FileLock':
        """Acquires lock or raises LockContentionError."""
        ...
    
    def __exit__(self, *args):
        """Releases lock."""
        ...
    
    def is_locked(self) -> bool:
        """Checks if file is currently locked (read-only check)."""
        ...
```

**Checklist:**
- [ ] Implement `fcntl.flock()` for Unix systems
- [ ] Implement `msvcrt.locking()` for Windows
- [ ] Timeout with exponential backoff (0.1s, 0.2s, 0.5s, 1s intervals)
- [ ] Raise `LockContentionError` (exit code 3) on timeout
- [ ] Unit tests: concurrent locking from multiple processes
- [ ] Unit tests: lock release on exception
- [ ] Unit tests: `is_locked()` accuracy

---

### 14. `src/excel_agent/core/serializers.py`
**Purpose:** Unified range parsing (A1, R1C1, Named Ranges, Table references).

**Features:**
- Parse all range formats to internal coordinates
- Reverse conversion (coordinates → A1 notation)
- Table column references (e.g., `Table1[Sales]`)
- Validation and error messages

**Interface:**
```python
@dataclass
class CellCoordinate:
    row: int
    col: int

@dataclass
class RangeCoordinate:
    min_row: int
    min_col: int
    max_row: int
    max_col: int

class RangeSerializer:
    def __init__(self, workbook: Workbook):
        """Initializes with workbook context for named ranges/tables."""
        ...
    
    def parse(self, range_str: str, sheet_name: Optional[str] = None) -> RangeCoordinate:
        """Parses A1, R1C1, Name, or Table[Column] to coordinates."""
        ...
    
    def to_a1(self, coord: RangeCoordinate) -> str:
        """Converts coordinates to A1 notation."""
        ...
    
    def to_r1c1(self, coord: RangeCoordinate) -> str:
        """Converts coordinates to R1C1 notation."""
        ...
```

**Checklist:**
- [ ] Parse A1 notation: `A1`, `A1:C10`, `Sheet1!A1:C10`
- [ ] Parse R1C1 notation: `R1C1`, `R1C1:R10C3`
- [ ] Parse Named Ranges: resolve via `workbook.defined_names`
- [ ] Parse Table references: `Table1`, `Table1[Column1]`, `Table1[#All]`
- [ ] Validate: raise `ValueError` for malformed inputs
- [ ] Reverse conversion: `to_a1()`, `to_r1c1()`
- [ ] Unit tests with hypothesis for property-based testing (roundtrip)
- [ ] Edge cases: single cell, full row (`A:A`), full column (`1:1`)

---

### 15. `src/excel_agent/core/version_hash.py`
**Purpose:** Geometry-aware workbook hashing for concurrency detection.

**Features:**
- Hash based on structure (sheet names, cell coordinates, formulas)
- Excludes cell values for performance
- Detects concurrent modifications

**Interface:**
```python
def compute_workbook_hash(workbook: Workbook) -> str:
    """Returns SHA256 hash of workbook geometry.
    Includes: sheet names, order, cell coordinates with formulas.
    Excludes: cell values, styles."""
    ...

def compute_sheet_hash(sheet: Worksheet) -> str:
    """Returns hash of single sheet geometry."""
    ...
```

**Checklist:**
- [ ] Iterate sheets in order
- [ ] For each sheet: hash (name, visibility, used_range coordinates)
- [ ] For cells with formulas: hash (row, col, formula_string)
- [ ] Ignore cell values (data_only=False mode)
- [ ] Return hex digest of SHA256
- [ ] Unit tests: identical structure = identical hash
- [ ] Unit tests: value change = same hash
- [ ] Unit tests: formula change = different hash
- [ ] Unit tests: sheet rename = different hash

---

### 16. `src/excel_agent/core/agent.py`
**Purpose:** The hub — stateful context manager for safe workbook manipulation.

**Features:**
- Context manager lifecycle (acquire lock, load, save, release)
- Integration with FileLock and version_hash
- `keep_vba=True` enforcement for .xlsm
- Concurrent modification detection

**Interface:**
```python
class ExcelAgent:
    def __init__(
        self,
        path: Path,
        *,
        mode: str = "rw",
        keep_vba: bool = True,
        lock_timeout: float = 30.0
    ):
        """Initializes agent for given workbook."""
        ...
    
    def __enter__(self) -> 'ExcelAgent':
        """Acquires lock, loads workbook, computes entry hash."""
        ...
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Saves workbook (if rw mode), verifies hash, releases lock."""
        ...
    
    @property
    def workbook(self) -> Workbook:
        """Returns openpyxl Workbook object."""
        ...
    
    @property
    def version_hash(self) -> str:
        """Returns current workbook version hash."""
        ...
    
    def verify_no_concurrent_modification(self) -> None:
        """Recomputes hash and raises ConcurrentModificationError if changed."""
        ...
```

**Checklist:**
- [ ] `__init__`: validate path exists (for mode 'r' or 'rw'), create parent dir (for mode 'w')
- [ ] `__enter__`: acquire FileLock, call `load_workbook(keep_vba=True, data_only=False)`
- [ ] `__enter__`: compute and store entry hash
- [ ] `__exit__`: if no exception and mode='rw', verify hash hasn't changed
- [ ] `__exit__`: call `workbook.save()`, release lock
- [ ] `__exit__`: if exception, release lock without saving
- [ ] `verify_no_concurrent_modification()`: re-read file, compare hashes
- [ ] Raise `ConcurrentModificationError` (maps to exit code 5) if hash mismatch
- [ ] Unit tests: successful load/save cycle
- [ ] Unit tests: lock contention with timeout
- [ ] Unit tests: concurrent modification detection (simulate external change)
- [ ] Unit tests: exception handling (lock release on error)
- [ ] Unit tests: .xlsm file preserves VBA project

---

**Phase 1 Exit Criteria:**
- [ ] All 4 core files pass unit tests with >95% coverage
- [ ] FileLock works on Linux (CI) and can be manually tested on Windows/macOS
- [ ] RangeSerializer handles all 4 input formats (A1, R1C1, Name, Table)
- [ ] ExcelAgent successfully loads sample.xlsx, makes a change, saves, and hash updates
- [ ] Hypothesis property tests pass for RangeSerializer (100 examples)
- [ ] No mypy errors in strict mode
- [ ] Integration test: two processes attempting to lock same file, one gets exit code 3

---

# Phase 2: Dependency Engine & Validation
**Duration:** 5 days  
**Goal:** Implement formula dependency graph and workbook validation.

## Files to Create

### 17. `src/excel_agent/core/dependency.py`
**Purpose:** Build and query formula dependency graph using `formulas` library.

**Features:**
- Parse all formulas in workbook into AST
- Build directed graph of cell dependencies
- Transitive closure for impact analysis
- Circular reference detection (topological sort)

**Interface:**
```python
@dataclass
class DependencyImpact:
    broken_references: int
    affected_sheets: List[str]
    sample_errors: List[str]
    suggestion: str

class DependencyTracker:
    def __init__(self, workbook: Workbook):
        """Initializes tracker for given workbook."""
        ...
    
    def build_graph(self, sheets: Optional[List[str]] = None) -> None:
        """Builds dependency graph for specified sheets (or all)."""
        ...
    
    def find_dependents(self, target: str) -> Set[str]:
        """Returns all cells that would break if target is deleted.
        Target format: 'Sheet1!A1'"""
        ...
    
    def impact_report(self, target_range: str, action: str) -> DependencyImpact:
        """Pre-flight check for deletion/modification.
        Returns impact report with broken_references count."""
        ...
    
    def detect_circular_references(self) -> List[List[str]]:
        """Returns list of circular dependency chains."""
        ...
    
    def get_graph_adjacency_list(self) -> Dict[str, List[str]]:
        """Exports graph as JSON-serializable adjacency list."""
        ...
```

**Checklist:**
- [ ] Integrate `formulas` library for AST parsing
- [ ] Iterate all sheets and cells with formulas
- [ ] Extract cell references from formula AST
- [ ] Build forward graph: `{cell: [cells_it_references]}`
- [ ] Build reverse graph: `{cell: [cells_that_reference_it]}`
- [ ] `find_dependents()`: perform transitive closure (DFS/BFS)
- [ ] `impact_report()`: count broken refs, sample first 5 errors
- [ ] `detect_circular_references()`: Tarjan's algorithm for strongly connected components
- [ ] Handle cross-sheet references: `'Sheet1'!A1`
- [ ] Handle named ranges in formulas
- [ ] Unit tests: simple chain (A1→B1→C1, deleting A1 affects B1 and C1)
- [ ] Unit tests: cross-sheet dependencies
- [ ] Unit tests: circular reference detection (A1=B1, B1=A1)
- [ ] Unit tests: large workbook (10 sheets, 1000 formulas) builds in <5s
- [ ] Performance optimization: lazy graph (only build on-demand)

---

### 18. `src/excel_agent/governance/schemas/range_input.schema.json`
**Purpose:** JSON schema for range input validation.

**Features:**
- Defines structure for range specifications
- Used by tools accepting range arguments

**Schema:**
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "object",
  "properties": {
    "range": {
      "oneOf": [
        {"type": "string", "pattern": "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$"},
        {"type": "object", "properties": {
          "start_row": {"type": "integer", "minimum": 1},
          "start_col": {"type": "integer", "minimum": 1},
          "end_row": {"type": "integer", "minimum": 1},
          "end_col": {"type": "integer", "minimum": 1}
        }, "required": ["start_row", "start_col"]}
      ]
    },
    "sheet": {"type": "string"}
  },
  "required": ["range"]
}
```

**Checklist:**
- [ ] Schema validates A1 notation strings
- [ ] Schema validates coordinate objects
- [ ] Optional `sheet` property
- [ ] Unit tests: valid inputs pass
- [ ] Unit tests: invalid inputs (malformed A1) fail

---

### 19. `src/excel_agent/governance/schemas/write_data.schema.json`
**Purpose:** JSON schema for cell data input.

**Schema:**
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "object",
  "properties": {
    "data": {
      "type": "array",
      "items": {
        "type": "array",
        "items": {
          "oneOf": [
            {"type": "string"},
            {"type": "number"},
            {"type": "boolean"},
            {"type": "null"}
          ]
        }
      }
    }
  },
  "required": ["data"]
}
```

**Checklist:**
- [ ] Schema validates 2D array of cells
- [ ] Supports string, number, boolean, null
- [ ] Unit tests for edge cases (empty array, nested nulls)

---

### 20. `src/excel_agent/governance/schemas/__init__.py`
**Purpose:** Schema loader utility.

**Interface:**
```python
def load_schema(schema_name: str) -> dict:
    """Loads JSON schema by name from schemas/ directory."""
    ...

def validate_against_schema(schema_name: str, data: dict) -> None:
    """Validates data, raises jsonschema.ValidationError."""
    ...
```

**Checklist:**
- [ ] `load_schema()` reads .schema.json files
- [ ] `validate_against_schema()` uses `jsonschema.validate()`
- [ ] Cache loaded schemas in memory
- [ ] Unit tests for all schema files

---

### 21. `tests/unit/test_dependency.py`
**Purpose:** Comprehensive unit tests for DependencyTracker.

**Test Cases:**
- [ ] Empty workbook (no formulas) → empty graph
- [ ] Single cell formula (A1=5) → no dependencies
- [ ] Single dependency (A1=B1) → graph shows A1 depends on B1
- [ ] Chain (A1=B1, B1=C1, C1=5) → find_dependents(C1) returns {B1, A1}
- [ ] Cross-sheet (Sheet1!A1=Sheet2!B1) → dependency tracked
- [ ] Circular reference (A1=B1, B1=A1) → detect_circular_references() returns cycle
- [ ] Impact report for deletion (deleting C1 breaks B1 and A1) → count=2
- [ ] Named range in formula (A1=SalesData) → resolves to range
- [ ] Large workbook stress test (1000 formulas, build in <5s)

---

**Phase 2 Exit Criteria:**
- [ ] DependencyTracker correctly identifies all dependencies in complex_formulas.xlsx fixture
- [ ] Circular reference detection works (test with workbook containing A1=B1, B1=C1, C1=A1)
- [ ] `impact_report()` returns accurate broken_references count
- [ ] Graph export as adjacency list is JSON-serializable
- [ ] All unit tests pass with >90% coverage
- [ ] Performance: 10-sheet, 1000-formula workbook analyzed in <5s

---

# Phase 3: Governance & Safety Layer
**Duration:** 3 days  
**Goal:** Implement approval tokens, audit trail, and safety protocols.

## Files to Create

### 22. `src/excel_agent/governance/token_manager.py`
**Purpose:** HMAC-SHA256 approval token generation and validation.

**Features:**
- Generate scoped tokens with expiration
- Validate tokens before destructive operations
- Prevent token reuse (nonce tracking)

**Interface:**
```python
class ApprovalTokenManager:
    def __init__(self, secret_key: Optional[str] = None):
        """Initializes with secret key (from env or generated)."""
        ...
    
    def generate_token(
        self,
        scope: str,
        resource: str,
        expires_in: int = 3600
    ) -> str:
        """Generates HMAC-SHA256 token.
        scope: 'sheet:delete', 'range:delete', etc.
        resource: 'Sheet1', 'A1:C10', etc.
        expires_in: seconds"""
        ...
    
    def validate_token(
        self,
        token: str,
        scope: str,
        resource: str
    ) -> bool:
        """Validates token. Returns True if valid, raises PermissionDeniedError otherwise."""
        ...
    
    def revoke_token(self, token: str) -> None:
        """Adds token to revocation list."""
        ...
```

**Implementation Details:**
- Token format: `{scope}:{resource}:{timestamp}:{hmac_signature}`
- HMAC signature over `scope|resource|timestamp` with secret key
- Secret key from env var `EXCEL_AGENT_SECRET` or randomly generated per session
- Maintain in-memory revocation list (nonce tracking)

**Checklist:**
- [ ] `generate_token()` creates HMAC with SHA256
- [ ] Token includes expiration timestamp
- [ ] `validate_token()` checks signature, expiration, scope match
- [ ] `validate_token()` raises `PermissionDeniedError` (exit code 4) if invalid
- [ ] `revoke_token()` prevents reuse
- [ ] Secret key loaded from env or generated (logged warning if generated)
- [ ] Unit tests: valid token passes
- [ ] Unit tests: expired token fails
- [ ] Unit tests: wrong scope fails
- [ ] Unit tests: tampered signature fails
- [ ] Unit tests: revoked token fails

---

### 23. `src/excel_agent/governance/audit_trail.py`
**Purpose:** Append-only audit log for all destructive operations.

**Features:**
- Log every operation requiring approval token
- JSON Lines format (.jsonl)
- Includes timestamp, user, tool, resource, outcome

**Interface:**
```python
class AuditTrail:
    def __init__(self, log_path: Path = Path(".excel_agent_audit.jsonl")):
        """Initializes audit trail logger."""
        ...
    
    def log_operation(
        self,
        tool: str,
        scope: str,
        resource: str,
        outcome: str,
        token_used: bool,
        details: Optional[dict] = None
    ) -> None:
        """Appends entry to audit log."""
        ...
    
    def query_logs(
        self,
        tool: Optional[str] = None,
        start_time: Optional[datetime] = None,
        end_time: Optional[datetime] = None
    ) -> List[dict]:
        """Queries audit log (for admin tools)."""
        ...
```

**Log Entry Format:**
```json
{
  "timestamp": "2026-04-08T14:30:22Z",
  "tool": "xls_delete_sheet.py",
  "scope": "sheet:delete",
  "resource": "Sheet1",
  "outcome": "success",
  "token_used": true,
  "user": "system",
  "pid": 12345,
  "details": {}
}
```

**Checklist:**
- [ ] `log_operation()` appends to .jsonl file (atomic write)
- [ ] Create log file if not exists
- [ ] Include ISO 8601 timestamp
- [ ] Include process ID for tracing
- [ ] `query_logs()` reads and filters entries
- [ ] Handle file rotation (optional: if log exceeds 10MB, rotate)
- [ ] Unit tests: log entry written correctly
- [ ] Unit tests: concurrent writes from multiple processes (file locking)
- [ ] Unit tests: query with time filters

---

### 24. `tests/unit/test_token_manager.py`
**Purpose:** Unit tests for ApprovalTokenManager.

**Test Cases:**
- [ ] Generate token with valid scope
- [ ] Validate correct token returns True
- [ ] Validate expired token raises PermissionDeniedError
- [ ] Validate token with wrong scope raises PermissionDeniedError
- [ ] Validate token with tampered signature raises PermissionDeniedError
- [ ] Revoke token and validate fails
- [ ] Secret key from environment variable is used
- [ ] Generated secret key triggers warning log

---

### 25. `tests/unit/test_audit_trail.py`
**Purpose:** Unit tests for AuditTrail.

**Test Cases:**
- [ ] Log entry appended to file
- [ ] Multiple entries maintain JSON Lines format
- [ ] Query by tool name returns filtered results
- [ ] Query by time range works
- [ ] Concurrent writes from 5 processes do not corrupt file

---

**Phase 3 Exit Criteria:**
- [ ] ApprovalTokenManager generates and validates tokens correctly
- [ ] Expired tokens are rejected with exit code 4
- [ ] Audit trail logs all operations to .jsonl file
- [ ] Audit log is human-readable and machine-parseable
- [ ] Concurrent write test (10 processes logging simultaneously) passes without corruption
- [ ] All unit tests pass with 100% coverage for governance module

---

# Phase 4: Read Tools (Category 2)
**Duration:** 4 days  
**Goal:** Implement all read-only introspection tools.

## Files to Create

### 26. `tools/read/xls_get_sheet_names.py`
**Purpose:** List all sheets in workbook.

**CLI Interface:**
```bash
xls_get_sheet_names.py --input sample.xlsx
```

**Output:**
```json
{
  "status": "success",
  "exit_code": 0,
  "workbook_version": "sha256:abc...",
  "data": {
    "sheets": [
      {"index": 0, "name": "Sheet1", "visibility": "visible"},
      {"index": 1, "name": "Sheet2", "visibility": "hidden"}
    ]
  }
}
```

**Checklist:**
- [ ] Import `ExcelAgent`, `json_io`, `cli_helpers`
- [ ] Parse `--input` argument
- [ ] Open workbook in read-only mode
- [ ] Iterate `workbook.sheetnames` and `workbook[name].sheet_state`
- [ ] Build JSON response with sheet index, name, visibility
- [ ] Print JSON to stdout
- [ ] Exit with code 0
- [ ] Unit test: sample.xlsx returns expected sheets
- [ ] Unit test: hidden sheet detected correctly
- [ ] CLI `--help` text is clear

---

### 27. `tools/read/xls_get_workbook_metadata.py`
**Purpose:** High-level workbook statistics.

**CLI Interface:**
```bash
xls_get_workbook_metadata.py --input sample.xlsx
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "sheet_count": 3,
    "total_formulas": 152,
    "named_ranges": 5,
    "tables": 2,
    "has_macros": false,
    "file_size_bytes": 45231
  }
}
```

**Checklist:**
- [ ] Count sheets
- [ ] Count cells with formulas across all sheets
- [ ] Count defined names (`len(workbook.defined_names)`)
- [ ] Count tables (ListObjects) by iterating sheets
- [ ] Check for macros using `MacroHandler.has_vba_project()`
- [ ] Get file size from `os.path.getsize()`
- [ ] Unit test: complex_formulas.xlsx returns accurate counts
- [ ] Unit test: .xlsm file shows `has_macros: true`

---

### 28. `tools/read/xls_read_range.py`
**Purpose:** Extract cell data from a range, with chunked streaming for large datasets.

**CLI Interface:**
```bash
xls_read_range.py --input sample.xlsx --range A1:C10 --sheet Sheet1
# For large ranges (>100k rows):
xls_read_range.py --input large.xlsx --range A1:Z100000 --chunked
```

**Output (non-chunked):**
```json
{
  "status": "success",
  "data": {
    "range": "A1:C10",
    "values": [
      ["Name", "Age", "City"],
      ["Alice", 30, "NYC"],
      ...
    ]
  }
}
```

**Output (chunked, JSON Lines):**
```jsonl
{"row": 1, "values": ["Name", "Age", "City"]}
{"row": 2, "values": ["Alice", 30, "NYC"]}
...
```

**Features:**
- Type preservation (dates as ISO 8601, booleans, numbers)
- Chunked mode for memory efficiency

**Checklist:**
- [ ] Parse `--range` using `RangeSerializer`
- [ ] Parse `--sheet` (optional, defaults to active sheet)
- [ ] Parse `--chunked` flag
- [ ] Non-chunked: read all cells into 2D array, serialize to JSON
- [ ] Chunked: yield rows one at a time as JSON Lines
- [ ] Handle cell types: string, number, boolean, date, None (empty)
- [ ] Convert dates to ISO 8601 strings
- [ ] Unit test: 10x10 range returns correct data
- [ ] Unit test: chunked mode for 100k rows completes in <3s
- [ ] Unit test: date formatting is ISO 8601
- [ ] Integration test: read range written by `xls_write_range.py` (roundtrip)

---

### 29. `tools/read/xls_get_defined_names.py`
**Purpose:** List all named ranges.

**Output:**
```json
{
  "status": "success",
  "data": {
    "named_ranges": [
      {"name": "SalesData", "scope": "Workbook", "refers_to": "Sheet1!$A$1:$C$100"},
      {"name": "TaxRate", "scope": "Sheet1", "refers_to": "Sheet1!$D$5"}
    ]
  }
}
```

**Checklist:**
- [ ] Iterate `workbook.defined_names`
- [ ] Extract name, scope (workbook-level or sheet-level), refers_to formula
- [ ] Unit test: workbook with global and local named ranges

---

### 30. `tools/read/xls_get_table_info.py`
**Purpose:** List Excel Tables (ListObjects).

**Output:**
```json
{
  "status": "success",
  "data": {
    "tables": [
      {
        "name": "Table1",
        "sheet": "Sheet1",
        "range": "A1:D100",
        "columns": ["Name", "Age", "City", "Salary"],
        "has_totals_row": false,
        "style": "TableStyleMedium2"
      }
    ]
  }
}
```

**Checklist:**
- [ ] Iterate sheets, then `sheet.tables` (openpyxl TableList)
- [ ] Extract table name, ref (range), column names, totals row, style
- [ ] Unit test: workbook with 2 tables

---

### 31. `tools/read/xls_get_formula.py`
**Purpose:** Get formula from a specific cell.

**CLI Interface:**
```bash
xls_get_formula.py --input sample.xlsx --cell A1 --sheet Sheet1
```

**Output:**
```json
{
  "status": "success",
  "data": {
    "cell": "A1",
    "formula": "=SUM(B1:B10)",
    "references": ["B1:B10"]
  }
}
```

**Checklist:**
- [ ] Parse `--cell` (single cell reference)
- [ ] Get cell object from sheet
- [ ] Check if `cell.value` starts with `=` (formula) or use `cell.data_type == 'f'`
- [ ] Extract formula string
- [ ] Parse references using openpyxl Tokenizer (optional enhancement)
- [ ] Return `null` if no formula
- [ ] Unit test: cell with formula
- [ ] Unit test: cell without formula returns null

---

### 32. `tools/read/xls_get_cell_style.py`
**Purpose:** Get formatting for a cell.

**Output:**
```json
{
  "status": "success",
  "data": {
    "font": {"name": "Arial", "size": 11, "bold": false, "color": "FF000000"},
    "fill": {"fgColor": "FFFFFFFF", "patternType": "solid"},
    "border": {"top": {"style": "thin", "color": "FF000000"}},
    "alignment": {"horizontal": "left", "vertical": "bottom"},
    "number_format": "General"
  }
}
```

**Checklist:**
- [ ] Access `cell.font`, `cell.fill`, `cell.border`, `cell.alignment`, `cell.number_format`
- [ ] Serialize to JSON (handle openpyxl objects)
- [ ] Unit test: formatted cell returns correct style

---

### 33. `src/excel_agent/core/chunked_io.py`
**Purpose:** Helper for streaming large range reads/writes using pandas.

**Interface:**
```python
def read_range_chunked(
    sheet: Worksheet,
    range_coord: RangeCoordinate,
    chunk_size: int = 10000
) -> Generator[List[List[Any]], None, None]:
    """Yields chunks of rows from range."""
    ...

def write_range_chunked(
    sheet: Worksheet,
    start_row: int,
    start_col: int,
    data_generator: Generator[List[List[Any]], None, None]
) -> int:
    """Writes chunks to sheet. Returns total rows written."""
    ...
```

**Checklist:**
- [ ] Use pandas `read_excel()` with `chunksize` for reading
- [ ] For writing: batch inserts every 10k rows, avoid cell-by-cell writes
- [ ] Unit test: 100k row read yields correct number of chunks
- [ ] Unit test: chunked write matches non-chunked write (data integrity)

---

### 34. `tests/integration/test_read_tools.py`
**Purpose:** Integration tests for all read tools.

**Test Cases:**
- [ ] Call `xls_get_sheet_names.py` on sample.xlsx via subprocess
- [ ] Verify JSON output is valid and matches expected sheets
- [ ] Call `xls_read_range.py` on 10x10 range, verify data
- [ ] Call `xls_read_range.py --chunked` on 100k rows, count lines
- [ ] Call all 7 read tools, verify exit code 0

---

**Phase 4 Exit Criteria:**
- [ ] All 7 read tools execute successfully
- [ ] JSON output schema is consistent across all tools
- [ ] Chunked read of 500k rows completes in <3s
- [ ] Date/time values serialize correctly (ISO 8601)
- [ ] All unit tests pass
- [ ] Integration tests simulate AI agent calling tools via subprocess

---

# Phase 5: Write & Create Tools (Category 3)
**Duration:** 3 days  
**Goal:** Implement workbook creation and data writing tools.

## Files to Create

### 35. `tools/write/xls_create_new.py`
**Purpose:** Create a blank workbook.

**CLI Interface:**
```bash
xls_create_new.py --output new.xlsx --sheets "Sheet1,Sheet2,Data"
```

**Features:**
- Create workbook with specified sheet names
- Default to single sheet "Sheet1" if not specified

**Checklist:**
- [ ] Parse `--output`, `--sheets` (comma-separated)
- [ ] Create `Workbook()` from openpyxl
- [ ] Remove default sheet, add specified sheets
- [ ] Save to output path
- [ ] Return JSON with created sheet names
- [ ] Unit test: created file opens successfully
- [ ] Unit test: sheet names match specification

---

### 36. `tools/write/xls_create_from_template.py`
**Purpose:** Clone from .xltx or .xltm template.

**CLI Interface:**
```bash
xls_create_from_template.py --template template.xltx --output output.xlsx --vars '{"company": "Acme Corp", "year": 2026}'
```

**Features:**
- Load template
- Substitute placeholders (e.g., `{{company}}`) with values from `--vars` JSON
- Save as new workbook

**Checklist:**
- [ ] Parse `--template`, `--output`, `--vars` (JSON object)
- [ ] Load template with openpyxl
- [ ] Iterate all cells, find text cells with `{{var}}` patterns
- [ ] Replace placeholders using `--vars` mapping
- [ ] Save to output path
- [ ] Unit test: template with `{{year}}` is replaced with 2026
- [ ] Unit test: formulas are preserved, not replaced

---

### 37. `tools/write/xls_write_range.py`
**Purpose:** Write data to a range with type inference.

**CLI Interface:**
```bash
xls_write_range.py --input sample.xlsx --output output.xlsx \
  --range A1 --sheet Sheet1 --data '[["Name", "Age"], ["Alice", 30]]'
```

**Features:**
- Type inference (string, number, boolean, date from ISO string)
- Expand range if data exceeds specified range
- Preserve existing styles (optional)

**Checklist:**
- [ ] Parse `--range`, `--sheet`, `--data` (JSON 2D array)
- [ ] Validate data against schema (`write_data.schema.json`)
- [ ] Use `ExcelAgent` context manager
- [ ] Parse range with `RangeSerializer`
- [ ] Iterate data rows and write to cells
- [ ] Infer types: if string is ISO date → convert to datetime
- [ ] If boolean string ("true"/"false") → convert to bool
- [ ] If numeric string → convert to number
- [ ] Handle formula strings starting with `=` (set as formula, not value)
- [ ] Save workbook
- [ ] Return JSON with `cells_modified` count
- [ ] Unit test: write 10x10 array, read back, verify roundtrip
- [ ] Unit test: date string "2026-04-08" is stored as datetime
- [ ] Unit test: formula string "=SUM(A1:A10)" is stored as formula

---

### 38. `tools/write/xls_write_cell.py`
**Purpose:** Write single cell with explicit type coercion.

**CLI Interface:**
```bash
xls_write_cell.py --input sample.xlsx --output output.xlsx \
  --cell A1 --value "2026-04-08" --type date
```

**Features:**
- Explicit type: `string`, `number`, `boolean`, `date`, `formula`
- Overrides inference

**Checklist:**
- [ ] Parse `--cell`, `--value`, `--type`
- [ ] Coerce value based on `--type`
- [ ] If type=date, parse ISO string to datetime
- [ ] If type=formula, ensure value starts with `=`
- [ ] Write to cell
- [ ] Unit test: type=date converts string correctly
- [ ] Unit test: type=formula sets cell.value correctly

---

### 39. `tests/integration/test_write_tools.py`
**Purpose:** Integration tests for write tools.

**Test Cases:**
- [ ] Create new workbook, verify it exists and has correct sheets
- [ ] Write range, read back, verify data matches
- [ ] Write cell with type=date, read back, verify it's datetime
- [ ] Create from template with variable substitution, verify replacement

---

**Phase 5 Exit Criteria:**
- [ ] All 4 write/create tools execute successfully
- [ ] Roundtrip test: write data, read data, verify equality
- [ ] Type inference correctly handles dates, booleans, numbers
- [ ] Template variable substitution works for {{placeholder}} syntax
- [ ] All unit and integration tests pass

---

# Phase 6: Structural Mutation Tools (Category 4)
**Duration:** 8 days  
**Goal:** Implement sheet/row/column manipulation with dependency checks and token enforcement.

## Files to Create

### 40. `tools/structure/xls_add_sheet.py`
**Purpose:** Add a new sheet to workbook.

**CLI Interface:**
```bash
xls_add_sheet.py --input sample.xlsx --output output.xlsx \
  --name "NewSheet" --position after:Sheet1
```

**Features:**
- Insert at specific position (before/after another sheet, or at index)

**Checklist:**
- [ ] Parse `--name`, `--position` (format: `before:SheetName`, `after:SheetName`, or `index:0`)
- [ ] Use `ExcelAgent`
- [ ] Create new sheet with `workbook.create_sheet()`
- [ ] Determine insert index based on position
- [ ] Save workbook
- [ ] Unit test: sheet added at correct position

---

### 41. `tools/structure/xls_delete_sheet.py` ⚠️
**Purpose:** Delete a sheet (requires approval token).

**CLI Interface:**
```bash
xls_delete_sheet.py --input sample.xlsx --output output.xlsx \
  --name "Sheet2" --token <approval_token>
```

**Features:**
- Pre-flight dependency check (cross-sheet formula scan)
- Approval token validation (scope: `sheet:delete`)
- Impact report if references exist

**Checklist:**
- [ ] Parse `--name`, `--token`
- [ ] Validate token with `ApprovalTokenManager.validate_token(scope='sheet:delete', resource=name)`
- [ ] Use `DependencyTracker.impact_report(target=name, action='delete')`
- [ ] If `broken_references > 0`, return warning JSON with impact
- [ ] If `--force` flag is present and token valid, proceed with deletion
- [ ] Log operation to audit trail
- [ ] Delete sheet with `workbook.remove(sheet)`
- [ ] Unit test: delete sheet with valid token succeeds
- [ ] Unit test: delete sheet without token fails with exit code 4
- [ ] Unit test: delete sheet with cross-sheet refs shows warning

---

### 42. `tools/structure/xls_rename_sheet.py` ⚠️
**Purpose:** Rename sheet and update all cross-sheet references.

**CLI Interface:**
```bash
xls_rename_sheet.py --input sample.xlsx --output output.xlsx \
  --old-name "Sheet1" --new-name "Data" --token <approval_token>
```

**Features:**
- Auto-updates all formulas referencing the sheet
- Token required due to formula side-effects

**Checklist:**
- [ ] Parse `--old-name`, `--new-name`, `--token`
- [ ] Validate token (scope: `sheet:rename`)
- [ ] Rename sheet: `sheet.title = new_name`
- [ ] Iterate all cells in all sheets
- [ ] For each formula, use openpyxl Tokenizer to find `'SheetName'!` references
- [ ] Replace old sheet name with new name in formula strings
- [ ] Update defined names referencing the sheet
- [ ] Unit test: rename sheet, verify formulas updated
- [ ] Unit test: formula `'Sheet1'!A1` becomes `'Data'!A1` after rename

---

### 43. `tools/structure/xls_insert_rows.py`
**Purpose:** Insert rows with style inheritance and formula offset updating.

**CLI Interface:**
```bash
xls_insert_rows.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --before-row 5 --count 3
```

**Features:**
- Insert blank rows before specified row
- Copy styles from adjacent row (optional)
- Auto-update formula references (A1 → A4 if inserting 3 rows above)

**Checklist:**
- [ ] Parse `--sheet`, `--before-row`, `--count`
- [ ] Use `sheet.insert_rows(idx=before_row, amount=count)`
- [ ] Copy styles from row `before_row` to new rows (optional: `--copy-style` flag)
- [ ] Formulas are automatically adjusted by openpyxl
- [ ] Unit test: insert 3 rows, verify formulas shifted correctly
- [ ] Unit test: styles copied if `--copy-style` flag present

---

### 44. `tools/structure/xls_delete_rows.py` ⚠️
**Purpose:** Delete rows with pre-flight impact report.

**CLI Interface:**
```bash
xls_delete_rows.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --start-row 5 --count 3 --token <token>
```

**Features:**
- Dependency check for formulas referencing deleted rows
- Impact report before deletion
- Token enforcement

**Checklist:**
- [ ] Parse `--sheet`, `--start-row`, `--count`, `--token`
- [ ] Validate token (scope: `range:delete`)
- [ ] Build target range: `Sheet1!A5:XFD7` (full rows)
- [ ] Use `DependencyTracker.impact_report()`
- [ ] If impact > 0 and no `--force`, return warning
- [ ] If `--force` and valid token, delete rows with `sheet.delete_rows()`
- [ ] Log to audit trail
- [ ] Unit test: delete rows with valid token
- [ ] Unit test: delete rows causing broken refs returns warning

---

### 45. `tools/structure/xls_insert_columns.py`
**Purpose:** Insert columns (similar to insert_rows).

**Checklist:**
- [ ] Parse `--before-column` (A, B, C, or column number)
- [ ] Use `sheet.insert_cols()`
- [ ] Copy styles from adjacent column
- [ ] Formulas auto-adjust

---

### 46. `tools/structure/xls_delete_columns.py` ⚠️
**Purpose:** Delete columns (similar to delete_rows).

**Checklist:**
- [ ] Parse `--start-column`, `--count`, `--token`
- [ ] Validate token (scope: `range:delete`)
- [ ] Dependency check
- [ ] Delete with `sheet.delete_cols()`
- [ ] Audit log

---

### 47. `tools/structure/xls_move_sheet.py`
**Purpose:** Reorder sheet position.

**CLI Interface:**
```bash
xls_move_sheet.py --input sample.xlsx --output output.xlsx \
  --name "Sheet2" --position 0
```

**Checklist:**
- [ ] Parse `--name`, `--position` (integer index)
- [ ] Use `workbook.move_sheet(sheet, offset=...)`
- [ ] Unit test: move sheet from position 2 to 0

---

### 48. `tests/integration/test_structure_tools.py`
**Purpose:** Integration tests for structural tools.

**Test Cases:**
- [ ] Add sheet, verify it appears in sheet list
- [ ] Delete sheet with token, verify removal
- [ ] Delete sheet without token, verify exit code 4
- [ ] Rename sheet, verify formulas updated (e.g., `'OldName'!A1` → `'NewName'!A1`)
- [ ] Insert rows, verify formula offsets updated (A1 → A4 after inserting 3 rows above)
- [ ] Delete rows with broken references, verify impact report returned
- [ ] Move sheet, verify new position

---

**Phase 6 Exit Criteria:**
- [ ] All 8 structural tools execute successfully
- [ ] Approval token enforcement works (invalid token → exit code 4)
- [ ] Dependency impact reports correctly identify broken references
- [ ] Formula offsets update correctly after insert/delete operations
- [ ] Audit trail logs all destructive operations
- [ ] Cross-sheet formula references update correctly after sheet rename
- [ ] All unit and integration tests pass with >90% coverage

---

# Phase 7: Cell Operations (Category 5)
**Duration:** 3 days  
**Goal:** Implement cell-level operations (merge, unmerge, delete range, update references).

## Files to Create

### 49. `tools/cells/xls_merge_cells.py`
**Purpose:** Merge a range of cells.

**CLI Interface:**
```bash
xls_merge_cells.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --range A1:C1
```

**Features:**
- Pre-check: warn if merging cells with data in non-anchor cells (data loss)

**Checklist:**
- [ ] Parse `--range`, `--sheet`
- [ ] Use `RangeSerializer` to get coordinates
- [ ] Check cells B1, C1 (non-anchor) for non-empty values
- [ ] If non-empty, return warning or fail (unless `--force`)
- [ ] Merge with `sheet.merge_cells(range_string)`
- [ ] Unit test: merge empty range succeeds
- [ ] Unit test: merge range with data in B1 returns warning

---

### 50. `tools/cells/xls_unmerge_cells.py`
**Purpose:** Unmerge cells.

**CLI Interface:**
```bash
xls_unmerge_cells.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --range A1:C1
```

**Checklist:**
- [ ] Parse `--range`, `--sheet`
- [ ] Unmerge with `sheet.unmerge_cells(range_string)`
- [ ] Unit test: unmerge restores grid

---

### 51. `tools/cells/xls_delete_range.py` ⚠️
**Purpose:** Delete range and shift cells.

**CLI Interface:**
```bash
xls_delete_range.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --range A1:C10 --shift up --token <token>
```

**Features:**
- Shift direction: `up`, `left`
- Token enforcement
- Impact report

**Checklist:**
- [ ] Parse `--range`, `--shift`, `--token`
- [ ] Validate token (scope: `range:delete`)
- [ ] Dependency check for deleted cells
- [ ] Delete range: iterate cells, set to None, then shift
- [ ] openpyxl limitation: manual shift implementation (move cells)
- [ ] Unit test: delete range with shift up

---

### 52. `tools/cells/xls_update_references.py`
**Purpose:** Batch-update cell references after structural changes.

**CLI Interface:**
```bash
xls_update_references.py --input sample.xlsx --output output.xlsx \
  --updates '[{"old": "A1", "new": "B5"}, {"old": "Sheet1!C1", "new": "Sheet2!D10"}]'
```

**Features:**
- AI can use this to fix broken references after deletions

**Checklist:**
- [ ] Parse `--updates` (JSON array of old/new pairs)
- [ ] Iterate all cells with formulas
- [ ] For each formula, tokenize and replace old references with new
- [ ] Update defined names
- [ ] Unit test: update A1 to B5 in formula `=A1*2` becomes `=B5*2`

---

### 53. `tests/integration/test_cell_operations.py`
**Purpose:** Integration tests for cell operations.

**Test Cases:**
- [ ] Merge cells, read back merged range
- [ ] Unmerge cells, verify individual cells
- [ ] Delete range with shift, verify cells moved
- [ ] Update references, verify formulas changed

---

**Phase 7 Exit Criteria:**
- [ ] Merge/unmerge tools work correctly
- [ ] Delete range with shift updates adjacent cells
- [ ] Update references tool fixes broken formulas
- [ ] All tests pass

---

# Phase 8: Formulas & Calculation (Category 6)
**Duration:** 4 days  
**Goal:** Implement formula manipulation and two-tier calculation engine.

## Files to Create

### 54. `src/excel_agent/calculation/tier1_engine.py`
**Purpose:** In-process calculation using `formulas` library.

**Interface:**
```python
class Tier1Calculator:
    def __init__(self, workbook_path: Path):
        """Loads workbook into formulas ExcelModel."""
        ...
    
    def calculate(self, recalculate_all: bool = True) -> dict:
        """Calculates formulas. Returns results and errors."""
        ...
    
    def get_cell_value(self, cell: str) -> Any:
        """Returns calculated value for a cell (e.g., 'Sheet1!A1')."""
        ...
```

**Checklist:**
- [ ] Use `formulas.ExcelModel().loads(workbook_path)`
- [ ] Call `model.calculate()` to evaluate
- [ ] Handle unsupported functions gracefully (return error marker)
- [ ] Unit test: workbook with SUM, IF, AVERAGE formulas
- [ ] Unit test: unsupported function (e.g., XLOOKUP) returns error

---

### 55. `src/excel_agent/calculation/tier2_libreoffice.py`
**Purpose:** LibreOffice headless recalculation wrapper.

**Interface:**
```python
class Tier2Calculator:
    def __init__(self, soffice_path: Optional[Path] = None):
        """Initializes with path to soffice binary."""
        ...
    
    def recalculate(self, workbook_path: Path, timeout: int = 60) -> dict:
        """Forces full recalculation via LibreOffice.
        Returns: formula_count, error_count, recalc_time_ms"""
        ...
```

**Implementation:**
- Call `soffice --headless --calc --convert-to xlsx <file>` (forces recalc on open)
- Or use a LibreOffice macro that opens, recalculates, saves, closes
- Parse stderr for errors
- Return JSON with stats

**Checklist:**
- [ ] Detect `soffice` binary (common paths: `/usr/bin/soffice`, `C:\Program Files\LibreOffice\...`)
- [ ] Execute headless command with `subprocess.run(timeout=timeout)`
- [ ] Capture stdout/stderr
- [ ] Parse output for error indicators
- [ ] Measure execution time
- [ ] Unit test: recalc sample.xlsx via LibreOffice
- [ ] Unit test: timeout after 60s if file is too large

---

### 56. `scripts/recalc.py`
**Purpose:** LibreOffice Python macro for recalculation.

**Features:**
- Minimal Python script that can be executed by LibreOffice's Python interpreter
- Opens workbook, calls `calculate()`, saves, exits

**Script:**
```python
#!/usr/bin/env python3
import sys
import uno
from com.sun.star.beans import PropertyValue

def recalc_workbook(file_path):
    # Connect to LibreOffice
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    
    # Load file
    url = uno.systemPathToFileUrl(file_path)
    doc = desktop.loadComponentFromURL(url, "_blank", 0, ())
    
    # Force recalculation
    doc.calculateAll()
    
    # Save and close
    doc.store()
    doc.close(True)

if __name__ == "__main__":
    recalc_workbook(sys.argv[1])
```

**Checklist:**
- [ ] Script accepts file path as argument
- [ ] Connects to LibreOffice UNO bridge (requires `soffice --accept` running)
- [ ] Calls `calculateAll()` on spreadsheet document
- [ ] Saves and closes
- [ ] Error handling for file not found, connection failures

**Alternative:** Use simpler headless conversion approach if UNO is too complex.

---

### 57. `src/excel_agent/calculation/error_detector.py`
**Purpose:** Scan workbook for formula errors.

**Interface:**
```python
def detect_errors(workbook: Workbook) -> List[dict]:
    """Scans all cells for error values.
    Returns list of errors: [{cell: 'Sheet1!A1', error: '#REF!'}]"""
    ...
```

**Checklist:**
- [ ] Iterate all sheets and cells
- [ ] Check if `cell.value` is an error string (`#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`)
- [ ] Return list of error locations
- [ ] Unit test: workbook with `#REF!` error is detected

---

### 58. `tools/formulas/xls_set_formula.py`
**Purpose:** Set formula in a cell with validation.

**CLI Interface:**
```bash
xls_set_formula.py --input sample.xlsx --output output.xlsx \
  --cell A1 --formula "=SUM(B1:B10)"
```

**Features:**
- Syntax validation using openpyxl Tokenizer
- Detect invalid cell references

**Checklist:**
- [ ] Parse `--cell`, `--formula`
- [ ] Ensure formula starts with `=`
- [ ] Tokenize formula to validate syntax
- [ ] Set `cell.value = formula`
- [ ] Unit test: valid formula sets correctly
- [ ] Unit test: invalid formula (e.g., `=SUM(`) returns validation error

---

### 59. `tools/formulas/xls_recalculate.py`
**Purpose:** Force recalculation using two-tier strategy.

**CLI Interface:**
```bash
xls_recalculate.py --input sample.xlsx --output output.xlsx --tier 1
# Or force Tier 2:
xls_recalculate.py --input sample.xlsx --output output.xlsx --tier 2
```

**Features:**
- Default: Try Tier 1, fall back to Tier 2 if errors
- Explicit tier selection via `--tier` flag

**Checklist:**
- [ ] Parse `--tier` (optional, defaults to auto)
- [ ] If tier=1 or auto: use `Tier1Calculator`
- [ ] If unsupported functions detected and auto: fall back to `Tier2Calculator`
- [ ] If tier=2: use `Tier2Calculator`
- [ ] Return JSON with `formula_count`, `error_count`, `recalc_time_ms`, `engine_used`
- [ ] Unit test: Tier 1 calculates SUM correctly
- [ ] Unit test: Tier 2 fallback for unsupported function

---

### 60. `tools/formulas/xls_detect_errors.py`
**Purpose:** Scan for formula errors.

**Output:**
```json
{
  "status": "success",
  "data": {
    "errors": [
      {"sheet": "Sheet1", "cell": "A1", "error": "#REF!"},
      {"sheet": "Sheet2", "cell": "B5", "error": "#DIV/0!"}
    ],
    "error_count": 2
  }
}
```

**Checklist:**
- [ ] Use `error_detector.detect_errors()`
- [ ] Return list of errors
- [ ] Unit test: workbook with multiple error types

---

### 61. `tools/formulas/xls_convert_to_values.py` ⚠️
**Purpose:** Replace formulas with calculated values (irreversible).

**CLI Interface:**
```bash
xls_convert_to_values.py --input sample.xlsx --output output.xlsx \
  --range A1:C10 --token <token>
```

**Features:**
- Token enforcement (scope: `formula:convert`)
- Warning about irreversibility

**Checklist:**
- [ ] Parse `--range`, `--token`
- [ ] Validate token
- [ ] For each cell in range, if formula: replace with calculated value
- [ ] Use Tier 1 or Tier 2 to get values first
- [ ] Save workbook
- [ ] Audit log
- [ ] Unit test: formula `=2+2` becomes value `4`

---

### 62. `tools/formulas/xls_copy_formula_down.py`
**Purpose:** Auto-fill formula down a column.

**CLI Interface:**
```bash
xls_copy_formula_down.py --input sample.xlsx --output output.xlsx \
  --source A1 --target A2:A100
```

**Features:**
- Auto-adjust relative references

**Checklist:**
- [ ] Parse `--source` (single cell with formula), `--target` (range)
- [ ] Get formula from source cell
- [ ] For each row in target, copy formula with adjusted row references
- [ ] openpyxl handles relative reference adjustment automatically
- [ ] Unit test: `=B1*2` in A1 copied to A2 becomes `=B2*2`

---

### 63. `tools/formulas/xls_define_name.py`
**Purpose:** Create or update named range.

**CLI Interface:**
```bash
xls_define_name.py --input sample.xlsx --output output.xlsx \
  --name "SalesData" --refers-to "Sheet1!$A$1:$C$100" --scope Workbook
```

**Checklist:**
- [ ] Parse `--name`, `--refers-to`, `--scope` (Workbook or SheetName)
- [ ] Create `DefinedName` object
- [ ] Add to `workbook.defined_names`
- [ ] Unit test: defined name resolves correctly

---

### 64. `tests/integration/test_calculation.py`
**Purpose:** Integration tests for calculation tools.

**Test Cases:**
- [ ] Set formula, recalculate, verify result
- [ ] Detect errors in workbook with `#REF!`
- [ ] Convert formulas to values, verify formulas removed
- [ ] Copy formula down, verify relative references adjusted
- [ ] Tier 1 calculation matches Excel (5 decimal places)
- [ ] Tier 2 calculation via LibreOffice completes successfully

---

**Phase 8 Exit Criteria:**
- [ ] Tier 1 calculation works for common functions (SUM, AVERAGE, IF)
- [ ] Tier 2 LibreOffice bridge recalculates workbook successfully
- [ ] Formula error detection finds all error types
- [ ] Convert to values tool removes formulas correctly
- [ ] All tests pass
- [ ] LibreOffice installed and working in CI environment

---

# Phase 9: Macro Safety Tools (Category 9)
**Duration:** 3 days  
**Goal:** Implement VBA inspection and safe macro management using `oletools`.

## Files to Create

### 65. `src/excel_agent/core/macro_handler.py`
**Purpose:** Safe VBA inspection and container management.

**Interface:**
```python
class MacroHandler:
    def has_vba_project(self, path: Path) -> bool:
        """Uses oletools.olevba to detect VBA macros."""
        ...
    
    def get_vba_modules(self, path: Path) -> List[dict]:
        """Returns list of VBA modules (name, type, size).
        Uses olevba.extract_macros()."""
        ...
    
    def scan_risk(self, path: Path) -> dict:
        """Returns security risk assessment.
        Uses olevba.detect_autoexec() and detect_suspicious()."""
        ...
    
    def has_digital_signature(self, path: Path) -> bool:
        """Checks for vbaProjectSignature.bin in OLE structure."""
        ...
    
    def extract_vba_project(self, xlsm_path: Path, output_bin: Path) -> None:
        """Extracts vbaProject.bin for safe injection elsewhere."""
        ...
```

**Checklist:**
- [ ] Import `oletools.olevba`
- [ ] `has_vba_project()`: use `olevba.detect_vba_macros()`
- [ ] `get_vba_modules()`: use `olevba.extract_macros()`, parse module list
- [ ] `scan_risk()`: use `detect_autoexec()`, `detect_suspicious()`
- [ ] `has_digital_signature()`: use `olefile` to check for signature stream
- [ ] `extract_vba_project()`: use `zipfile` to extract `xl/vbaProject.bin`
- [ ] Unit test: .xlsm with macros detected
- [ ] Unit test: .xlsx without macros returns False
- [ ] Unit test: auto-exec trigger (AutoOpen) detected in scan_risk
- [ ] Unit test: suspicious keyword (Shell) detected

---

### 66. `tools/macros/xls_has_macros.py`
**Purpose:** Boolean check for VBA presence.

**Output:**
```json
{
  "status": "success",
  "data": {
    "has_macros": true,
    "file_type": "xlsm"
  }
}
```

**Checklist:**
- [ ] Use `MacroHandler.has_vba_project()`
- [ ] Return boolean
- [ ] Unit test: .xlsm with macros
- [ ] Unit test: .xlsx without macros

---

### 67. `tools/macros/xls_inspect_macros.py`
**Purpose:** List VBA modules and signature status.

**Output:**
```json
{
  "status": "success",
  "data": {
    "modules": [
      {"name": "Module1", "type": "Standard", "lines": 45},
      {"name": "ThisWorkbook", "type": "Class", "lines": 12}
    ],
    "digitally_signed": true
  }
}
```

**Checklist:**
- [ ] Use `MacroHandler.get_vba_modules()`, `has_digital_signature()`
- [ ] Return module list and signature status
- [ ] Unit test: .xlsm with 2 modules

---

### 68. `tools/macros/xls_validate_macro_safety.py`
**Purpose:** Security risk scan.

**Output:**
```json
{
  "status": "warning",
  "data": {
    "risk_level": "high",
    "auto_executable": ["AutoOpen", "Workbook_Open"],
    "suspicious_keywords": ["Shell", "CreateObject"],
    "iocs": ["http://malicious.com/payload.exe"],
    "recommendation": "Manual review required before enabling macros."
  }
}
```

**Checklist:**
- [ ] Use `MacroHandler.scan_risk()`
- [ ] Return risk level: low/medium/high
- [ ] List auto-exec triggers, suspicious keywords, IOCs
- [ ] Unit test: .xlsm with AutoOpen and Shell keyword

---

### 69. `tools/macros/xls_remove_macros.py` ⚠️⚠️
**Purpose:** Strip VBA project (double-token required).

**CLI Interface:**
```bash
xls_remove_macros.py --input sample.xlsm --output sample_clean.xlsx \
  --token <token1> --confirm-token <token2>
```

**Features:**
- Requires two separate tokens (paranoid safety)
- Converts .xlsm to .xlsx
- Removes vbaProject.bin from ZIP archive

**Checklist:**
- [ ] Parse `--token`, `--confirm-token`
- [ ] Validate both tokens (scope: `macro:remove`)
- [ ] Open .xlsm with `zipfile`
- [ ] Remove `xl/vbaProject.bin` and `xl/vbaProjectSignature.bin` (if exists)
- [ ] Save as .xlsx (change extension)
- [ ] Audit log
- [ ] Unit test: .xlsm becomes .xlsx without VBA

---

### 70. `tools/macros/xls_inject_vba_project.py` ⚠️
**Purpose:** Inject pre-extracted vbaProject.bin.

**CLI Interface:**
```bash
xls_inject_vba_project.py --input sample.xlsx --output sample.xlsm \
  --vba-bin project.bin --signature signature.bin --token <token>
```

**Features:**
- Inject trusted VBA project
- Optionally inject signature for signed macros
- Token enforcement

**Checklist:**
- [ ] Parse `--vba-bin`, `--signature` (optional), `--token`
- [ ] Validate token (scope: `macro:inject`)
- [ ] Open .xlsx with `zipfile`
- [ ] Add `xl/vbaProject.bin` to archive
- [ ] If signature provided, add `xl/vbaProjectSignature.bin`
- [ ] Update `[Content_Types].xml` and workbook.xml to mark as macro-enabled
- [ ] Save as .xlsm
- [ ] Unit test: .xlsx + vbaProject.bin = .xlsm with macros

---

### 71. `tests/integration/test_macro_tools.py`
**Purpose:** Integration tests for macro tools.

**Test Cases:**
- [ ] Detect macros in .xlsm
- [ ] Inspect modules in .xlsm
- [ ] Validate macro safety, detect AutoOpen
- [ ] Remove macros with double token, verify .xlsx output
- [ ] Inject VBA project into .xlsx, verify .xlsm output with macros

---

**Phase 9 Exit Criteria:**
- [ ] All 5 macro tools execute successfully
- [ ] `oletools` integration correctly detects auto-exec and suspicious keywords
- [ ] Digital signature detection works
- [ ] Macro removal produces clean .xlsx file
- [ ] Macro injection produces functional .xlsm file
- [ ] All tests pass with 100% coverage for macro_handler.py

---

# Phase 10: Objects & Visualization (Category 7)
**Duration:** 5 days  
**Goal:** Implement chart, table, image, comment, and data validation tools.

## Files to Create

### 72. `tools/objects/xls_add_table.py`
**Purpose:** Convert range to Excel Table (ListObject).

**CLI Interface:**
```bash
xls_add_table.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --range A1:D100 --name "SalesTable" --style TableStyleMedium2
```

**Features:**
- Apply table style
- Auto-detect headers

**Checklist:**
- [ ] Parse `--range`, `--name`, `--style`
- [ ] Use openpyxl `Table` class
- [ ] Create table: `table = Table(displayName=name, ref=range_string)`
- [ ] Set style: `table.tableStyleInfo = TableStyleInfo(name=style)`
- [ ] Add to sheet: `sheet.add_table(table)`
- [ ] Unit test: created table appears in `sheet.tables`

---

### 73. `tools/objects/xls_add_chart.py`
**Purpose:** Create chart from data range.

**CLI Interface:**
```bash
xls_add_chart.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --chart-type bar --data-range A1:B10 \
  --position E5 --title "Sales by Region"
```

**Features:**
- Chart types: bar, line, pie, scatter
- Anchor to cell position

**Checklist:**
- [ ] Parse `--chart-type`, `--data-range`, `--position`, `--title`
- [ ] Use openpyxl chart classes (`BarChart`, `LineChart`, `PieChart`, `ScatterChart`)
- [ ] Create chart object, set data reference
- [ ] Set title
- [ ] Anchor chart to cell with `sheet.add_chart(chart, position)`
- [ ] Unit test: chart appears in workbook, opens without errors

---

### 74. `tools/objects/xls_add_image.py`
**Purpose:** Insert image into sheet.

**CLI Interface:**
```bash
xls_add_image.py --input sample.xlsx --output output.xlsx \
  --sheet Sheet1 --image logo.png --position A1 --lock-aspect
```

**Features:**
- Preserve aspect ratio
- Anchor to cell

**Checklist:**
- [ ] Parse `--image`, `--position`, `--lock-aspect` (boolean flag)
- [ ] Use openpyxl `Image` class
- [ ] Load image: `img = Image(image_path)`
- [ ] Anchor: `img.anchor = position`
- [ ] Add to sheet: `sheet.add_image(img)`
- [ ] Unit test: image embedded in workbook

---

### 75. `tools/objects/xls_add_comment.py`
**Purpose:** Add threaded comment to cell.

**CLI Interface:**
```bash
xls_add_comment.py --input sample.xlsx --output output.xlsx \
  --cell A1 --text "Review this value" --author "Agent"
```

**Checklist:**
- [ ] Parse `--cell`, `--text`, `--author`
- [ ] Use openpyxl `Comment` class
- [ ] Create comment: `comment = Comment(text, author)`
- [ ] Assign to cell: `cell.comment = comment`
- [ ] Unit test: comment appears in Excel

---

### 76. `tools/objects/xls_set_data_validation.py`
**Purpose:** Add dropdown list or numeric constraint.

**CLI Interface:**
```bash
xls_set_data_validation.py --input sample.xlsx --output output.xlsx \
  --range A1:A10 --type list --formula '"Option1,Option2,Option3"'
```

**Features:**
- Validation types: list, whole, decimal, date, textLength
- Custom formulas

**Checklist:**
- [ ] Parse `--range`, `--type`, `--formula`
- [ ] Use openpyxl `DataValidation` class
- [ ] Create validation: `dv = DataValidation(type=type, formula1=formula)`
- [ ] Add to sheet: `sheet.add_data_validation(dv)`
- [ ] Apply to range: `dv.add(range_string)`
- [ ] Unit test: dropdown list appears in Excel

---

### 77. `tests/integration/test_objects.py`
**Purpose:** Integration tests for object tools.

**Test Cases:**
- [ ] Add table, verify it exists in `sheet.tables`
- [ ] Add chart, open workbook in Excel, verify chart renders
- [ ] Add image, verify image embedded
- [ ] Add comment, verify comment text
- [ ] Add data validation, verify dropdown works

---

**Phase 10 Exit Criteria:**
- [ ] All 5 object tools execute successfully
- [ ] Generated workbooks open in Excel without repair prompts
- [ ] Charts, images, tables render correctly
- [ ] Data validation dropdowns work
- [ ] All tests pass

---

# Phase 11: Formatting & Style (Category 8)
**Duration:** 3 days  
**Goal:** Implement cell and range formatting tools.

## Files to Create

### 78. `tools/formatting/xls_format_range.py`
**Purpose:** Apply comprehensive formatting from JSON spec.

**CLI Interface:**
```bash
xls_format_range.py --input sample.xlsx --output output.xlsx \
  --range A1:C10 --style '{
    "font": {"name": "Arial", "size": 12, "bold": true, "color": "FF0000"},
    "fill": {"fgColor": "FFFF00", "patternType": "solid"},
    "alignment": {"horizontal": "center", "vertical": "middle"}
  }'
```

**Checklist:**
- [ ] Parse `--style` (JSON object)
- [ ] Validate against style schema
- [ ] Use openpyxl `Font`, `Fill`, `Border`, `Alignment` classes
- [ ] Apply to all cells in range
- [ ] Unit test: formatting applied correctly

---

### 79. `tools/formatting/xls_set_column_width.py`
**Purpose:** Set column width.

**CLI Interface:**
```bash
xls_set_column_width.py --input sample.xlsx --output output.xlsx \
  --columns "A,B,C" --width auto
# Or fixed width:
xls_set_column_width.py --columns "A" --width 15
```

**Checklist:**
- [ ] Parse `--columns`, `--width` (auto or numeric)
- [ ] If auto: sample first 100 rows, calculate max content width
- [ ] Set `sheet.column_dimensions[col].width`
- [ ] Unit test: auto-width sets reasonable value
- [ ] Unit test: fixed width sets exact value

---

### 80. `tools/formatting/xls_freeze_panes.py`
**Purpose:** Freeze rows/columns for scrolling.

**CLI Interface:**
```bash
xls_freeze_panes.py --input sample.xlsx --output output.xlsx \
  --cell A2  # Freezes row 1
```

**Checklist:**
- [ ] Parse `--cell` (freeze point)
- [ ] Set `sheet.freeze_panes = cell`
- [ ] Unit test: freeze panes set correctly

---

### 81. `tools/formatting/xls_apply_conditional_formatting.py`
**Purpose:** Add conditional formatting rules.

**CLI Interface:**
```bash
xls_apply_conditional_formatting.py --input sample.xlsx --output output.xlsx \
  --range A1:A10 --type colorScale --colors '["FF0000", "FFFF00", "00FF00"]'
```

**Features:**
- ColorScale, DataBar, IconSet

**Checklist:**
- [ ] Parse `--type`, rule-specific parameters
- [ ] Use openpyxl conditional formatting classes
- [ ] Add rule to sheet
- [ ] Unit test: rule appears in Excel

---

### 82. `tools/formatting/xls_set_number_format.py`
**Purpose:** Apply number format codes.

**CLI Interface:**
```bash
xls_set_number_format.py --input sample.xlsx --output output.xlsx \
  --range A1:A10 --format "$#,##0.00"
```

**Checklist:**
- [ ] Parse `--format` (Excel format code)
- [ ] Set `cell.number_format`
- [ ] Presets: `currency`, `percentage`, `date`, `time`
- [ ] Unit test: currency format displays with $

---

### 83. `tests/integration/test_formatting.py`
**Purpose:** Integration tests for formatting tools.

**Test Cases:**
- [ ] Format range with bold red text, verify in Excel
- [ ] Set column width to auto, verify reasonable width
- [ ] Freeze panes at A2, verify row 1 frozen
- [ ] Apply color scale, verify gradient in Excel
- [ ] Set number format to currency, verify $ symbol

---

**Phase 11 Exit Criteria:**
- [ ] All 5 formatting tools execute successfully
- [ ] Formatting persists when opening in Excel
- [ ] Auto-width calculates reasonable values
- [ ] Conditional formatting renders correctly
- [ ] All tests pass

---

# Phase 12: Export & Interop (Category 10)
**Duration:** 2 days  
**Goal:** Implement export tools (PDF, CSV, JSON).

## Files to Create

### 84. `tools/export/xls_export_pdf.py`
**Purpose:** Export workbook or sheet to PDF via LibreOffice.

**CLI Interface:**
```bash
xls_export_pdf.py --input sample.xlsx --output sample.pdf --sheet Sheet1
```

**Features:**
- Uses LibreOffice headless for conversion

**Checklist:**
- [ ] Parse `--sheet` (optional, exports all if not specified)
- [ ] Call `soffice --headless --convert-to pdf --outdir <dir> <file>`
- [ ] If `--sheet` specified, hide other sheets before export, then unhide
- [ ] Unit test: PDF generated successfully
- [ ] Unit test: multi-sheet workbook exports all sheets

---

### 85. `tools/export/xls_export_csv.py`
**Purpose:** Export sheet to CSV.

**CLI Interface:**
```bash
xls_export_csv.py --input sample.xlsx --output data.csv --sheet Sheet1 --encoding utf-8
```

**Checklist:**
- [ ] Parse `--sheet`, `--encoding`
- [ ] Read all rows from sheet
- [ ] Write to CSV using `csv.writer`
- [ ] Handle cell types (convert dates to ISO strings)
- [ ] Unit test: CSV roundtrip (write to Excel, export to CSV, verify match)

---

### 86. `tools/export/xls_export_json.py`
**Purpose:** Export sheet or range as structured JSON.

**CLI Interface:**
```bash
xls_export_json.py --input sample.xlsx --output data.json \
  --sheet Sheet1 --range A1:C100 --format records
# Formats: records (array of objects), values (2D array), columns (object of arrays)
```

**Output (records format):**
```json
[
  {"Name": "Alice", "Age": 30, "City": "NYC"},
  {"Name": "Bob", "Age": 25, "City": "LA"}
]
```

**Checklist:**
- [ ] Parse `--format` (records, values, columns)
- [ ] If records: first row is headers, subsequent rows are data objects
- [ ] If values: 2D array (like `xls_read_range`)
- [ ] If columns: `{"Name": ["Alice", "Bob"], "Age": [30, 25]}`
- [ ] Unit test: export as records matches expected structure

---

### 87. `tests/integration/test_export.py`
**Purpose:** Integration tests for export tools.

**Test Cases:**
- [ ] Export to PDF, verify file size > 0 and is valid PDF
- [ ] Export to CSV, verify correct encoding and delimiter
- [ ] Export to JSON (records), verify structure matches
- [ ] Export large dataset (100k rows) to CSV in <5s

---

**Phase 12 Exit Criteria:**
- [ ] All 3 export tools execute successfully
- [ ] PDF export produces valid PDF files
- [ ] CSV export handles special characters and encoding
- [ ] JSON export produces valid, parseable JSON
- [ ] All tests pass

---

# Phase 13: Integration Testing & Documentation
**Duration:** 3 days  
**Goal:** E2E workflows, comprehensive documentation, and agent simulation tests.

## Files to Create

### 88. `tests/integration/test_clone_modify_workflow.py`
**Purpose:** Simulate full AI agent workflow.

**Workflow:**
1. Clone workbook
2. Read metadata
3. Modify data (write range)
4. Insert rows
5. Recalculate formulas
6. Validate workbook
7. Export to PDF

**Checklist:**
- [ ] Use `subprocess` to call tools sequentially
- [ ] Parse JSON outputs, pass data between tools
- [ ] Verify final workbook state
- [ ] Assert all exit codes are 0
- [ ] Measure total workflow time

---

### 89. `tests/integration/test_formula_dependency_workflow.py`
**Purpose:** Test dependency-aware deletion workflow.

**Workflow:**
1. Load workbook with cross-sheet formulas
2. Generate dependency report
3. Attempt to delete sheet (should warn)
4. Update references
5. Delete sheet successfully

**Checklist:**
- [ ] Simulate agent receiving impact report
- [ ] Agent calls `xls_update_references.py`
- [ ] Agent retries deletion with token
- [ ] Verify no broken references in final workbook

---

### 90. `docs/API.md`
**Purpose:** Complete CLI reference for all 53 tools.

**Structure:**
- Tool name
- Purpose
- CLI interface with all flags
- Input schema
- Output schema
- Exit codes
- Examples

**Checklist:**
- [ ] Document all 53 tools
- [ ] Consistent formatting
- [ ] Examples for common use cases
- [ ] Exit code reference table

---

### 91. `docs/WORKFLOWS.md`
**Purpose:** Common AI agent workflows.

**Workflows:**
1. Clone-Modify-Validate-Export
2. Dependency-Aware Deletion
3. Formula Recalculation with Error Handling
4. Template-Based Report Generation
5. Macro Inspection and Safe Removal

**Checklist:**
- [ ] Step-by-step instructions
- [ ] JSON payloads for each step
- [ ] Expected outputs
- [ ] Error handling guidance

---

### 92. `docs/GOVERNANCE.md`
**Purpose:** Token system and safety protocols.

**Content:**
- Token scope definitions
- Token generation guide
- Audit trail format
- Security best practices

**Checklist:**
- [ ] Document all 7 token scopes
- [ ] Example token generation command
- [ ] Audit log query examples
- [ ] Revocation procedure

---

### 93. `docs/DEVELOPMENT.md`
**Purpose:** Contributor guide.

**Content:**
- Project structure
- Development setup
- Testing strategy
- Code style (black, mypy)
- PR process

**Checklist:**
- [ ] Setup instructions for dev environment
- [ ] How to run tests
- [ ] How to add a new tool
- [ ] Coding standards

---

**Phase 13 Exit Criteria:**
- [ ] All integration tests pass (5+ E2E workflows)
- [ ] Documentation covers all 53 tools
- [ ] Workflow guides enable an AI agent to complete complex tasks
- [ ] Governance documentation is clear and enforceable
- [ ] README links to all documentation

---

# Phase 14: Performance Optimization & Hardening
**Duration:** 3 days  
**Goal:** Optimize for large files, add resilience, final QA.

## Tasks

### 94. Performance Benchmarks
**Purpose:** Measure and optimize critical paths.

**Benchmarks:**
- [ ] Read 500k rows in <3s (chunked mode)
- [ ] Write 500k rows in <5s (chunked mode)
- [ ] Build dependency graph for 10-sheet, 1000-formula workbook in <5s
- [ ] Recalculate 1000 formulas (Tier 1) in <500ms
- [ ] File locking acquire/release in <100ms

**Checklist:**
- [ ] Create `tests/performance/` directory
- [ ] Benchmark script for each critical operation
- [ ] Identify bottlenecks (profiling with `cProfile`)
- [ ] Optimize: use `openpyxl` write-only mode for large writes
- [ ] Document performance characteristics in README

---

### 95. Error Handling Hardening
**Purpose:** Graceful degradation and clear error messages.

**Checklist:**
- [ ] All tools catch exceptions and return exit code 5 with JSON error
- [ ] File not found → exit code 2
- [ ] Lock timeout → exit code 3
- [ ] Invalid token → exit code 4
- [ ] Concurrent modification → exit code 5 with message
- [ ] Malformed JSON input → exit code 1 with validation error details

---

### 96. Security Audit
**Purpose:** Ensure no vulnerabilities in file handling.

**Checklist:**
- [ ] Verify `defusedxml` is used for all XML parsing
- [ ] No arbitrary code execution via formula injection (formulas treated as data)
- [ ] Token secret key is not logged or exposed
- [ ] Audit trail is append-only (permissions check)
- [ ] LibreOffice subprocess runs with timeout (no infinite hangs)
- [ ] File paths validated (no directory traversal attacks)

---

### 97. Cross-Platform Testing
**Purpose:** Ensure tools work on Windows, macOS, Linux.

**Checklist:**
- [ ] CI runs on ubuntu-latest (Linux)
- [ ] Manual test on Windows 11 (WSL or native Python)
- [ ] Manual test on macOS (Homebrew Python)
- [ ] File locking works on all platforms
- [ ] LibreOffice paths detected correctly
- [ ] Path separators handled correctly (use `pathlib.Path`)

---

### 98. Final QA Pass
**Purpose:** Run full test suite, fix any remaining issues.

**Checklist:**
- [ ] Run `pytest` with coverage, achieve >90%
- [ ] Run `black --check .`, fix formatting
- [ ] Run `mypy src/`, fix type errors
- [ ] Run `flake8 src/`, fix linting issues
- [ ] Manually test 10 random tools via CLI
- [ ] Generate coverage report HTML, review untested branches
- [ ] Fix any failing tests

---

**Phase 14 Exit Criteria:**
- [ ] All performance benchmarks meet targets
- [ ] 100% of critical paths have error handling
- [ ] Security audit passes (no vulnerabilities)
- [ ] Tests pass on Linux, Windows, macOS
- [ ] Code coverage >90%
- [ ] No linting or type errors
- [ ] Project is production-ready

---

# Summary: Master Execution Plan Checklist

## Overall Project Phases

- [ ] **Phase 0:** Project Scaffolding & Infrastructure (2 days)
- [ ] **Phase 1:** Core Foundation (5 days)
- [ ] **Phase 2:** Dependency Engine & Validation (5 days)
- [ ] **Phase 3:** Governance & Safety Layer (3 days)
- [ ] **Phase 4:** Read Tools (4 days)
- [ ] **Phase 5:** Write & Create Tools (3 days)
- [ ] **Phase 6:** Structural Mutation Tools (8 days)
- [ ] **Phase 7:** Cell Operations (3 days)
- [ ] **Phase 8:** Formulas & Calculation (4 days)
- [ ] **Phase 9:** Macro Safety Tools (3 days)
- [ ] **Phase 10:** Objects & Visualization (5 days)
- [ ] **Phase 11:** Formatting & Style (3 days)
- [ ] **Phase 12:** Export & Interop (2 days)
- [ ] **Phase 13:** Integration Testing & Documentation (3 days)
- [ ] **Phase 14:** Performance Optimization & Hardening (3 days)

**Total Duration:** ~50 working days (10 weeks)

## Final Deliverables

- [ ] 53 production-ready CLI tools
- [ ] Comprehensive core library (`excel_agent`)
- [ ] Full test coverage (>90%)
- [ ] Complete documentation (API, workflows, governance, development)
- [ ] CI/CD pipeline with automated testing
- [ ] Performance benchmarks demonstrating sub-3s large file reads
- [ ] Security-hardened token system with audit trail
- [ ] Cross-platform compatibility (Linux, Windows, macOS)

---

**This Master Execution Plan provides a comprehensive, phased roadmap for building `excel-agent-tools` to the highest standards of meticulous design, AI-native architecture, and governance-first safety.** Each phase builds upon the previous, ensuring dependencies are respected and allowing for iterative testing and validation. Every file has a clear purpose, interface specification, and success criteria.
