# Phase 2: PLAN — Master Execution Plan

Let me first conduct targeted research to ensure every technical detail in this plan reflects the absolute latest state as of April 2026.Now I have all the research data needed. Let me synthesize everything into the comprehensive Master Execution Plan, incorporating all validated findings and the latest confirmed library versions.

---

# 📋 Master Execution Plan: `excel-agent-tools` v1.0.0

## Research-Validated Technology Stack (As of April 8, 2026)

All versions below are confirmed via direct PyPI/documentation research:

| Component | Package | Confirmed Version | Source |
|:---|:---|:---|:---|
| Core I/O | `openpyxl` | 3.1.5 | PyPI |
| XML Security | `defusedxml` | 0.7.1 | PyPI |
| Macro Inspection | `oletools` | 0.60.2 | PyPI |
| Formula Engine (Tier 1) | `formulas` | 1.3.4 (Released Mar 11, 2026) | PyPI |
| HMAC Security | `hmac` (stdlib) | Uses `compare_digest()` per RFC 2104 | Python 3.14.3 docs |
| Python Floor | ≥3.12 | oletools recommends latest Python 3.x (3.12 for now) | oletools wiki |

**Critical security notes:**
- By default openpyxl does not guard against quadratic blowup or billion laughs xml attacks. To guard against these attacks install defusedxml. → `defusedxml` is **mandatory**.
- `compare_digest` uses an approach designed to prevent timing analysis by avoiding content-based short circuiting behaviour, making it appropriate for cryptography. → All token validation **must** use this.
- The Python XML documentation explicitly recommends the use of the defusedxml package for security purposes. However, this package has not received published updates since early 2021. → Functional but stale; same mitigation strategy as oletools.

---

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
│       │   ├── agent.py                    # ExcelAgent context manager hub
│       │   ├── locking.py                  # Cross-platform OS-level file locking
│       │   ├── serializers.py              # Unified range parsing (A1/R1C1/Name/Table)
│       │   ├── dependency.py               # Formula dependency graph engine
│       │   ├── macro_handler.py            # oletools-based VBA inspection (behind Protocol)
│       │   ├── chunked_io.py               # Streaming I/O for large datasets
│       │   └── version_hash.py             # Geometry-aware workbook hashing
│       ├── governance/
│       │   ├── __init__.py
│       │   ├── token_manager.py            # HMAC-SHA256 scoped approval tokens
│       │   ├── audit_trail.py              # Pluggable audit logging (JSONL default)
│       │   └── schemas/
│       │       ├── __init__.py             # Schema loader utility
│       │       ├── range_input.schema.json
│       │       ├── write_data.schema.json
│       │       ├── style_spec.schema.json
│       │       └── token_request.schema.json
│       ├── calculation/
│       │   ├── __init__.py
│       │   ├── tier1_engine.py             # In-process calc via `formulas` library
│       │   ├── tier2_libreoffice.py        # LibreOffice headless recalc wrapper
│       │   └── error_detector.py           # Formula error scanner
│       └── utils/
│           ├── __init__.py
│           ├── exit_codes.py               # Standardized exit code enum
│           ├── json_io.py                  # JSON response builder & serializer
│           ├── cli_helpers.py              # Reusable CLI argument parsing
│           └── exceptions.py               # Custom exception hierarchy
├── tools/
│   ├── governance/
│   │   ├── xls_clone_workbook.py           # Atomic copy to /work/ directory
│   │   ├── xls_validate_workbook.py        # OOXML compliance & broken ref check
│   │   ├── xls_approve_token.py            # Generate scoped HMAC-SHA256 tokens
│   │   ├── xls_version_hash.py             # Geometry hash (structure + formulas)
│   │   ├── xls_lock_status.py              # Check OS-level file lock state
│   │   └── xls_dependency_report.py        # Full dependency graph export as JSON
│   ├── read/
│   │   ├── xls_read_range.py              # Extract data as JSON, chunked streaming
│   │   ├── xls_get_sheet_names.py         # Sheet index, name, visibility
│   │   ├── xls_get_defined_names.py       # Global and sheet-scoped named ranges
│   │   ├── xls_get_table_info.py          # ListObject schema & metadata
│   │   ├── xls_get_cell_style.py          # Full style inspection as JSON
│   │   ├── xls_get_formula.py             # Formula string + parsed references
│   │   └── xls_get_workbook_metadata.py   # High-level workbook statistics
│   ├── write/
│   │   ├── xls_create_new.py              # Create blank workbook with sheets
│   │   ├── xls_create_from_template.py    # Clone from .xltx/.xltm + variable sub
│   │   ├── xls_write_range.py             # Write 2D data with type inference
│   │   └── xls_write_cell.py              # Single-cell write with explicit type
│   ├── structure/
│   │   ├── xls_add_sheet.py               # Add sheet at position
│   │   ├── xls_delete_sheet.py      ⚠️    # Delete sheet (token + dependency check)
│   │   ├── xls_rename_sheet.py      ⚠️    # Rename + auto-update cross-sheet refs
│   │   ├── xls_insert_rows.py             # Insert rows with style inheritance
│   │   ├── xls_delete_rows.py       ⚠️    # Delete rows (token + impact report)
│   │   ├── xls_insert_columns.py          # Insert columns
│   │   ├── xls_delete_columns.py    ⚠️    # Delete columns (token + impact report)
│   │   └── xls_move_sheet.py              # Reorder sheet position
│   ├── cells/
│   │   ├── xls_merge_cells.py             # Merge with hidden data pre-check
│   │   ├── xls_unmerge_cells.py           # Restore grid from merged range
│   │   ├── xls_delete_range.py      ⚠️    # Shift cells up/left after clearing
│   │   └── xls_update_references.py       # Batch-update cell refs after changes
│   ├── formulas/
│   │   ├── xls_set_formula.py             # Inject formula with syntax validation
│   │   ├── xls_recalculate.py             # Two-tier recalc (formulas → LibreOffice)
│   │   ├── xls_detect_errors.py           # Scan for #REF!, #VALUE!, #DIV/0!, etc.
│   │   ├── xls_convert_to_values.py  ⚠️   # Replace formulas with values (irreversible)
│   │   ├── xls_copy_formula_down.py       # Auto-fill formula with ref adjustment
│   │   └── xls_define_name.py             # Create/update named ranges
│   ├── objects/
│   │   ├── xls_add_table.py               # Convert range to Excel Table
│   │   ├── xls_add_chart.py               # Bar, Line, Pie, Scatter charts
│   │   ├── xls_add_image.py               # Insert image with aspect preservation
│   │   ├── xls_add_comment.py             # Threaded comments
│   │   └── xls_set_data_validation.py     # Dropdown lists, numeric constraints
│   ├── formatting/
│   │   ├── xls_format_range.py            # Fonts, fills, borders from JSON spec
│   │   ├── xls_set_column_width.py        # Auto-fit or fixed width
│   │   ├── xls_freeze_panes.py            # Freeze rows/columns for scrolling
│   │   ├── xls_apply_conditional_formatting.py  # ColorScale, DataBar, IconSet
│   │   └── xls_set_number_format.py       # Currency, %, date format codes
│   ├── macros/
│   │   ├── xls_has_macros.py              # Boolean VBA presence check
│   │   ├── xls_inspect_macros.py          # List VBA modules + signature status
│   │   ├── xls_validate_macro_safety.py   # Risk scan: auto-exec, Shell, IOCs
│   │   ├── xls_remove_macros.py     ⚠️⚠️  # Strip VBA (double-token)
│   │   └── xls_inject_vba_project.py ⚠️   # Inject pre-extracted .bin
│   └── export/
│       ├── xls_export_pdf.py              # Export via LibreOffice headless
│       ├── xls_export_csv.py              # Sheet → CSV with encoding control
│       └── xls_export_json.py             # Sheet/range → structured JSON
├── scripts/
│   ├── recalc.py                          # LibreOffice headless recalc wrapper
│   ├── install_libreoffice.sh             # Setup script for CI
│   └── generate_test_files.py             # Creates test .xlsx/.xlsm fixtures
├── tests/
│   ├── __init__.py
│   ├── conftest.py                        # Shared fixtures and helpers
│   ├── fixtures/
│   │   ├── sample.xlsx                    # Basic 3-sheet workbook
│   │   ├── sample_with_macros.xlsm        # .xlsm with AutoOpen VBA macro
│   │   ├── large_dataset.xlsx             # 500k rows for perf testing
│   │   ├── complex_formulas.xlsx          # 10 sheets, 1000+ cross-sheet formulas
│   │   ├── circular_refs.xlsx             # Workbook with intentional circular refs
│   │   ├── template.xltx                  # Template with {{placeholder}} vars
│   │   └── vbaProject.bin                 # Extracted VBA binary for injection tests
│   ├── unit/
│   │   ├── test_agent.py
│   │   ├── test_locking.py
│   │   ├── test_serializers.py
│   │   ├── test_version_hash.py
│   │   ├── test_dependency.py
│   │   ├── test_token_manager.py
│   │   ├── test_audit_trail.py
│   │   ├── test_macro_handler.py
│   │   ├── test_tier1_engine.py
│   │   ├── test_error_detector.py
│   │   ├── test_chunked_io.py
│   │   ├── test_json_io.py
│   │   └── test_exit_codes.py
│   ├── integration/
│   │   ├── test_read_tools.py
│   │   ├── test_write_tools.py
│   │   ├── test_structure_tools.py
│   │   ├── test_cell_operations.py
│   │   ├── test_calculation.py
│   │   ├── test_macro_tools.py
│   │   ├── test_objects.py
│   │   ├── test_formatting.py
│   │   ├── test_export.py
│   │   ├── test_clone_modify_workflow.py
│   │   └── test_formula_dependency_workflow.py
│   ├── property/
│   │   └── test_range_serializer.py       # Hypothesis property-based tests
│   └── performance/
│       ├── bench_read_large.py
│       ├── bench_write_large.py
│       └── bench_dependency_graph.py
└── docs/
    ├── DESIGN.md                          # Architecture blueprint
    ├── API.md                             # CLI reference for all 53 tools
    ├── WORKFLOWS.md                       # Common agent workflow recipes
    ├── GOVERNANCE.md                      # Token & safety protocol docs
    └── DEVELOPMENT.md                     # Contributing guide
```

---

# Phase 0: Project Scaffolding & Infrastructure

**Duration:** 2 days
**Goal:** Establish project structure, tooling, CI/CD, and development environment so all subsequent phases have a stable foundation.
**Dependency:** None (entry point)

---

## Files to Create

### 1. `README.md`

**Purpose:** Project overview, installation guide, quick-start workflow, and navigation hub.

**Features:**
- Project description and design philosophy (Governance-First, AI-Native, Headless)
- Installation instructions (pip, from source, LibreOffice headless)
- Quick-start 3-step workflow (clone → modify → recalc)
- Link tree to all documentation in `docs/`
- CI status badge, coverage badge, PyPI version badge

**Checklist:**
- [ ] Project tagline: "53 governance-first CLI tools for AI agents to safely read, write, and manipulate Excel workbooks"
- [ ] Installation instructions targeting Python ≥3.12
- [ ] LibreOffice headless installation guide per platform (apt, brew, choco)
- [ ] Quick-start code example demonstrating `clone → read → modify → validate → export`
- [ ] Link table: DESIGN.md, API.md, WORKFLOWS.md, GOVERNANCE.md, DEVELOPMENT.md
- [ ] Badges: CI (GitHub Actions), coverage (codecov), PyPI version, Python version
- [ ] Security notice: defusedxml mandatory, macro safety scanning with oletools
- [ ] License: MIT

---

### 2. `LICENSE`

**Purpose:** MIT License file.

**Checklist:**
- [ ] MIT License text
- [ ] Copyright year: 2026
- [ ] Copyright holder name placeholder

---

### 3. `pyproject.toml`

**Purpose:** Modern Python project metadata (PEP 518/621). Single source of truth for build, deps, tools.

**Features:**
- Build system (setuptools with `src/` layout)
- Project metadata with all 53 tool entry points
- Dependency version pinning with minimum floors
- Optional dependency groups: `[dev]`, `[test]`, `[libreoffice]`
- Tool configs: black, mypy (strict), pytest, ruff

**Checklist:**
- [ ] `[build-system]` with `setuptools>=68.0` and `setuptools-scm`
- [ ] `[project]` metadata: `name="excel-agent-tools"`, `version="1.0.0"`, `requires-python=">=3.12"`
- [ ] `dependencies`: `openpyxl>=3.1.5`, `defusedxml>=0.7.1`, `oletools>=0.60`, `formulas[excel]>=1.3.0`, `pandas>=2.0.0`, `jsonschema>=4.0.0`
- [ ] `[project.optional-dependencies]` — `dev`: pytest, pytest-cov, hypothesis, black, mypy, ruff, pre-commit; `libreoffice`: (marker for Tier 2 engine docs)
- [ ] `[project.scripts]` — all 53 tool entry points (`xls-clone-workbook = "excel_agent.tools.governance.xls_clone_workbook:main"`, etc.)
- [ ] `[tool.black]` — `line-length = 99`, `target-version = ["py312"]`
- [ ] `[tool.mypy]` — `strict = true`, `warn_return_any = true`, `disallow_any_generics = true`, `no_implicit_reexport = true`
- [ ] `[tool.pytest.ini_options]` — `testpaths = ["tests"]`, `markers` for `slow`, `libreoffice`, `integration`
- [ ] `[tool.ruff]` — `select = ["E", "F", "I", "N", "W", "UP", "S", "B"]`, `line-length = 99`

---

### 4. `requirements.txt`

**Purpose:** Pinned runtime dependencies with hashes for supply-chain security.

**Checklist:**
- [ ] `openpyxl==3.1.5` — read/write Excel 2010 xlsx/xlsm/xltx/xltm files
- [ ] `defusedxml==0.7.1` — contains several Python-only workarounds and fixes for denial of service and other vulnerabilities in Python's XML libraries
- [ ] `oletools==0.60.2` — can detect, extract and analyse VBA macros, OLE objects, Excel 4 macros (XLM) and DDE links
- [ ] `formulas[excel]==1.3.4` — implements an interpreter for Excel formulas; compiles Excel workbooks to python and executes without using the Excel COM server. Hence, Excel is not needed.
- [ ] `pandas>=2.1.0` (internal chunked I/O only)
- [ ] `jsonschema>=4.20.0` (input validation)
- [ ] All entries include `--hash` pins for reproducibility

---

### 5. `requirements-dev.txt`

**Purpose:** Development, testing, and linting dependencies.

**Checklist:**
- [ ] `pytest>=8.0.0`
- [ ] `pytest-cov>=5.0.0`
- [ ] `hypothesis>=6.100.0`
- [ ] `black>=24.0.0`
- [ ] `mypy>=1.10.0`
- [ ] `ruff>=0.5.0`
- [ ] `pre-commit>=3.7.0`
- [ ] `types-defusedxml>=0.7.0` — a PEP 561 type stub package for the defusedxml package. It can be used by type-checking tools like mypy, pyright, pytype, PyCharm, etc. to check code that uses defusedxml.
- [ ] `-r requirements.txt` (include runtime deps)

---

### 6. `.gitignore`

**Purpose:** Exclude generated files, temp files, and sensitive data from version control.

**Checklist:**
- [ ] Python caches: `__pycache__/`, `*.pyc`, `*.pyo`, `.pytest_cache/`, `.mypy_cache/`, `.ruff_cache/`
- [ ] Virtual envs: `venv/`, `.venv/`, `env/`
- [ ] IDEs: `.vscode/`, `.idea/`, `*.swp`, `*.swo`
- [ ] Build artifacts: `build/`, `dist/`, `*.egg-info/`, `.eggs/`
- [ ] Test outputs: `.coverage`, `htmlcov/`, `.tox/`, `coverage.xml`
- [ ] Working files: `/work/`, `*.tmp.xlsx`, `.~lock.*#` (LibreOffice lock files)
- [ ] Audit trail: `.excel_agent_audit.jsonl` (unless explicitly committed)
- [ ] Secrets: `.env`, `*.pem`, `*.key`

---

### 7. `.github/workflows/ci.yml`

**Purpose:** GitHub Actions CI/CD pipeline with matrix testing and quality gates.

**Features:**
- Multi-Python matrix (3.12, 3.13)
- LibreOffice headless installation for Tier 2 tests
- Linting gate (black, mypy, ruff)
- Test execution with coverage reporting
- Coverage enforcement (≥90% for merge)

**Checklist:**
- [ ] Trigger: `push` to `main`, all `pull_request`
- [ ] Matrix: `python-version: ["3.12", "3.13"]`, `os: [ubuntu-latest]`
- [ ] Install LibreOffice: `sudo apt-get install -y libreoffice-calc`
- [ ] Install deps: `pip install -r requirements.txt -r requirements-dev.txt && pip install -e .`
- [ ] Lint step: `black --check src/ tools/ tests/`, `ruff check src/ tools/`, `mypy src/`
- [ ] Test step: `pytest --cov=excel_agent --cov-report=xml --cov-fail-under=90 -m "not slow"`
- [ ] Integration test step (separate job): `pytest -m integration`
- [ ] Upload coverage to codecov
- [ ] Fail build on any lint error, test failure, or coverage < 90%

---

### 8. `setup.py` (Legacy Compatibility Shim)

**Purpose:** Minimal fallback for older pip versions that don't support `pyproject.toml`.

**Checklist:**
- [ ] Single line: `from setuptools import setup; setup()`
- [ ] Comment explaining it delegates to `pyproject.toml`

---

### 9. `src/excel_agent/__init__.py`

**Purpose:** Package initialization, version export, and public API convenience imports.

**Checklist:**
- [ ] `__version__ = "1.0.0"`
- [ ] Convenience imports: `ExcelAgent`, `DependencyTracker`, `ApprovalTokenManager`, `AuditTrail`
- [ ] `__all__` list restricting public API surface
- [ ] Module docstring describing the package purpose

---

### 10. `src/excel_agent/utils/__init__.py`

**Purpose:** Utils package init.

**Checklist:**
- [ ] Empty or minimal `__all__` list

---

### 11. `src/excel_agent/utils/exit_codes.py`

**Purpose:** Standardized exit code constants for all 53 tools.

**Interface:**
```python
from enum import IntEnum

class ExitCode(IntEnum):
    """Standardized exit codes for all excel-agent-tools."""
    SUCCESS = 0             # Operation completed successfully
    VALIDATION_ERROR = 1    # Input validation failed (malformed JSON, bad range, schema error)
    FILE_NOT_FOUND = 2      # Input file does not exist or is not readable
    LOCK_CONTENTION = 3     # File is locked by another process, timeout exceeded
    PERMISSION_DENIED = 4   # Approval token invalid, expired, revoked, or wrong scope
    INTERNAL_ERROR = 5      # Unexpected error (bug, corrupt file, LibreOffice crash)
```

**Checklist:**
- [ ] Define `ExitCode(IntEnum)` with docstrings for each member
- [ ] Helper function `exit_with(code: ExitCode, message: str) -> NoReturn` that prints JSON error and `sys.exit()`
- [ ] Export in `__init__.py`
- [ ] Unit test: all 6 codes have distinct integer values

---

### 12. `src/excel_agent/utils/json_io.py`

**Purpose:** Standardized JSON output formatting with consistent response schema.

**Interface:**
```python
from typing import Any
from datetime import datetime
from pathlib import Path

def build_response(
    status: str,                           # "success" | "error" | "warning" | "denied"
    data: Any,
    *,
    workbook_version: str = "",
    impact: dict[str, Any] | None = None,  # {"cells_modified": N, "formulas_updated": N}
    warnings: list[str] | None = None,
    exit_code: int = 0,
    guidance: str | None = None,           # Prescriptive guidance for agent on denial
) -> dict[str, Any]:
    """Builds standardized JSON response envelope."""
    ...

def print_json(data: dict[str, Any], *, indent: int = 2) -> None:
    """Prints JSON to stdout. Never writes to stderr."""
    ...

class ExcelAgentEncoder(json.JSONEncoder):
    """Custom encoder for datetime, Path, bytes, Decimal."""
    ...
```

**Checklist:**
- [ ] `build_response()` returns dict matching the universal response schema
- [ ] `ExcelAgentEncoder` handles: `datetime` → ISO 8601, `Path` → string, `bytes` → hex, `Decimal` → float
- [ ] `print_json()` outputs to `sys.stdout` only (no stderr pollution that would confuse agent)
- [ ] Response always includes `"timestamp"` (ISO 8601 UTC) for audit correlation
- [ ] Unit test: `datetime` serialization roundtrip
- [ ] Unit test: `None` data produces `"data": null`
- [ ] Unit test: nested dicts serialize correctly

---

### 13. `src/excel_agent/utils/cli_helpers.py`

**Purpose:** Reusable CLI argument parsing, path validation, and JSON input handling.

**Interface:**
```python
import argparse
from pathlib import Path

def add_common_args(parser: argparse.ArgumentParser) -> None:
    """Adds --input, --output, --sheet, --format flags."""
    ...

def add_governance_args(parser: argparse.ArgumentParser) -> None:
    """Adds --token, --force, --acknowledge-impact flags."""
    ...

def validate_input_path(path: str) -> Path:
    """Validates file exists and is readable. Raises SystemExit(2) if not."""
    ...

def validate_output_path(path: str) -> Path:
    """Validates parent directory exists and is writable."""
    ...

def load_json_stdin() -> dict[str, Any]:
    """Reads JSON from stdin, validates it's a dict. Raises SystemExit(1) on malformed."""
    ...

def parse_range_arg(range_str: str) -> str:
    """Basic pre-validation of range argument format."""
    ...
```

**Checklist:**
- [ ] `add_common_args()`: `--input` (required), `--output` (optional), `--sheet` (optional), `--format` (json/jsonl)
- [ ] `add_governance_args()`: `--token` (string), `--force` (flag), `--acknowledge-impact` (flag)
- [ ] `validate_input_path()`: checks `Path.exists()`, `Path.is_file()`, readable permission; returns `Path` or exits with code 2
- [ ] `validate_output_path()`: checks parent dir exists and is writable; creates parent if `--force`
- [ ] `load_json_stdin()`: reads from `sys.stdin`, `json.loads()`, validates top-level is dict; exits with code 1 on malformed
- [ ] Unit test: valid path passes
- [ ] Unit test: nonexistent path exits with code 2 and JSON error message
- [ ] Unit test: malformed JSON stdin exits with code 1

---

### 14. `src/excel_agent/utils/exceptions.py`

**Purpose:** Custom exception hierarchy for structured error handling across all tools.

**Interface:**
```python
class ExcelAgentError(Exception):
    """Base exception for all excel-agent errors."""
    exit_code: int = 5

class FileNotFoundError(ExcelAgentError):
    exit_code: int = 2

class LockContentionError(ExcelAgentError):
    exit_code: int = 3

class PermissionDeniedError(ExcelAgentError):
    exit_code: int = 4

class ConcurrentModificationError(ExcelAgentError):
    exit_code: int = 5

class ValidationError(ExcelAgentError):
    exit_code: int = 1

class ImpactDeniedError(ExcelAgentError):
    """Raised when a destructive operation would break formula references."""
    exit_code: int = 1
    impact_report: dict  # The full impact analysis
    guidance: str        # Prescriptive next step for the agent
```

**Checklist:**
- [ ] All exceptions inherit from `ExcelAgentError`
- [ ] Each has `exit_code` attribute mapping to `ExitCode` enum
- [ ] `ImpactDeniedError` includes `impact_report` and `guidance` for agent prescriptive denial
- [ ] `__str__()` returns human-readable message
- [ ] Unit test: each exception maps to correct exit code

---

### 15. `tests/__init__.py` and `tests/conftest.py`

**Purpose:** Test infrastructure: shared fixtures, temp directories, sample workbook builders.

**Conftest Features:**
```python
@pytest.fixture
def sample_workbook(tmp_path: Path) -> Path:
    """Creates a basic 3-sheet workbook with formulas."""
    ...

@pytest.fixture
def macro_workbook(tmp_path: Path) -> Path:
    """Creates an .xlsm workbook with VBA project."""
    ...

@pytest.fixture
def large_workbook(tmp_path: Path) -> Path:
    """Creates a 100k-row workbook for performance tests."""
    ...

@pytest.fixture
def complex_formula_workbook(tmp_path: Path) -> Path:
    """Creates 10-sheet workbook with 1000+ cross-sheet formulas."""
    ...

@pytest.fixture
def token_manager() -> ApprovalTokenManager:
    """Returns configured token manager with test secret."""
    ...
```

**Checklist:**
- [ ] `tests/__init__.py` is empty
- [ ] `conftest.py` provides at least 5 fixtures: `sample_workbook`, `macro_workbook`, `large_workbook`, `complex_formula_workbook`, `token_manager`
- [ ] All fixtures use `tmp_path` for isolation (no test pollution)
- [ ] `sample_workbook` has: Sheet1 (data + formulas), Sheet2 (cross-sheet refs), Sheet3 (named ranges)
- [ ] Pytest markers registered: `slow`, `libreoffice`, `integration`, `property`

---

### 16. `scripts/generate_test_files.py`

**Purpose:** Programmatic generation of all test fixture files. Reproducible, no binary blobs in repo.

**Checklist:**
- [ ] Generates `sample.xlsx`: 3 sheets, 50 cells with data, 20 formulas including cross-sheet
- [ ] Generates `complex_formulas.xlsx`: 10 sheets, 1000+ formulas, named ranges, tables
- [ ] Generates `circular_refs.xlsx`: Intentional circular references (A1=B1, B1=C1, C1=A1)
- [ ] Generates `template.xltx`: Template with `{{company}}`, `{{year}}`, `{{author}}` placeholders
- [ ] Generates `large_dataset.xlsx`: 500k rows × 10 columns (uses write-only mode for speed)
- [ ] Does NOT generate `.xlsm` (macro binaries must be pre-extracted; see test fixture docs)
- [ ] Script is idempotent: running twice produces identical files

---

**Phase 0 Exit Criteria:**
- [ ] All 16 files/directories created and pass linting (`black --check`, `ruff check`, `mypy`)
- [ ] CI pipeline runs successfully (even with no tests — empty test suite is OK for now)
- [ ] `pip install -e .` succeeds from project root
- [ ] `python -c "from excel_agent import __version__; print(__version__)"` prints `1.0.0`
- [ ] All 53 entry points registered (even if they just print "not yet implemented")
- [ ] `soffice --headless --version` works on CI runner
- [ ] README renders correctly on GitHub with badge placeholders
- [ ] `scripts/generate_test_files.py` creates all fixture files successfully

---

# Phase 1: Core Foundation

**Duration:** 5 days
**Goal:** Implement the central hub: ExcelAgent context manager, cross-platform file locking, unified range parsing, and geometry-aware version hashing.
**Dependency:** Phase 0 complete

---

## Files to Create

### 17. `src/excel_agent/core/__init__.py`

**Purpose:** Core package init with convenience imports.

**Checklist:**
- [ ] Import and re-export: `ExcelAgent`, `FileLock`, `RangeSerializer`, `CellCoordinate`, `RangeCoordinate`
- [ ] `__all__` list

---

### 18. `src/excel_agent/core/locking.py`

**Purpose:** Cross-platform atomic file locking. Prevents concurrent agent access to the same workbook.

**Interface:**
```python
from pathlib import Path

class FileLock:
    """OS-level file lock with timeout, retry, and contention detection."""

    def __init__(self, path: Path, *, timeout: float = 30.0, poll_interval: float = 0.1):
        ...

    def __enter__(self) -> 'FileLock':
        """Acquires exclusive lock. Raises LockContentionError on timeout."""
        ...

    def __exit__(self, exc_type: type | None, exc_val: BaseException | None, exc_tb: object) -> None:
        """Releases lock. Always releases even on exception."""
        ...

    @staticmethod
    def is_locked(path: Path) -> bool:
        """Non-blocking check: is this file currently locked by any process?"""
        ...
```

**Implementation Details:**
- Unix: `fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)` with polling loop
- Windows: `msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)` with polling loop
- Lock file: `.{filename}.lock` adjacent to target (avoids modifying the Excel file)
- Timeout uses exponential backoff: 0.1s → 0.2s → 0.4s → 0.8s → 1.0s (capped)

**Checklist:**
- [ ] Unix implementation using `fcntl.flock()` with `LOCK_EX | LOCK_NB`
- [ ] Windows implementation using `msvcrt.locking()` with `LK_NBLCK`
- [ ] Platform detection via `sys.platform` or `os.name`
- [ ] Exponential backoff polling with cap at 1.0s interval
- [ ] Raise `LockContentionError` (exit code 3) when timeout expires
- [ ] Lock file cleanup in `__exit__` even on exception (finally block)
- [ ] `is_locked()` attempts non-blocking acquire, immediately releases; returns bool
- [ ] Unit test: acquire and release cycle succeeds
- [ ] Unit test: second lock attempt within timeout window raises `LockContentionError`
- [ ] Unit test: lock released after exception in context body
- [ ] Unit test: `is_locked()` returns `True` while lock held, `False` after release
- [ ] Integration test: two `subprocess` processes competing for same lock; one gets exit code 3

---

### 19. `src/excel_agent/core/serializers.py`

**Purpose:** Unified range parsing that converts any Excel reference format to internal coordinates and back.

**Interface:**
```python
from dataclasses import dataclass
from openpyxl import Workbook

@dataclass(frozen=True)
class CellCoordinate:
    row: int    # 1-indexed
    col: int    # 1-indexed

@dataclass(frozen=True)
class RangeCoordinate:
    sheet: str | None       # None = active sheet
    min_row: int
    min_col: int
    max_row: int | None     # None = single cell
    max_col: int | None     # None = single cell

class RangeSerializer:
    """Parses A1, R1C1, Named Range, Table[Column] references to coordinates."""

    def __init__(self, workbook: Workbook | None = None):
        """Workbook context enables named range and table resolution."""
        ...

    def parse(self, range_str: str, *, default_sheet: str | None = None) -> RangeCoordinate:
        """Parses any supported format to RangeCoordinate."""
        ...

    def to_a1(self, coord: RangeCoordinate) -> str:
        """RangeCoordinate → 'Sheet1!A1:C10' or 'A1:C10'."""
        ...

    def to_r1c1(self, coord: RangeCoordinate) -> str:
        """RangeCoordinate → 'R1C1:R10C3'."""
        ...

    @staticmethod
    def col_letter_to_number(letter: str) -> int:
        """'A' → 1, 'Z' → 26, 'AA' → 27, 'XFD' → 16384."""
        ...

    @staticmethod
    def col_number_to_letter(number: int) -> str:
        """1 → 'A', 27 → 'AA', 16384 → 'XFD'."""
        ...
```

**Supported Formats:**
| Format | Example | Notes |
|:---|:---|:---|
| A1 single cell | `A1`, `$A$1` | Absolute markers stripped |
| A1 range | `A1:C10`, `$A$1:$C$10` | |
| Cross-sheet | `Sheet1!A1:C10`, `'Sheet Name'!A1` | Quoted names for spaces |
| R1C1 single | `R1C1` | Converted to 1-indexed |
| R1C1 range | `R1C1:R10C3` | |
| Named Range | `SalesData` | Resolved via `workbook.defined_names` |
| Table ref | `Table1`, `Table1[Sales]`, `Table1[#All]` | Resolved via sheet tables |
| Full row/col | `A:A`, `1:1` | `max_row`/`max_col` = None for "entire" |

**Checklist:**
- [ ] Parse A1 notation with optional sheet prefix and absolute markers
- [ ] Parse R1C1 notation
- [ ] Resolve named ranges via `workbook.defined_names` (requires workbook context)
- [ ] Resolve table references via iterating `sheet.tables` (requires workbook context)
- [ ] Handle quoted sheet names with single quotes (`'My Sheet'!A1`)
- [ ] `to_a1()` reverse conversion (round-trip fidelity)
- [ ] `to_r1c1()` reverse conversion
- [ ] `col_letter_to_number()` handles A-Z, AA-AZ, ..., XFD
- [ ] `col_number_to_letter()` reverse
- [ ] Raise `ValidationError` for malformed inputs with clear error message
- [ ] Edge cases: single cell (max_row=None), full row `1:1`, full column `A:A`
- [ ] Unit tests: 20+ test cases covering all formats
- [ ] Property-based test (Hypothesis): `to_a1(parse(x)) == x` roundtrip for random valid A1 strings

---

### 20. `src/excel_agent/core/version_hash.py`

**Purpose:** Geometry-aware workbook hashing for detecting concurrent modifications. Hashes *structure and formulas*, not values (values change on recalc; structure changes indicate mutation).

**Interface:**
```python
from openpyxl import Workbook

def compute_workbook_hash(workbook: Workbook) -> str:
    """SHA-256 hash of workbook geometry.
    Includes: sheet names (in order), visibility, cell coordinates with formulas.
    Excludes: cell values, styles, metadata.
    Returns: 'sha256:' + hex digest."""
    ...

def compute_sheet_hash(sheet: 'Worksheet') -> str:
    """SHA-256 hash of single sheet geometry."""
    ...

def compute_file_hash(path: Path) -> str:
    """SHA-256 hash of file bytes on disk. For concurrent modification detection."""
    ...
```

**Hashing Algorithm:**
1. Sort sheets by index
2. For each sheet: hash `(name, visibility, min_row, max_row, min_col, max_col)`
3. For each cell with a formula: hash `(sheet_name, row, col, formula_string)`
4. Combine all hashes in deterministic order → SHA-256 hex digest
5. Prefix with `sha256:` for self-describing format

**Checklist:**
- [ ] Iterate sheets in `workbook.sheetnames` order (deterministic)
- [ ] For each sheet: include `title`, `sheet_state` (visible/hidden/veryHidden)
- [ ] For cells with formulas: include `(row, col, formula_string)` — NOT values
- [ ] Use `hashlib.sha256()` with incremental `.update()` for memory efficiency
- [ ] Return `f"sha256:{digest.hexdigest()}"` format
- [ ] `compute_file_hash()` reads file in 64KB chunks for large file support
- [ ] Unit test: identical workbooks produce identical hashes
- [ ] Unit test: changing a cell *value* produces the SAME hash (values excluded)
- [ ] Unit test: changing a cell *formula* produces a DIFFERENT hash
- [ ] Unit test: renaming a sheet produces a different hash
- [ ] Unit test: adding/removing a sheet produces a different hash
- [ ] Unit test: reordering sheets produces a different hash

---

### 21. `src/excel_agent/core/agent.py`

**Purpose:** The central hub — stateful context manager integrating locking, loading, saving, and hash verification for safe workbook manipulation.

**Interface:**
```python
from pathlib import Path
from openpyxl import Workbook

class ExcelAgent:
    """Stateful context manager for safe, locked, hash-verified workbook manipulation."""

    def __init__(
        self,
        path: Path,
        *,
        mode: str = "rw",              # "r" = read-only, "rw" = read-write
        keep_vba: bool = True,         # Preserve VBA projects in .xlsm
        lock_timeout: float = 30.0,    # Seconds before LockContentionError
        data_only: bool = False,       # False = preserve formulas; True = read cached values
    ):
        ...

    def __enter__(self) -> 'ExcelAgent':
        """1. Acquire FileLock  2. Load workbook  3. Compute entry hash."""
        ...

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """1. If no exception and mode='rw': verify hash, save  2. Release lock."""
        ...

    @property
    def workbook(self) -> Workbook:
        """Returns the openpyxl Workbook object. Raises if not entered."""
        ...

    @property
    def path(self) -> Path:
        """Returns the workbook file path."""
        ...

    @property
    def version_hash(self) -> str:
        """Returns the entry-time geometry hash."""
        ...

    def verify_no_concurrent_modification(self) -> None:
        """Re-reads file bytes, compares hash. Raises ConcurrentModificationError if changed."""
        ...
```

**Lifecycle:**
```
__enter__:
  1. self._lock = FileLock(self._path, timeout=lock_timeout)
  2. self._lock.__enter__()
  3. self._wb = load_workbook(str(path), keep_vba=keep_vba, data_only=data_only)
  4. self._entry_hash = compute_file_hash(path)  # byte-level hash for modification detection
  5. self._geometry_hash = compute_workbook_hash(self._wb)  # geometry hash for version reporting
  6. return self

__exit__ (no exception, mode='rw'):
  1. current_file_hash = compute_file_hash(self._path)
  2. if current_file_hash != self._entry_hash: raise ConcurrentModificationError
  3. self._wb.save(str(self._path))
  4. self._lock.__exit__(...)

__exit__ (with exception):
  1. self._lock.__exit__(...) — release lock without saving
  2. Do NOT save (prevent partial writes from corrupting file)
```

**Checklist:**
- [ ] `__init__`: validate `mode` is `"r"` or `"rw"`; validate path exists for `"r"` and `"rw"` modes
- [ ] `__enter__`: acquire lock FIRST, then load workbook, then hash
- [ ] `__enter__`: call `load_workbook(keep_vba=True, data_only=False)` — ALWAYS preserve formulas unless explicit
- [ ] `__exit__` (success path): verify file hash unchanged, save workbook, release lock
- [ ] `__exit__` (error path): release lock WITHOUT saving; re-raise exception
- [ ] `verify_no_concurrent_modification()`: re-reads file bytes, compares with entry hash
- [ ] Raise `ConcurrentModificationError` (exit code 5) if file modified externally during session
- [ ] `.xlsm` detection: auto-set `keep_vba=True` if file extension is `.xlsm`
- [ ] Property guards: `workbook` raises `RuntimeError` if accessed outside context
- [ ] Unit test: successful load → modify → save cycle, verify file changed on disk
- [ ] Unit test: read-only mode (`mode="r"`) does NOT save even if workbook modified in memory
- [ ] Unit test: exception in context body → lock released, file unchanged
- [ ] Unit test: simulated external modification (write to file between enter and exit) → raises `ConcurrentModificationError`
- [ ] Unit test: `.xlsm` file preserves VBA project after save (verify `xl/vbaProject.bin` exists in ZIP)
- [ ] Unit test: lock timeout when another process holds lock → exit code 3

---

### 22. `tests/unit/test_locking.py`

**Checklist:**
- [ ] Test: acquire lock, verify lock file exists, release, verify lock file cleaned up
- [ ] Test: `is_locked()` returns `True` while held, `False` after release
- [ ] Test: second lock on same file within timeout raises `LockContentionError`
- [ ] Test: lock released even when context body raises exception
- [ ] Test: concurrent access via `multiprocessing.Process` — one succeeds, one gets exit code 3

---

### 23. `tests/unit/test_serializers.py`

**Checklist:**
- [ ] Test: `"A1"` → `RangeCoordinate(None, 1, 1, None, None)`
- [ ] Test: `"A1:C10"` → `RangeCoordinate(None, 1, 1, 10, 3)`
- [ ] Test: `"Sheet1!A1:C10"` → `RangeCoordinate("Sheet1", 1, 1, 10, 3)`
- [ ] Test: `"'Sheet Name'!A1"` → `RangeCoordinate("Sheet Name", 1, 1, None, None)`
- [ ] Test: `"R1C1:R10C3"` → same as A1:C10
- [ ] Test: `"$A$1:$C$10"` → absolute markers stripped, same result as `"A1:C10"`
- [ ] Test: Named range resolution (requires workbook fixture)
- [ ] Test: Table reference resolution (requires workbook fixture)
- [ ] Test: `col_letter_to_number("A")=1`, `"Z"=26`, `"AA"=27`, `"XFD"=16384`
- [ ] Test: Roundtrip `to_a1(parse(x)) == x` for 20+ inputs
- [ ] Test: Malformed input raises `ValidationError`

---

### 24. `tests/unit/test_version_hash.py`

**Checklist:**
- [ ] Test: identical workbooks produce identical hashes
- [ ] Test: value-only change → same geometry hash
- [ ] Test: formula change → different geometry hash
- [ ] Test: sheet rename → different geometry hash
- [ ] Test: sheet add/remove → different geometry hash
- [ ] Test: sheet reorder → different geometry hash
- [ ] Test: `compute_file_hash()` changes when file is modified

---

### 25. `tests/unit/test_agent.py`

**Checklist:**
- [ ] Test: basic enter → modify → exit → file saved with new data
- [ ] Test: read-only mode → file NOT modified
- [ ] Test: exception in body → file NOT modified, lock released
- [ ] Test: concurrent modification detection raises `ConcurrentModificationError`
- [ ] Test: `.xlsm` preserves VBA project
- [ ] Test: `version_hash` property returns `sha256:...` format string

---

### 26. `tests/property/test_range_serializer.py`

**Purpose:** Hypothesis property-based tests for RangeSerializer roundtrip fidelity.

**Checklist:**
- [ ] Strategy: generate random valid A1 strings (`[A-Z]{1,3}[0-9]{1,7}`)
- [ ] Strategy: generate random valid ranges (`cell:cell`)
- [ ] Property: `to_a1(parse(generated_a1)) == generated_a1`
- [ ] Property: `col_number_to_letter(col_letter_to_number(x)) == x` for all valid column letters
- [ ] Run at least 200 examples per property

---

**Phase 1 Exit Criteria:**
- [ ] All 10 files (including tests) pass with >95% coverage
- [ ] FileLock works on Linux (CI) — Windows tested manually or skipped with marker
- [ ] RangeSerializer handles all 8 input formats documented above
- [ ] ExcelAgent context manager correctly loads, modifies, saves, and detects concurrent modification
- [ ] Hypothesis property tests pass for RangeSerializer (200+ examples)
- [ ] No mypy errors in strict mode
- [ ] `ExcelAgent` with `.xlsm` file preserves VBA project

---

# Phase 2: Dependency Engine & Schema Validation

**Duration:** 5 days
**Goal:** Implement the formula dependency graph (powered by `formulas` library) and JSON schema validation infrastructure.
**Dependency:** Phase 1 complete (needs `ExcelAgent`, `RangeSerializer`)

---

## Files to Create

### 27. `src/excel_agent/core/dependency.py`

**Purpose:** Build and query the workbook's formula dependency graph. This is the **most safety-critical component** — it powers pre-flight impact reports that prevent agents from breaking formula chains.

**Design Notes:** `formulas` implements an interpreter for Excel formulas, which parses and compile Excel formulas expressions. Moreover, it compiles Excel workbooks to python and executes without using the Excel COM server. Hence, Excel is not needed. The `formulas` library can handle circular references with `circular=True` and plot the dependency graph that depicts relationships between Excel cells.

**Interface:**
```python
from dataclasses import dataclass, field
from openpyxl import Workbook

@dataclass
class ImpactReport:
    """Pre-flight impact analysis for destructive operations."""
    status: str                         # "safe" | "warning" | "critical"
    broken_references: int              # Number of formulas that would produce #REF!
    affected_sheets: list[str]          # Sheets containing affected formulas
    sample_errors: list[str]            # First 10 affected cells: "Sheet1!B4 → #REF!"
    circular_refs_affected: bool        # Whether action affects a circular ref chain
    suggestion: str                     # Prescriptive guidance for the agent
    details: dict[str, list[str]]       # Per-sheet list of affected cells

class DependencyTracker:
    """Builds and queries the workbook's formula dependency graph."""

    def __init__(self, workbook: Workbook):
        ...

    def build_graph(self, *, sheets: list[str] | None = None) -> None:
        """Parses all formulas, builds directed dependency graph.
        Uses openpyxl Tokenizer for cell reference extraction."""
        ...

    @property
    def is_built(self) -> bool:
        """Whether the graph has been built."""
        ...

    def find_dependents(self, target: str) -> set[str]:
        """Returns all cells that would be affected if target is deleted/changed.
        Performs transitive closure: A→B→C means deleting A affects B and C.
        Target format: 'Sheet1!A1' or 'A1' (active sheet)."""
        ...

    def find_precedents(self, cell: str) -> set[str]:
        """Returns all cells that the given cell depends on (its inputs)."""
        ...

    def impact_report(self, target_range: str, *, action: str = "delete") -> ImpactReport:
        """Pre-flight check for destructive operations.
        Returns structured impact analysis with prescriptive guidance."""
        ...

    def detect_circular_references(self) -> list[list[str]]:
        """Returns list of circular dependency cycles (Tarjan's SCC algorithm)."""
        ...

    def get_adjacency_list(self) -> dict[str, list[str]]:
        """Exports full graph as JSON-serializable adjacency list."""
        ...

    def get_stats(self) -> dict[str, int]:
        """Returns: total_cells, total_formulas, total_edges, circular_chains."""
        ...
```

**Graph Construction Algorithm:**
1. Iterate all sheets (or specified subset)
2. For each cell: if `cell.data_type == 'f'` (formula cell):
   a. Tokenize formula string using openpyxl's `Tokenizer`
   b. Extract all `OPERAND` tokens of subtype `RANGE` → these are cell references
   c. Normalize references to absolute format: `Sheet1!A1`
   d. Add edges: `referenced_cell → formula_cell` (forward graph)
   e. Add reverse edges: `formula_cell → referenced_cell` (precedent graph)
3. For `find_dependents()`: BFS/DFS on forward graph from target → transitive closure
4. For `detect_circular_references()`: Tarjan's strongly connected components algorithm

**Checklist:**
- [ ] Use openpyxl's `Tokenizer` class for formula parsing (shared foundation with `formulas` library)
- [ ] Build forward graph: `{cell: set(cells_that_reference_it)}` — "who depends on me?"
- [ ] Build reverse graph: `{cell: set(cells_it_references)}` — "who do I depend on?"
- [ ] `find_dependents()`: BFS traversal on forward graph, returns transitive closure
- [ ] `find_precedents()`: direct lookup on reverse graph
- [ ] `impact_report()`: compute dependents for all cells in target range, aggregate into `ImpactReport`
- [ ] `impact_report()` `suggestion` field: if broken_references > 0, suggest `"Run xls_update_references.py --target='...' before retrying"`
- [ ] `detect_circular_references()`: Tarjan's SCC algorithm; returns list of cycles
- [ ] Handle cross-sheet references: `'Sheet1'!A1` normalized form
- [ ] Handle named ranges in formulas: resolve via workbook.defined_names, expand to cell refs
- [ ] Handle range references: `A1:C10` expands to all individual cells in range (for precision)
- [ ] Lazy construction: `build_graph()` must be explicitly called; not on `__init__`
- [ ] Performance: build graph for 10-sheet, 1000-formula workbook in <5 seconds
- [ ] Unit test: empty workbook → empty graph, `get_stats()` returns all zeros
- [ ] Unit test: `A1=5` (value, not formula) → not in graph
- [ ] Unit test: `A1=B1` → graph shows A1 depends on B1
- [ ] Unit test: chain `A1=B1, B1=C1, C1=5` → `find_dependents("C1")` returns `{"B1", "A1"}`
- [ ] Unit test: cross-sheet `Sheet1!A1=Sheet2!B1` → dependency tracked
- [ ] Unit test: circular `A1=B1, B1=A1` → `detect_circular_references()` returns `[["A1", "B1"]]`
- [ ] Unit test: `impact_report("Sheet1!C1", action="delete")` returns `broken_references=2` for chain
- [ ] Unit test: large workbook fixture (1000 formulas) builds in <5s (mark as `@pytest.mark.slow`)

---

### 28. `src/excel_agent/governance/__init__.py`

**Checklist:**
- [ ] Import and re-export: `ApprovalTokenManager`, `AuditTrail`

---

### 29. `src/excel_agent/governance/schemas/__init__.py`

**Purpose:** Schema loader utility with caching.

**Interface:**
```python
from pathlib import Path

_SCHEMA_CACHE: dict[str, dict] = {}

def load_schema(schema_name: str) -> dict:
    """Loads JSON schema by name from schemas/ directory. Caches in memory."""
    ...

def validate_against_schema(schema_name: str, data: dict) -> None:
    """Validates data against named schema. Raises jsonschema.ValidationError."""
    ...
```

**Checklist:**
- [ ] `load_schema()` reads `.schema.json` files from the schemas directory
- [ ] In-memory caching: load once, reuse across calls
- [ ] `validate_against_schema()` uses `jsonschema.validate(data, schema)`
- [ ] Raise `ValidationError` (exit code 1) with clear path to invalid field
- [ ] Unit test: valid data passes
- [ ] Unit test: invalid data raises with descriptive message

---

### 30. `src/excel_agent/governance/schemas/range_input.schema.json`

**Purpose:** Validates range input arguments across all tools.

**Checklist:**
- [ ] Supports A1 string format (pattern: `^[A-Za-z]+[0-9]+(:[A-Za-z]+[0-9]+)?$`)
- [ ] Supports cross-sheet format (`SheetName!A1:C10`)
- [ ] Supports coordinate object: `{"start_row": int, "start_col": int, "end_row": int, "end_col": int}`
- [ ] Optional `sheet` property
- [ ] Unit test: valid A1 passes; invalid `"ZZZZZ"` fails

---

### 31. `src/excel_agent/governance/schemas/write_data.schema.json`

**Purpose:** Validates cell data arrays for write operations.

**Checklist:**
- [ ] 2D array: `{"data": [[cell, cell, ...], ...]}` 
- [ ] Cell types: string, number, boolean, null
- [ ] Unit test: valid 3×3 array passes; non-array fails

---

### 32. `src/excel_agent/governance/schemas/style_spec.schema.json`

**Purpose:** Validates style/formatting specification JSON.

**Checklist:**
- [ ] Font properties: name, size, bold, italic, color (hex)
- [ ] Fill properties: fgColor, bgColor, patternType
- [ ] Border properties: top, bottom, left, right (each with style, color)
- [ ] Alignment: horizontal, vertical, wrapText, textRotation
- [ ] Number format: string pattern

---

### 33. `src/excel_agent/governance/schemas/token_request.schema.json`

**Purpose:** Validates token generation request structure.

**Checklist:**
- [ ] Required: `scope` (enum of valid scopes), `target_file` (string)
- [ ] Optional: `ttl_seconds` (integer, default 300, max 3600)
- [ ] Valid scopes: `sheet:delete`, `sheet:rename`, `range:delete`, `formula:convert`, `macro:remove`, `macro:inject`, `structure:modify`

---

### 34. `tests/unit/test_dependency.py`

**Comprehensive test suite for `DependencyTracker`.**

**Checklist:**
- [ ] Empty workbook → empty graph
- [ ] Single value cell (no formula) → not in graph
- [ ] Single dependency `A1=B1` → correct edges
- [ ] Chain `A1=B1, B1=C1` → transitive `find_dependents("C1")` = `{"B1", "A1"}`
- [ ] Cross-sheet dependency → tracked correctly
- [ ] Circular reference → detected by `detect_circular_references()`
- [ ] Impact report counts broken references correctly
- [ ] Impact report includes correct `suggestion` text
- [ ] Named range in formula → resolved and tracked
- [ ] Multi-cell range in formula (`=SUM(A1:A10)`) → all 10 cells tracked as precedents
- [ ] Performance: 1000 formulas across 10 sheets builds in <5s

---

**Phase 2 Exit Criteria:**
- [ ] `DependencyTracker.build_graph()` correctly identifies all dependencies in `complex_formulas.xlsx`
- [ ] Circular reference detection works for 2-cell and 3-cell cycles
- [ ] `impact_report()` returns accurate `broken_references` count with prescriptive `suggestion`
- [ ] Graph export (`get_adjacency_list()`) is JSON-serializable
- [ ] All JSON schemas load and validate correctly
- [ ] Performance: 10-sheet, 1000-formula workbook analyzed in <5 seconds
- [ ] All unit tests pass with >90% coverage on `dependency.py`

---

# Phase 3: Governance & Safety Layer

**Duration:** 3 days
**Goal:** Implement HMAC-SHA256 approval tokens (with TTL, nonce, file-hash binding), pluggable audit trail, and the safety enforcement protocol.
**Dependency:** Phase 2 complete (needs schemas)

---

## Files to Create

### 35. `src/excel_agent/governance/token_manager.py`

**Purpose:** HMAC-SHA256 scoped approval token system with replay protection.

**Security Foundation:** When comparing the output of hexdigest() to an externally supplied digest during a verification routine, it is recommended to use the compare_digest() function instead of the == operator to reduce the vulnerability to timing attacks. And per Python 3.10+, the function uses OpenSSL's CRYPTO_memcmp() internally when available.

**Interface:**
```python
from dataclasses import dataclass

@dataclass(frozen=True)
class ApprovalToken:
    """Immutable token structure."""
    scope: str                  # e.g., "sheet:delete"
    target_file_hash: str       # SHA-256 of target workbook (prevents cross-file reuse)
    nonce: str                  # UUID4, one-time use
    issued_at: float            # Unix timestamp (UTC)
    ttl_seconds: int            # Time-to-live (default: 300 = 5 minutes)
    signature: str              # HMAC-SHA256(secret, scope|hash|nonce|issued_at|ttl)

class ApprovalTokenManager:
    """Generates and validates scoped HMAC-SHA256 approval tokens."""

    VALID_SCOPES = frozenset({
        "sheet:delete", "sheet:rename", "range:delete",
        "formula:convert", "macro:remove", "macro:inject", "structure:modify",
    })

    def __init__(self, *, secret_key: str | None = None):
        """Secret from EXCEL_AGENT_SECRET env var, or generated per session."""
        ...

    def generate_token(
        self,
        scope: str,
        target_file_hash: str,
        *,
        ttl_seconds: int = 300,
    ) -> str:
        """Generates scoped approval token. Returns serialized token string."""
        ...

    def validate_token(
        self,
        token_str: str,
        *,
        expected_scope: str,
        expected_file_hash: str,
    ) -> ApprovalToken:
        """Validates token. Returns parsed token if valid.
        Raises PermissionDeniedError if invalid/expired/wrong-scope/reused."""
        ...

    def revoke_token(self, nonce: str) -> None:
        """Adds nonce to revocation set (prevents reuse)."""
        ...
```

**Token Format (serialized):** `base64(json({"scope", "target_file_hash", "nonce", "issued_at", "ttl_seconds", "signature"}))`

**Validation Steps (in order):**
1. Deserialize and parse JSON
2. Verify `scope == expected_scope`
3. Verify `target_file_hash == expected_file_hash`
4. Verify `issued_at + ttl_seconds > current_time` (not expired)
5. Verify `nonce` not in revocation set (not reused)
6. Recompute HMAC signature over `scope|target_file_hash|nonce|issued_at|ttl_seconds`
7. Compare using `hmac.compare_digest()` — used to safely compare two digests to prevent a type of side-channel attack called a timing attack. It returns True if a and b are equal, and False otherwise, but it does so in a constant-time manner.
8. Add nonce to used set (single-use enforcement)

**Checklist:**
- [ ] Secret key from `os.environ.get("EXCEL_AGENT_SECRET")` or `secrets.token_hex(32)` if not set
- [ ] Log warning if secret key is auto-generated (not reproducible across sessions)
- [ ] `generate_token()`: UUID4 nonce, current UTC timestamp, HMAC-SHA256 signature
- [ ] `validate_token()`: all 8 validation steps in order
- [ ] Use `hmac.compare_digest()` for signature comparison (NEVER `==`)
- [ ] Raise `PermissionDeniedError` with descriptive reason for each failure mode
- [ ] In-memory nonce tracking set (prevents replay within session)
- [ ] `revoke_token()` adds nonce to revocation set
- [ ] Scope validation: reject unknown scopes
- [ ] TTL range: 1–3600 seconds (reject out-of-range)
- [ ] Unit test: valid token passes validation
- [ ] Unit test: expired token (TTL=1, sleep 2s) raises `PermissionDeniedError`
- [ ] Unit test: wrong scope raises `PermissionDeniedError`
- [ ] Unit test: wrong file hash raises `PermissionDeniedError`
- [ ] Unit test: tampered signature raises `PermissionDeniedError`
- [ ] Unit test: reused nonce raises `PermissionDeniedError`
- [ ] Unit test: revoked token raises `PermissionDeniedError`
- [ ] Unit test: secret from env var is used when set
- [ ] Unit test: auto-generated secret logs warning

---

### 36. `src/excel_agent/governance/audit_trail.py`

**Purpose:** Pluggable audit logging for all destructive operations. Default: JSONL file.

**Interface:**
```python
from typing import Protocol
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

@dataclass
class AuditEvent:
    timestamp: str          # ISO 8601 UTC
    tool: str               # e.g., "xls_delete_sheet"
    scope: str              # Token scope used
    resource: str           # What was affected (e.g., "Sheet1", "A1:C10")
    action: str             # "delete", "rename", "convert", etc.
    outcome: str            # "success", "denied", "error"
    token_used: bool        # Whether governance token was required
    file_hash: str          # Workbook hash at time of operation
    pid: int                # Process ID for tracing
    details: dict           # Additional context

class AuditBackend(Protocol):
    """Protocol for pluggable audit backends."""
    def log_event(self, event: AuditEvent) -> None: ...
    def query_events(
        self,
        *,
        tool: str | None = None,
        start_time: datetime | None = None,
        end_time: datetime | None = None,
        limit: int = 100,
    ) -> list[AuditEvent]: ...

class JsonlAuditBackend:
    """Default: append-only JSONL file."""
    def __init__(self, log_path: Path = Path(".excel_agent_audit.jsonl")):
        ...

class NullAuditBackend:
    """No-op backend for testing or when auditing is disabled."""
    ...

class CompositeAuditBackend:
    """Fan-out to multiple backends simultaneously."""
    def __init__(self, *backends: AuditBackend):
        ...

class AuditTrail:
    """Singleton-ish audit trail manager."""
    def __init__(self, backend: AuditBackend | None = None):
        """Defaults to JsonlAuditBackend if not specified."""
        ...

    def log_operation(
        self,
        tool: str,
        scope: str,
        resource: str,
        action: str,
        outcome: str,
        *,
        token_used: bool = False,
        file_hash: str = "",
        details: dict | None = None,
    ) -> None:
        """Creates AuditEvent and delegates to backend."""
        ...
```

**JSONL Entry Format:**
```json
{"timestamp":"2026-04-08T14:30:22Z","tool":"xls_delete_sheet","scope":"sheet:delete","resource":"Sheet1","action":"delete","outcome":"success","token_used":true,"file_hash":"sha256:abc...","pid":12345,"details":{}}
```

**Checklist:**
- [ ] `JsonlAuditBackend.log_event()`: atomic append to `.jsonl` file (open, write, flush, close)
- [ ] Create log file if it doesn't exist
- [ ] ISO 8601 UTC timestamp via `datetime.now(timezone.utc).isoformat()`
- [ ] Include `os.getpid()` for multi-process tracing
- [ ] `query_events()`: read file, filter by tool/time range, return list
- [ ] `NullAuditBackend`: no-op for tests
- [ ] `CompositeAuditBackend`: iterates backends, calls `log_event()` on each
- [ ] File-level locking during write (prevent corruption from concurrent processes)
- [ ] Unit test: single event logged correctly
- [ ] Unit test: multiple events maintain valid JSONL (one JSON object per line)
- [ ] Unit test: query by tool name filters correctly
- [ ] Unit test: query by time range filters correctly
- [ ] Unit test: `NullAuditBackend` does not write anything
- [ ] Unit test: `CompositeAuditBackend` forwards to all backends
- [ ] Integration test: 10 concurrent processes logging → no corruption

---

### 37. `tests/unit/test_token_manager.py`

**Checklist:** (See token_manager.py checklist above for all 9 unit test cases)

---

### 38. `tests/unit/test_audit_trail.py`

**Checklist:** (See audit_trail.py checklist above for all 7 unit test cases)

---

**Phase 3 Exit Criteria:**
- [ ] Token generation and validation works correctly for all 7 scopes
- [ ] Expired tokens rejected with exit code 4
- [ ] Replay protection: reused nonce rejected
- [ ] File-hash binding: token generated for `file_A` cannot be used for `file_B`
- [ ] Audit trail logs all operations to `.jsonl` file without corruption
- [ ] Pluggable backend architecture works (Null, JSONL, Composite)
- [ ] 100% coverage on `token_manager.py` and `audit_trail.py`

---

# Phase 4: Governance & Read Tools

**Duration:** 5 days
**Goal:** Implement all 6 governance CLI tools and all 7 read-only introspection tools. Plus the chunked I/O helper.
**Dependency:** Phase 3 complete

---

## Files to Create

### 39. `src/excel_agent/core/chunked_io.py`

**Purpose:** Streaming read/write for large datasets (>100k rows) without loading entire workbook into memory.

**Interface:**
```python
from typing import Generator, Any
from openpyxl.worksheet.worksheet import Worksheet

def read_range_chunked(
    sheet: Worksheet,
    min_row: int, min_col: int, max_row: int, max_col: int,
    *,
    chunk_size: int = 10_000,
) -> Generator[list[list[Any]], None, None]:
    """Yields chunks of rows from specified range."""
    ...

def count_used_rows(sheet: Worksheet) -> int:
    """Returns count of rows with data (not just sheet.max_row which may be inflated)."""
    ...
```

**Checklist:**
- [ ] `read_range_chunked()`: iterate rows in chunks of `chunk_size`, yield 2D list per chunk
- [ ] Handle cell types: string, number, boolean, datetime, None
- [ ] Convert datetime cells to ISO 8601 strings
- [ ] Memory: never hold more than `chunk_size` rows in memory
- [ ] `count_used_rows()`: walk rows from bottom to find actual data extent
- [ ] Unit test: 100k-row read yields correct number of chunks (10 chunks × 10k rows)
- [ ] Unit test: chunked read produces identical data to non-chunked read
- [ ] Performance test: 500k rows in <3 seconds (`@pytest.mark.slow`)

---

### 40–45. Governance Tools (`tools/governance/`)

Each tool follows an identical pattern: argparse CLI → ExcelAgent → core logic → JSON output.

#### 40. `tools/governance/xls_clone_workbook.py`
**Purpose:** Atomic copy of workbook to `/work/` directory for safe mutation.

**CLI:** `xls_clone_workbook.py --input financials.xlsx [--output-dir ./work/]`

**Checklist:**
- [ ] Generate timestamped filename: `{name}_{YYYYMMDDTHHMMSS}_{hash[:8]}.xlsx`
- [ ] Atomic copy using `shutil.copy2()` (preserves metadata)
- [ ] Compute and return entry hash of clone
- [ ] Return JSON: `{"clone_path": "...", "source_hash": "...", "clone_hash": "..."}`
- [ ] Unit test: clone is byte-identical to source
- [ ] Unit test: clone path matches expected naming convention

#### 41. `tools/governance/xls_validate_workbook.py`
**Purpose:** OOXML compliance check, broken reference detection, circular ref scan.

**CLI:** `xls_validate_workbook.py --input workbook.xlsx`

**Checklist:**
- [ ] Load workbook and verify it opens without error
- [ ] Run `DependencyTracker.detect_circular_references()`
- [ ] Scan for `#REF!`, `#NAME?` error values in cells
- [ ] Check for orphaned named ranges
- [ ] Return JSON: `{"valid": bool, "errors": [...], "warnings": [...], "circular_refs": [...]}`

#### 42. `tools/governance/xls_approve_token.py`
**Purpose:** Generate scoped HMAC-SHA256 approval token.

**CLI:** `xls_approve_token.py --scope sheet:delete --file workbook.xlsx [--ttl 300]`

**Checklist:**
- [ ] Validate scope is in `VALID_SCOPES`
- [ ] Compute file hash of target workbook
- [ ] Generate token via `ApprovalTokenManager`
- [ ] Return JSON: `{"token": "...", "scope": "...", "expires_at": "...", "file_hash": "..."}`

#### 43. `tools/governance/xls_version_hash.py`
**Purpose:** Compute geometry hash of workbook.

**CLI:** `xls_version_hash.py --input workbook.xlsx`

**Checklist:**
- [ ] Compute both geometry hash and file hash
- [ ] Return JSON: `{"geometry_hash": "sha256:...", "file_hash": "sha256:..."}`

#### 44. `tools/governance/xls_lock_status.py`
**Purpose:** Check if a workbook is currently locked.

**CLI:** `xls_lock_status.py --input workbook.xlsx`

**Checklist:**
- [ ] Use `FileLock.is_locked()` for non-blocking check
- [ ] Return JSON: `{"locked": bool, "lock_file_exists": bool}`

#### 45. `tools/governance/xls_dependency_report.py`
**Purpose:** Full dependency graph export as JSON adjacency list.

**CLI:** `xls_dependency_report.py --input workbook.xlsx [--sheet Sheet1]`

**Checklist:**
- [ ] Build dependency graph via `DependencyTracker`
- [ ] Export via `get_adjacency_list()` and `get_stats()`
- [ ] Return JSON: `{"stats": {...}, "graph": {...}, "circular_refs": [...]}`

---

### 46–52. Read Tools (`tools/read/`)

#### 46. `tools/read/xls_get_sheet_names.py`
**CLI:** `--input sample.xlsx`
**Output:** `{"sheets": [{"index": 0, "name": "Sheet1", "visibility": "visible"}, ...]}`
**Checklist:** iterate `workbook.sheetnames` + `sheet.sheet_state`, unit test with hidden sheet

#### 47. `tools/read/xls_get_workbook_metadata.py`
**CLI:** `--input sample.xlsx`
**Output:** `{"sheet_count", "total_formulas", "named_ranges", "tables", "has_macros", "file_size_bytes"}`
**Checklist:** count formulas across all sheets, count `defined_names`, count tables, check file size

#### 48. `tools/read/xls_read_range.py`
**CLI:** `--input sample.xlsx --range A1:C10 --sheet Sheet1 [--chunked]`
**Checklist:** normal mode returns 2D array; chunked mode emits JSONL; dates → ISO 8601; 500k rows <3s

#### 49. `tools/read/xls_get_defined_names.py`
**Output:** `{"named_ranges": [{"name", "scope", "refers_to"}, ...]}`
**Checklist:** iterate `workbook.defined_names`, extract scope (workbook vs sheet)

#### 50. `tools/read/xls_get_table_info.py`
**Output:** `{"tables": [{"name", "sheet", "range", "columns", "has_totals_row", "style"}, ...]}`
**Checklist:** iterate `sheet.tables` per sheet

#### 51. `tools/read/xls_get_formula.py`
**CLI:** `--input sample.xlsx --cell A1 --sheet Sheet1`
**Output:** `{"cell": "A1", "formula": "=SUM(B1:B10)", "references": ["B1:B10"]}`
**Checklist:** check `cell.data_type`, return formula or `null`, optionally parse references

#### 52. `tools/read/xls_get_cell_style.py`
**CLI:** `--input sample.xlsx --cell A1`
**Output:** `{"font": {...}, "fill": {...}, "border": {...}, "alignment": {...}, "number_format": "..."}`
**Checklist:** serialize openpyxl style objects to JSON

---

### 53. `tests/integration/test_read_tools.py`

**Checklist:**
- [ ] Each of 7 read tools called via `subprocess.run()` against fixture
- [ ] Verify exit code 0 for all
- [ ] Verify JSON output is parseable and matches expected schema
- [ ] Chunked read of large fixture completes in <3s

---

**Phase 4 Exit Criteria:**
- [ ] All 13 tools (6 governance + 7 read) execute and return valid JSON
- [ ] Chunked read of 500k rows in <3 seconds
- [ ] All governance tools correctly integrate with core components
- [ ] Integration tests pass calling tools via subprocess (simulating agent)

---

# Phase 5: Write & Create Tools

**Duration:** 3 days
**Goal:** Implement workbook creation and data writing tools.
**Dependency:** Phase 4 (uses ExcelAgent, RangeSerializer, schemas)

## Files: 4 Tools + Integration Test

### 54. `tools/write/xls_create_new.py`
### 55. `tools/write/xls_create_from_template.py`
### 56. `tools/write/xls_write_range.py`
### 57. `tools/write/xls_write_cell.py`
### 58. `tests/integration/test_write_tools.py`

*(Detailed specifications identical to the draft reference — see draft Phase 5 items 35–39. All checklists preserved.)*

**Phase 5 Exit Criteria:**
- [ ] All 4 tools execute and return valid JSON
- [ ] Roundtrip: write data → read data → verify equality
- [ ] Type inference: dates, booleans, numbers auto-detected
- [ ] Template substitution for `{{placeholder}}` patterns works
- [ ] Formulas in data strings (starting with `=`) stored as formulas, not values

---

# Phase 6: Structural Mutation Tools

**Duration:** 8 days
**Goal:** Implement sheet/row/column manipulation with dependency checks and governance token enforcement.
**Dependency:** Phase 5 + Phase 2 (DependencyTracker) + Phase 3 (TokenManager, AuditTrail)

## Files: 8 Tools + Integration Test

### 59. `tools/structure/xls_add_sheet.py`
### 60. `tools/structure/xls_delete_sheet.py` ⚠️ (token: `sheet:delete`)
### 61. `tools/structure/xls_rename_sheet.py` ⚠️ (token: `sheet:rename`)
### 62. `tools/structure/xls_insert_rows.py`
### 63. `tools/structure/xls_delete_rows.py` ⚠️ (token: `range:delete`)
### 64. `tools/structure/xls_insert_columns.py`
### 65. `tools/structure/xls_delete_columns.py` ⚠️ (token: `range:delete`)
### 66. `tools/structure/xls_move_sheet.py`
### 67. `tests/integration/test_structure_tools.py`

*(Detailed specifications match draft Phase 6 items 40–48. Key addition from research: all token-gated tools include denial-with-prescriptive-guidance in error responses.)*

**Critical Enhancement from Research:** Every `⚠️` tool implements the denial-with-guidance pattern:
```json
{
    "status": "denied",
    "exit_code": 1,
    "denial_reason": "Operation would break 7 formula references across 3 sheets",
    "guidance": "Run xls_update_references.py --target='Sheet1!A5:A10' before retrying",
    "impact": {"broken_references": 7, "affected_sheets": ["Sheet1", "Sheet2", "Summary"]},
    "stale_output_warning": "Do not proceed with cached data from prior reads of affected cells"
}
```

**Phase 6 Exit Criteria:**
- [ ] All 8 tools execute; token-gated tools reject without valid token (exit code 4)
- [ ] `xls_delete_sheet` runs dependency check, returns impact report if refs would break
- [ ] `xls_rename_sheet` auto-updates all cross-sheet formula references
- [ ] Insert/delete rows correctly adjusts formula references in affected sheets
- [ ] Audit trail logs all destructive operations

---

# Phase 7: Cell Operations

**Duration:** 3 days
**Dependency:** Phase 6

### 68–71. `tools/cells/xls_merge_cells.py`, `xls_unmerge_cells.py`, `xls_delete_range.py` ⚠️, `xls_update_references.py`
### 72. `tests/integration/test_cell_operations.py`

*(Specifications match draft Phase 7 items 49–53.)*

---

# Phase 8: Formulas & Calculation Engine

**Duration:** 5 days
**Goal:** Implement two-tier calculation engine and formula manipulation tools.
**Dependency:** Phase 7 + Phase 2 (DependencyTracker)

## Files to Create

### 73. `src/excel_agent/calculation/__init__.py`

### 74. `src/excel_agent/calculation/tier1_engine.py`

**Purpose:** In-process calculation using the `formulas` library. `formulas` implements an interpreter for Excel formulas, which parses and compile Excel formulas expressions. Moreover, it compiles Excel workbooks to python and executes without using the Excel COM server. Hence, Excel is not needed.

**Key capability:** Spreadsheet models can also be converted into a portable JSON representation. This is useful when the model needs to be versioned, inspected, or executed without the original workbook.

**Interface:**
```python
from pathlib import Path
from dataclasses import dataclass

@dataclass
class CalculationResult:
    formula_count: int
    calculated_count: int
    error_count: int
    unsupported_functions: list[str]   # Functions formulas library doesn't support
    recalc_time_ms: float
    engine: str                         # "tier1_formulas" or "tier2_libreoffice"

class Tier1Calculator:
    """In-process Excel calculation via `formulas` library."""

    def __init__(self, workbook_path: Path):
        ...

    def calculate(self, *, circular: bool = False) -> CalculationResult:
        """Calculates all formulas. Set circular=True for workbooks with circular refs."""
        ...

    def get_cell_value(self, cell_ref: str) -> object:
        """Returns calculated value for specific cell (e.g., 'Sheet1!A1')."""
        ...

    def get_unsupported_functions(self) -> list[str]:
        """Returns list of Excel functions not supported by Tier 1."""
        ...

    def write_results(self, output_path: Path) -> None:
        """Writes recalculated workbook to disk."""
        ...
```

**Checklist:**
- [ ] Load workbook via `formulas.ExcelModel().loads(path).finish()` — If you have or could have circular references, add `circular=True` to finish method.
- [ ] Call `xl_model.calculate()` to evaluate all formulas
- [ ] Write results via `xl_model.write(dirpath=output_dir)`
- [ ] Track unsupported functions: catch exceptions per formula, collect function names
- [ ] Return `CalculationResult` with timing, counts, error list
- [ ] Unit test: `=2+2` evaluates to `4`
- [ ] Unit test: `=SUM(A1:A10)` evaluates correctly
- [ ] Unit test: `=IF(A1>5, "high", "low")` evaluates correctly
- [ ] Unit test: unsupported function returns error marker (does not crash)
- [ ] Performance test: 1000 formulas in <500ms

### 75. `src/excel_agent/calculation/tier2_libreoffice.py`

**Interface:**
```python
class Tier2Calculator:
    def __init__(self, *, soffice_path: Path | None = None):
        ...
    def recalculate(self, workbook_path: Path, output_path: Path, *, timeout: int = 60) -> CalculationResult:
        ...
    @staticmethod
    def is_available() -> bool:
        """Checks if LibreOffice is installed and accessible."""
        ...
```

**Checklist:**
- [ ] Auto-detect `soffice` binary on PATH or common install locations
- [ ] Execute `soffice --headless --calc --convert-to xlsx --outdir <dir> <file>`
- [ ] `subprocess.run()` with `timeout` parameter
- [ ] Capture stderr for error detection
- [ ] `is_available()` runs `soffice --version` to check
- [ ] Unit test (marked `@pytest.mark.libreoffice`): recalc produces valid file
- [ ] Unit test: timeout handling for very large files

### 76. `src/excel_agent/calculation/error_detector.py`

**Checklist:**
- [ ] Scan all cells for error values: `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`
- [ ] Return list of `{"sheet", "cell", "error", "formula"}` dicts
- [ ] Unit test: detect each error type in fixture

### 77. `scripts/recalc.py`

LibreOffice headless wrapper script. (Specifications match draft item 56.)

### 78–83. Formula Tools (`tools/formulas/`)

- `xls_set_formula.py` — syntax validation via openpyxl Tokenizer
- `xls_recalculate.py` — auto: try Tier 1, fallback to Tier 2; explicit `--tier` flag
- `xls_detect_errors.py` — uses `error_detector.py`
- `xls_convert_to_values.py` ⚠️ — token `formula:convert`; irreversible
- `xls_copy_formula_down.py` — auto-adjust relative references
- `xls_define_name.py` — create/update named ranges

### 84. `tests/integration/test_calculation.py`

**Phase 8 Exit Criteria:**
- [ ] Tier 1 correctly calculates SUM, AVERAGE, IF, VLOOKUP (common functions)
- [ ] Tier 2 LibreOffice bridge works in CI
- [ ] Auto-fallback: Tier 1 failure → Tier 2 invocation
- [ ] Error detection finds all 7 error types
- [ ] Convert-to-values replaces formulas with computed results

---

# Phase 9: Macro Safety Tools

**Duration:** 3 days
**Dependency:** Phase 3 (TokenManager), Phase 0 (oletools installed)

### 85. `src/excel_agent/core/macro_handler.py`

**Purpose:** Safe VBA inspection behind a `Protocol` abstraction (future-proofed against oletools maintenance risk).

**Design:** oletools is a package of python tools to analyze Microsoft OLE2 files. It is based on the olefile parser. It also provides tools to analyze RTF files and files based on the OpenXML format such as MS Office 2007+ documents. oletools can detect, extract and analyse VBA macros, OLE objects, Excel 4 macros (XLM) and DDE links.

**Interface:**
```python
from typing import Protocol

class MacroAnalyzer(Protocol):
    """Abstraction layer for swappable macro analysis backends."""
    def has_macros(self, path: Path) -> bool: ...
    def extract_modules(self, path: Path) -> list[dict]: ...
    def detect_auto_exec(self, path: Path) -> list[dict]: ...
    def detect_suspicious(self, path: Path) -> list[dict]: ...
    def scan_risk(self, path: Path) -> dict: ...
    def has_digital_signature(self, path: Path) -> bool: ...

class OletoolsMacroAnalyzer:
    """oletools-backed implementation of MacroAnalyzer."""
    ...
```

**Checklist:**
- [ ] `Protocol` interface defined (`MacroAnalyzer`) for future backend swaps
- [ ] `OletoolsMacroAnalyzer` implements all methods using `oletools.olevba`
- [ ] `has_macros()`: uses `olevba.detect_vba_macros()`
- [ ] `extract_modules()`: uses `olevba.extract_macros()`, returns `[{"name", "type", "code_size"}]`
- [ ] `detect_auto_exec()`: uses `olevba.detect_autoexec()`, returns triggers
- [ ] `detect_suspicious()`: uses `olevba.detect_suspicious()`, returns suspicious keywords
- [ ] `scan_risk()`: aggregates auto_exec + suspicious + IOCs → risk level (low/medium/high)
- [ ] `has_digital_signature()`: checks for `vbaProjectSignature.bin` in ZIP structure
- [ ] **Hard pre-condition on inject:** `scan_risk()` MUST be called on any `.bin` before injection
- [ ] Unit test: `.xlsm` with macros detected
- [ ] Unit test: `.xlsx` without macros returns False
- [ ] Unit test: AutoOpen trigger detected
- [ ] Unit test: Shell keyword detected as suspicious

### 86–90. Macro Tools (`tools/macros/`)

- `xls_has_macros.py` — boolean check
- `xls_inspect_macros.py` — module list + signature status
- `xls_validate_macro_safety.py` — risk scan with IOCs
- `xls_remove_macros.py` ⚠️⚠️ — double-token, converts `.xlsm` → `.xlsx`
- `xls_inject_vba_project.py` ⚠️ — injects pre-extracted `.bin` (MUST scan first)

### 91. `tests/integration/test_macro_tools.py`

---

# Phase 10: Objects & Visualization

**Duration:** 4 days
**Dependency:** Phase 5

### 92–96. Object Tools (`tools/objects/`)

- `xls_add_table.py` — ListObject with style
- `xls_add_chart.py` — Bar, Line, Pie, Scatter
- `xls_add_image.py` — with aspect ratio preservation
- `xls_add_comment.py` — threaded comments
- `xls_set_data_validation.py` — dropdowns, constraints

### 97. `tests/integration/test_objects.py`

---

# Phase 11: Formatting & Style

**Duration:** 3 days
**Dependency:** Phase 5

### 98–102. Formatting Tools (`tools/formatting/`)

- `xls_format_range.py` — comprehensive JSON-driven formatting
- `xls_set_column_width.py` — auto-fit or fixed
- `xls_freeze_panes.py` — freeze point
- `xls_apply_conditional_formatting.py` — ColorScale, DataBar, IconSet
- `xls_set_number_format.py` — currency, %, date format codes

### 103. `tests/integration/test_formatting.py`

---

# Phase 12: Export & Interop

**Duration:** 2 days
**Dependency:** Phase 8 (needs LibreOffice for PDF)

### 104–106. Export Tools (`tools/export/`)

- `xls_export_pdf.py` — LibreOffice headless conversion
- `xls_export_csv.py` — with encoding control
- `xls_export_json.py` — records, values, or columns format

### 107. `tests/integration/test_export.py`

---

# Phase 13: End-to-End Integration & Documentation

**Duration:** 3 days
**Goal:** Full workflow simulation, comprehensive docs, agent-ready validation.
**Dependency:** All previous phases

## Files to Create

### 108. `tests/integration/test_clone_modify_workflow.py`

**Full Agent Workflow Simulation:**
```
1. xls_clone_workbook.py → get clone_path
2. xls_get_workbook_metadata.py → understand structure
3. xls_read_range.py → read data
4. xls_write_range.py → modify data
5. xls_insert_rows.py → add rows
6. xls_recalculate.py → recalc formulas
7. xls_validate_workbook.py → verify integrity
8. xls_export_pdf.py → generate output
```

**Checklist:**
- [ ] All tools called via `subprocess.run()` (simulating AI agent tool execution)
- [ ] JSON outputs parsed and chained between steps
- [ ] Final workbook passes validation
- [ ] Total workflow time measured and reported

### 109. `tests/integration/test_formula_dependency_workflow.py`

**Dependency-Aware Deletion Workflow:**
```
1. xls_dependency_report.py → get impact analysis
2. xls_delete_sheet.py → attempt delete → denied with guidance
3. xls_update_references.py → fix references per guidance
4. xls_approve_token.py → generate token
5. xls_delete_sheet.py --token → successful deletion
6. xls_validate_workbook.py → no broken references
```

### 110. `docs/DESIGN.md` — Architecture blueprint (this document)
### 111. `docs/API.md` — CLI reference for all 53 tools
### 112. `docs/WORKFLOWS.md` — 5 common agent workflow recipes with JSON examples
### 113. `docs/GOVERNANCE.md` — Token scopes, audit trail, safety protocols
### 114. `docs/DEVELOPMENT.md` — Contributing guide, code standards, how to add a new tool

---

# Phase 14: Performance Optimization & Security Hardening

**Duration:** 3 days
**Goal:** Meet all performance benchmarks, harden error handling, security audit, cross-platform validation.
**Dependency:** Phase 13

### Tasks

**Performance Benchmarks:**
- [ ] Read 500k rows in <3s (chunked)
- [ ] Write 500k rows in <5s (openpyxl write-only mode)
- [ ] Dependency graph for 10-sheet/1000-formula workbook in <5s
- [ ] Tier 1 recalc of 1000 formulas in <500ms
- [ ] File lock acquire/release in <100ms

**Error Handling Hardening:**
- [ ] Every tool catches all exceptions → JSON error response with correct exit code
- [ ] No unhandled exceptions leak stack traces to stdout (agent sees only JSON)
- [ ] `FileNotFoundError` → exit code 2
- [ ] `LockContentionError` → exit code 3
- [ ] `PermissionDeniedError` → exit code 4
- [ ] All others → exit code 5 with sanitized message

**Security Audit:**
- [ ] Verify `defusedxml` is imported and active (openpyxl uses it when installed)
- [ ] No arbitrary code execution: formulas treated as data strings, never `eval()`'d
- [ ] Token secret key never logged or exposed in JSON output
- [ ] Audit trail is append-only (no delete/modify capability)
- [ ] LibreOffice subprocess runs with `timeout` (no infinite hangs)
- [ ] File paths validated: reject `../` traversal, symlink following
- [ ] Dependency pinning with hashes in `requirements.txt`

**Cross-Platform:**
- [ ] CI on Ubuntu (Linux)
- [ ] Manual verification on macOS (or add to CI matrix)
- [ ] File locking tested on Windows (manual or WSL)
- [ ] All paths use `pathlib.Path` (never raw string concatenation)

**Final QA:**
- [ ] `pytest --cov=excel_agent --cov-report=html` → coverage >90%
- [ ] `black --check src/ tools/ tests/` → pass
- [ ] `mypy --strict src/` → no errors
- [ ] `ruff check src/ tools/` → no errors
- [ ] Manually test 10 random tools via CLI
- [ ] Verify all 53 entry points are registered and show `--help`

---

# Summary: Master Execution Plan Overview

| Phase | Duration | Files | Tools | Key Deliverable |
|:---|:---|:---|:---|:---|
| **Phase 0:** Scaffolding | 2 days | 16 | 0 | Project structure, CI, deps |
| **Phase 1:** Core Foundation | 5 days | 10 | 0 | ExcelAgent, FileLock, RangeSerializer, VersionHash |
| **Phase 2:** Dependency Engine | 5 days | 8 | 0 | DependencyTracker, JSON schemas |
| **Phase 3:** Governance Layer | 3 days | 4 | 0 | ApprovalTokenManager, AuditTrail |
| **Phase 4:** Governance + Read Tools | 5 days | 15 | 13 | 6 governance + 7 read CLI tools |
| **Phase 5:** Write Tools | 3 days | 5 | 4 | Create + write tools |
| **Phase 6:** Structure Tools | 8 days | 9 | 8 | Token-gated sheet/row/col mutation |
| **Phase 7:** Cell Operations | 3 days | 5 | 4 | Merge, unmerge, delete range, update refs |
| **Phase 8:** Formulas + Calc | 5 days | 12 | 6 | Two-tier calc engine + formula tools |
| **Phase 9:** Macro Safety | 3 days | 7 | 5 | oletools-backed VBA analysis + management |
| **Phase 10:** Objects | 4 days | 6 | 5 | Tables, charts, images, comments, validation |
| **Phase 11:** Formatting | 3 days | 6 | 5 | Styles, column width, freeze, conditional format |
| **Phase 12:** Export | 2 days | 4 | 3 | PDF, CSV, JSON export |
| **Phase 13:** Integration + Docs | 3 days | 7 | 0 | E2E workflows, full documentation |
| **Phase 14:** Hardening | 3 days | 0 | 0 | Performance, security, cross-platform |
| **TOTAL** | **~57 days (≈12 weeks)** | **~114 files** | **53 tools** | |

---

## Critical Path Dependencies

```
Phase 0 (Scaffold)
  └→ Phase 1 (Core: Agent, Lock, Serializer, Hash)
       └→ Phase 2 (Dependency Graph + Schemas)
            └→ Phase 3 (Tokens + Audit)
                 ├→ Phase 4 (Governance Tools + Read Tools)
                 │    └→ Phase 5 (Write Tools)
                 │         ├→ Phase 6 (Structure Tools) ← needs Phase 2 + 3
                 │         │    └→ Phase 7 (Cell Operations)
                 │         ├→ Phase 10 (Objects)
                 │         └→ Phase 11 (Formatting)
                 ├→ Phase 8 (Formulas + Calc Engine) ← needs Phase 2
                 │    └→ Phase 12 (Export) ← needs Phase 8 for LibreOffice
                 └→ Phase 9 (Macro Safety) ← needs Phase 3
                      └→ Phase 13 (E2E + Docs) ← needs ALL above
                           └→ Phase 14 (Hardening)
```

**Parallelizable pairs** (if multiple developers available):
- Phase 10 (Objects) + Phase 11 (Formatting) can run in parallel
- Phase 8 (Calc Engine) + Phase 9 (Macros) can run in parallel
- Phase 6 (Structure) + Phase 8 (Calc) can overlap if core is done

---

## Final Deliverables Checklist

- [ ] **53** production-ready CLI tools with `--help`, JSON output, exit codes 0–5
- [ ] **Core library** (`excel_agent`): ExcelAgent, DependencyTracker, MacroAnalyzer, ApprovalTokenManager, AuditTrail
- [ ] **Two-tier calculation engine**: `formulas` (in-process, ~50ms for 10k formulas) + LibreOffice headless (full fidelity)
- [ ] **7 governance token scopes** with HMAC-SHA256, TTL, nonce, file-hash binding, `compare_digest()`
- [ ] **Pluggable audit trail**: JSONL default, Null for testing, Composite for fan-out
- [ ] **Full test suite**: >90% coverage, unit + integration + property-based + performance
- [ ] **Complete documentation**: DESIGN.md, API.md, WORKFLOWS.md, GOVERNANCE.md, DEVELOPMENT.md
- [ ] **CI/CD pipeline**: GitHub Actions, multi-Python matrix, LibreOffice installed, coverage gate
- [ ] **Cross-platform**: Linux (CI), macOS (manual), Windows (manual)
- [ ] **Security hardened**: defusedxml mandatory, no `eval()`, token secrets protected, path traversal blocked

---

**This Master Execution Plan is ready for Phase 3 (VALIDATE) — your explicit confirmation before any code is written.**

Shall I proceed to implementation of Phase 0, or would you like to discuss any adjustments to this plan?
