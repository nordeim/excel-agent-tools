# CLAUDE.md - AI Coding Agent Briefing
## excel-agent-tools v1.0.0

**Last Updated:** April 10, 2026
**Status:** Production-Ready | All 53 Tools Implemented | Phase 14 Complete
**Current Phase:** Phase 14 Complete (Hardening, Security & Release Preparation)

---

## Executive Summary

`excel-agent-tools` is a **production-grade Python CLI suite** of 53 stateless tools enabling AI agents to safely read, write, calculate, and export Excel workbooks without Microsoft Excel or COM dependencies.

### Key Metrics
| Metric | Value |
|--------|-------|
| **Total Tools** | 53 (100% implemented) |
| **Source Files** | 86 Python modules + 8 new Phase 14 modules |
| **Test Files** | 36 test modules |
| **Total Tests** | 430+ tests (343 unit tests passing) |
| **Coverage** | >90% |
| **Documentation** | 10+ MD files |
| **Entry Points** | 53 CLI commands |
| **SDK** | AgentClient with retry/backoff |

### Design Philosophy
1. **Governance-First**: Destructive ops require HMAC-SHA256 scoped tokens
2. **Formula Integrity**: Dependency graphs block mutations breaking `#REF!` chains
3. **AI-Native Contracts**: Strict JSON stdout, standardized exit codes (0-5)
4. **Headless Operation**: Zero Excel dependency, runs on any server
5. **Distributed-Ready**: Pluggable state backends (Redis) for multi-agent deployments

---

## Phase 14 Accomplishments (April 10, 2026)

### 1. Agent Orchestration SDK (`src/excel_agent/sdk/`)
**New Feature**: Python client for simplified AI framework integration

```python
from excel_agent.sdk import AgentClient

client = AgentClient(secret_key="your-secret")

# Clone and modify
clone_path = client.clone("data.xlsx", output_dir="./work")
data = client.read_range(clone_path, "A1:C10")

# Safe structural edit with automatic retry
from excel_agent.sdk import ImpactDeniedError

try:
    token = client.generate_token("sheet:delete", clone_path)
    client.run("structure.xls_delete_sheet",
              input=clone_path, name="OldSheet", token=token)
except ImpactDeniedError as e:
    print(f"Denied: {e.guidance}")
    print(f"Impact: {e.impact}")
```

**Key Classes:**
- `AgentClient`: Main SDK client with retry logic
- `ImpactDeniedError`: Structured denial with guidance
- `TokenRequiredError`: Token authentication error
- `ToolExecutionError`: General execution failure

### 2. Pre-commit Configuration (`.pre-commit-config.yaml`)
**Security Hardening**: Automated code quality and security scanning

```yaml
# Hooks include:
# - detect-secrets: Prevents accidental secret commits
# - detect-private-key: Blocks private key files
# - black: Code formatting
# - ruff: Linting
# - mypy: Type checking
```

**Usage:**
```bash
pre-commit install  # One-time setup
pre-commit run --all-files  # Manual run
git commit -m "message"  # Automatic hooks on commit
```

### 3. Distributed State Protocols (`src/excel_agent/governance/`)
**Scalability Feature**: Pluggable storage for multi-agent deployments

**New Files:**
- `stores.py`: Protocol definitions (`TokenStore`, `AuditBackend`)
- `backends/redis.py`: Redis implementation

**Usage:**
```python
from excel_agent.governance.backends.redis import RedisTokenStore
from excel_agent.governance.token_manager import ApprovalTokenManager

# Distributed token tracking
redis_store = RedisTokenStore("redis://localhost:6379")
manager = ApprovalTokenManager(
    secret="my-secret",
    nonce_store=redis_store
)
```

### 4. Dependency Version Updates
**Accuracy Improvement**: Pinned to actual installed versions

| Package | Old Version | New Version |
|---------|-------------|-------------|
| pandas | >=2.1.0 | >=3.0.0 |
| jsonschema | >=4.23.0 | >=4.26.0 |
| pytest | >=8.0.0 | >=9.0.0 |
| black | >=24.0.0 | >=26.0.0 |

**New Optional Extras:**
```toml
[project.optional-dependencies]
redis = ["redis>=6.0.0"]
security = [
    "cyclonedx-python-lib>=9.0.0",
    "detect-secrets>=1.5.0",
    "safety>=3.7.0",
]
```

### 5. Chunked I/O Test Fix
**Bug Fix**: Corrected test expectation for JSONL chunked output

```python
# Fixed in tests/integration/test_clone_modify_workflow.py:306
# Before: assert chunk.get("status") == "success"  # Wrong - JSONL has no envelope
# After:  assert "values" in chunk  # Correct - direct data
```

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│ AI Agent / Orchestrator │
│ (Claude, GPT, LangChain, AutoGen) │
└───────────────────────┬───────────────────────────────────────────┘
│ JSON stdin/stdout
┌───────────────────────▼───────────────────────────────────────────┐
│ Agent SDK Layer (NEW in Phase 14) │
│ ┌─────────────────────────────────────────┐ │
│ │ AgentClient: retry, parse, token mgmt │ │
│ └─────────────────────────────────────────┘ │
└───────────────────────┬───────────────────────────────────────────┘
│ subprocess
┌───────────────────────▼───────────────────────────────────────────┐
│ CLI Tool Layer (53 Tools) │
│ ┌──────────┬──────────┬──────────┬──────────┬──────────┐ │
│ │Governance│ Read │ Write │ Structure│ Cells │ │
│ │ (6) │ (7) │ (4) │ (8) │ (4) │ │
│ ├──────────┼──────────┼──────────┼──────────┼──────────┤ │
│ │ Formulas │ Objects │ Formatting│ Macros │ Export │ │
│ │ (6) │ (5) │ (5) │ (5) │ (3) │ │
│ └──────────┴──────────┴──────────┴──────────┴──────────┘ │
└───────────────────────┬───────────────────────────────────────────┘
│ _tool_base.run_tool()
┌───────────────────────▼───────────────────────────────────────────┐
│ Core Hub Layer │
│ ┌─────────────────┬─────────────────┬──────────────────┐ │
│ │ ExcelAgent │ DependencyTrack │ TokenManager │ │
│ │ (Context Mgr) │ (Graph) │ (HMAC-SHA256) │ │
│ ├─────────────────┼─────────────────┼──────────────────┤ │
│ │ FileLock │ RangeSerial │ AuditTrail │ │
│ │ (OS-level) │ (A1/R1C1) │ (JSONL) │ │
│ ├─────────────────┼─────────────────┼──────────────────┤ │
│ │ VersionHash │ MacroHandler │ ChunkedIO │ │
│ │ (Geometry) │ (oletools) │ (Streaming) │ │
│ ├─────────────────┼─────────────────┼──────────────────┤ │
│ │ Distributed │ Redis Backend │ InMemory │ │
│ │ State (NEW) │ (NEW) │ (Default) │ │
│ └─────────────────┴─────────────────┴──────────────────┘ │
└───────────────────────┬───────────────────────────────────────────┘
│ Libraries
┌───────────────────────▼───────────────────────────────────────────┐
│ Library Layer │
│ ┌──────────┬──────────┬──────────┬──────────┬──────────┐ │
│ │openpyxl │ formulas │ oletools │defusedxml│ jsonschema│ │
│ │(I/O) │(Tier 1) │(Macros) │(Security)│(Schemas) │ │
│ └──────────┴──────────┴──────────┴──────────┴──────────┘ │
└───────────────────────────────────────────────────────────────────┘
```

---

## Core Components Deep Dive

### 1. ExcelAgent (src/excel_agent/core/agent.py)

**Purpose**: Stateful context manager for safe workbook manipulation

**Lifecycle**:
```
__enter__:
1. Acquire FileLock (exclusive, timeout=30s)
2. Load workbook via openpyxl (keep_vba=True, data_only=False)
3. Compute entry file hash (SHA-256 for concurrent modification detection)
4. Compute geometry hash (structure + formulas, excludes values)

__exit__ (success, mode='rw'):
1. Re-read file hash from disk
2. If changed → raise ConcurrentModificationError (do NOT save)
3. Save workbook + fsync
4. Release lock

__exit__ (exception):
1. Release lock WITHOUT saving
2. Re-raise exception
```

**Key Properties**:
- Thread-safe: One agent per workbook per process
- Fail-safe: Lock released even on exception
- Formula-preserving: Always loads formulas (not cached values)

### 2. Exit Codes (src/excel_agent/utils/exit_codes.py)

| Code | Name | Meaning | Agent Action |
|------|------|---------|-------------|
| 0 | SUCCESS | Operation completed | Parse JSON, proceed |
| 1 | VALIDATION_ERROR | Input rejected | Fix input, retry |
| 2 | FILE_NOT_FOUND | Path invalid | Verify path |
| 3 | LOCK_CONTENTION | File locked | Exponential backoff |
| 4 | PERMISSION_DENIED | Token invalid | Generate new token |
| 5 | INTERNAL_ERROR | Unexpected failure | Alert operator |

### 3. JSON Response Schema

Every tool outputs exactly one JSON object:

```json
{
  "status": "success" | "error" | "warning" | "denied",
  "exit_code": 0,
  "timestamp": "2026-04-08T14:30:22Z",
  "workbook_version": "sha256:abc...",
  "data": { /* tool-specific */ },
  "impact": {
    "cells_modified": 0,
    "formulas_updated": 0,
    "rows_inserted": 0
  },
  "warnings": [],
  "guidance": "..." // Present when denied
}
```

### 4. Tool Base Pattern (src/excel_agent/tools/_tool_base.py)

All tools follow this pattern:

```python
def _run() -> dict:
    parser = create_parser("Description")
    parser.add_argument("--input", required=True)
    args = parser.parse_args()

    with ExcelAgent(path, mode="rw") as agent:
        # Core logic here
        return build_response(
            "success",
            {"result": "..."},
            impact={"cells_modified": n}
        )

def main() -> None:
    run_tool(_run)
```

---

## Tool Catalog (53 Tools)

### Governance (6 tools)
| Tool | CLI | Token Required | Purpose |
|------|-----|----------------|---------|
| xls-clone-workbook | `--input X --output-dir Y` | No | Atomic copy to /work/ |
| xls-validate-workbook | `--input X` | No | OOXML compliance check |
| xls-approve-token | `--scope S --file X` | No | Generate HMAC token |
| xls-version-hash | `--input X` | No | Compute geometry hash |
| xls-lock-status | `--input X` | No | Check lock state |
| xls-dependency-report | `--input X` | No | Export dependency graph |

### Read (7 tools)
| Tool | CLI | Description |
|------|-----|-------------|
| xls-read-range | `--input X --range A1:C10 [--chunked]` | Extract data as JSON |
| xls-get-sheet-names | `--input X` | List all sheets |
| xls-get-workbook-metadata | `--input X` | High-level stats |
| xls-get-defined-names | `--input X` | Named ranges |
| xls-get-table-info | `--input X` | Table objects |
| xls-get-cell-style | `--input X --cell A1` | Full style JSON |
| xls-get-formula | `--input X --cell A1` | Formula string |

### Write (4 tools)
| Tool | CLI | Notes |
|------|-----|-------|
| xls-create-new | `--output X [--sheets A,B]` | Create blank workbook |
| xls-create-from-template | `--template T --output X --vars JSON` | Substitute {{vars}} |
| xls-write-range | `--input X --range A1 --data JSON` | 2D array write |
| xls-write-cell | `--input X --cell A1 --value V [--type T]` | Single cell |

### Structure (8 tools) - ⚠️ Token Required
| Tool | Scope | Description |
|------|-------|-------------|
| xls-add-sheet | - | Add new sheet |
| xls-delete-sheet | `sheet:delete` | Remove sheet with impact check |
| xls-rename-sheet | `sheet:rename` | Rename + update refs |
| xls-insert-rows | - | Insert with style inheritance |
| xls-delete-rows | `range:delete` | Delete with impact check |
| xls-insert-columns | - | Insert columns |
| xls-delete-columns | `range:delete` | Delete columns |
| xls-move-sheet | - | Reorder sheets |

### Cells (4 tools)
| Tool | Token | Description |
|------|-------|-------------|
| xls-merge-cells | No | Merge range |
| xls-unmerge-cells | No | Restore grid |
| xls-delete-range | `range:delete` | Delete with shift |
| xls-update-references | No | Batch update refs |

### Formulas (6 tools)
| Tool | Token | Description |
|------|-------|-------------|
| xls-set-formula | No | Inject formula |
| xls-recalculate | No | Tier 1→Tier 2 fallback |
| xls-detect-errors | No | Scan for #REF!, etc. |
| xls-convert-to-values | `formula:convert` | Replace with values |
| xls-copy-formula-down | No | Auto-fill |
| xls-define-name | No | Create named range |

### Objects (5 tools)
| Tool | Description |
|------|-------------|
| xls-add-table | Convert range to Table |
| xls-add-chart | Bar/Line/Pie/Scatter |
| xls-add-image | Insert with aspect ratio |
| xls-add-comment | Threaded comments |
| xls-set-data-validation | Dropdown/constraints |

### Formatting (5 tools)
| Tool | Description |
|------|-------------|
| xls-format-range | JSON-driven formatting |
| xls-set-column-width | Auto-fit or fixed |
| xls-freeze-panes | Freeze rows/cols |
| xls-apply-conditional-formatting | ColorScale/DataBar/IconSet |
| xls-set-number-format | Currency, %, dates |

### Macros (5 tools) - ⚠️⚠️ Double Token
| Tool | Token | Description |
|------|-------|-------------|
| xls-has-macros | No | Boolean check |
| xls-inspect-macros | No | Module + signature |
| xls-validate-macro-safety | No | Risk scan |
| xls-remove-macros | `macro:remove` ×2 | Strip VBA |
| xls-inject-vba-project | `macro:inject` | Inject .bin |

### Export (3 tools)
| Tool | CLI | Notes |
|------|-----|-------|
| xls-export-pdf | `--input X --outfile Y` | LibreOffice headless |
| xls-export-csv | `--input X --outfile Y` | UTF-8 |
| xls-export-json | `--input X --outfile Y --orient O` | records/values/columns |

---

## Project Structure

```
excel-agent-tools/
├── 📄 pyproject.toml                    # 53 entry points, deps, tool configs
├── 📄 README.md                         # Project overview
├── 📄 Project_Architecture_Document.md  # Deep architecture
├── 📄 CLAUDE.md                         # THIS FILE - Agent briefing
├── 📄 CHANGELOG.md                      # Phase 14 additions
├── 📄 .pre-commit-config.yaml           # NEW: Pre-commit hooks
│
├── 📂 src/excel_agent/
│ ├── 📄 __init__.py                     # Lazy imports, version 1.0.0
│ │
│ ├── 📂 core/                           # Foundation layer
│ │ ├── 📄 agent.py                      # ExcelAgent context manager
│ │ ├── 📄 locking.py                    # FileLock (fcntl/msvcrt)
│ │ ├── 📄 serializers.py                # RangeSerializer (A1/R1C1/Named/Table)
│ │ ├── 📄 dependency.py               # DependencyTracker + Tarjan SCC
│ │ ├── 📄 version_hash.py              # SHA-256 geometry hashing
│ │ ├── 📄 formula_updater.py            # Reference shifting
│ │ ├── 📄 chunked_io.py                 # Streaming for >100k rows
│ │ ├── 📄 type_coercion.py              # JSON → Python types
│ │ └── 📄 style_serializer.py          # Style serialization
│ │
│ ├── 📂 governance/                     # Security & Compliance
│ │ ├── 📄 token_manager.py              # ApprovalTokenManager (HMAC-SHA256)
│ │ │                                    #   + NonceStore support (Phase 14)
│ │ ├── 📄 audit_trail.py                # AuditTrail backends
│ │ ├── 📄 stores.py                     # NEW: TokenStore/AuditBackend Protocols
│ │ ├── 📂 backends/                     # NEW: Pluggable backends
│ │ │ ├── 📄 __init__.py
│ │ │ └── 📄 redis.py                    # NEW: Redis implementations
│ │ └── 📂 schemas/                      # JSON Schema files
│ │
│ ├── 📂 calculation/                    # Two-tier engine
│ │ ├── 📄 tier1_engine.py               # `formulas` library wrapper
│ │ ├── 📄 tier2_libreoffice.py          # LibreOffice headless
│ │ └── 📄 error_detector.py             # Formula error scanner
│ │
│ ├── 📂 sdk/                           # NEW: Agent Orchestration SDK
│ │ ├── 📄 __init__.py                   # NEW: SDK exports
│ │ └── 📄 client.py                     # NEW: AgentClient + exceptions
│ │
│ ├── 📂 utils/                          # Shared utilities
│ │ ├── 📄 exit_codes.py                 # ExitCode enum (0-5)
│ │ ├── 📄 json_io.py                    # build_response(), ExcelAgentEncoder
│ │ ├── 📄 cli_helpers.py                # argparse patterns
│ │ ├── 📄 exceptions.py                 # ExcelAgentError hierarchy
│ │ └── 📄 __init__.py
│ │
│ └── 📂 tools/                          # 53 CLI tools (10 categories)
│   ├── 📄 _tool_base.py                 # Base runner for all tools
│   ├── 📂 governance/                   # 6 tools
│   ├── 📂 read/                         # 7 tools
│   ├── 📂 write/                        # 4 tools
│   ├── 📂 structure/                    # 8 tools
│   ├── 📂 cells/                        # 4 tools
│   ├── 📂 formulas/                     # 6 tools
│   ├── 📂 objects/                      # 5 tools
│   ├── 📂 formatting/                   # 5 tools
│   ├── 📂 macros/                       # 5 tools
│   └── 📂 export/                       # 3 tools
│
├── 📂 tests/
│ ├── 📄 __init__.py
│ ├── 📄 conftest.py                     # Shared fixtures
│ ├── 📂 unit/                           # 20+ test modules
│ ├── 📂 integration/                    # 10+ test modules
│ └── 📂 property/                       # Hypothesis fuzzing tests
│
├── 📂 docs/
│ ├── 📄 DESIGN.md                       # Architecture blueprint
│ ├── 📄 API.md                          # CLI reference (all 53 tools)
│ ├── 📄 WORKFLOWS.md                    # 5 production recipes
│ ├── 📄 GOVERNANCE.md                   # Token lifecycle
│ └── 📄 DEVELOPMENT.md                  # Contributor guide
│
└── 📂 scripts/
  └── 📄 install_libreoffice.sh          # CI setup script
```

---

## Development Workflow

### Standard Operating Procedure (Meticulous Approach)

```
┌─────────────────────────────────────────────────────────────────┐
│                                                                 │
│  ANALYZE → PLAN → VALIDATE → IMPLEMENT → VERIFY → DELIVER      │
│                                                                 │
│  • Deep requirement   • Phases,    • Write code   • Test      │
│    mining               checklists   modular        coverage    │
│  • Research           • Decision     documented                 │
│  • Risk assessment    points       • Continuous                 │
│  • User testing                        testing                  │
│    confirm     • Follow style                                   │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### Adding a New Tool

1. **Create** `src/excel_agent/tools/<category>/xls_<name>.py`
2. **Implement** `_run() -> dict` following `_tool_base` pattern
3. **Register** entry point in `pyproject.toml` under `[project.scripts]`
4. **Test** in `tests/unit/test_<category>_tools.py`
5. **Integration** test in `tests/integration/test_<category>_workflow.py`
6. **Document** in `docs/API.md`

### CI/CD Gates (All MUST Pass)

| Gate | Command | Threshold |
|------|---------|-----------|
| Formatting | `black --check src/ tests/` | 0 violations |
| Linting | `ruff check src/` | 0 errors |
| Type Check | `mypy --strict src/` | 0 errors |
| Tests | `pytest --cov=excel_agent --cov-fail-under=90` | ≥90% |
| Integration | `pytest -m integration` | 100% pass |

---

## Critical Implementation Notes

### 1. Export Tool Parameter

**IMPORTANT**: Export tools use `--outfile` NOT `--output`:

```bash
# CORRECT
xls-export-pdf --input data.xlsx --outfile output.pdf

# WRONG (argparse conflict with common args)
xls-export-pdf --input data.xlsx --output output.pdf
```

### 2. Token Scopes

Valid scopes for `xls-approve-token`:
- `sheet:delete` - Remove entire sheet
- `sheet:rename` - Rename with ref update
- `range:delete` - Delete rows/columns/ranges
- `formula:convert` - Formulas → values (irreversible)
- `macro:remove` - Strip VBA (requires 2 tokens)
- `macro:inject` - Inject VBA project
- `structure:modify` - Batch structural changes

### 3. Impact Denial Pattern

When destructive operation breaks formulas:

```json
{
  "status": "denied",
  "exit_code": 1,
  "denial_reason": "Operation would break 7 formula references",
  "guidance": "Run xls-update-references --updates '[{\"old\": \"...\", \"new\": \"...\"}]' before retrying",
  "impact": {
    "broken_references": 7,
    "affected_sheets": ["Sheet1", "Sheet2"]
  }
}
```

### 4. Environment Variable

Set `EXCEL_AGENT_SECRET` for token operations:

```bash
export EXCEL_AGENT_SECRET="256-bit-hex-secret-key"
```

### 5. LibreOffice Requirement

- **Tier 2 Calculation**: Optional, provides full formula coverage
- **PDF Export**: Requires LibreOffice headless
- **Ubuntu/Debian**: `sudo apt-get install -y libreoffice-calc`

### 6. Tier 1 Calculation Workflow (CRITICAL)

**The `formulas` library operates on disk files, NOT in-memory workbooks.**

```python
# WRONG: This will calculate stale file
with ExcelAgent(path, mode="rw") as agent:
    agent.workbook["Sheet1"]["A1"] = 42  # In-memory only
    Tier1Calculator(path).calculate()  # Calculates old file!

# CORRECT: Save before calculating
with ExcelAgent(path, mode="rw") as agent:
    agent.workbook["Sheet1"]["A1"] = 42
    # ExcelAgent.__exit__ saves automatically

# Now calculate
Tier1Calculator(path).calculate()
```

**Workflow:** Save changes → Run Tier 1 → Reload workbook

---

## Common Issues & Solutions

### Issue: File Lock Not Released
**Cause**: Exception in context body before `__exit__`
**Solution**: `FileLock.__exit__` always releases lock (in `finally` block)

### Issue: `#REF!` Errors After Structural Change
**Cause**: Deleted cells were referenced by formulas
**Solution**: Run `xls-dependency-report` before destructive ops, fix refs with `xls-update-references`

### Issue: Token Validation Fails
**Cause**: Token expired, wrong scope, or file hash mismatch
**Solution**: Generate new token with correct scope and TTL

### Issue: Chunked Read Returns JSONL Not JSON
**Cause**: `--chunked` flag emits one JSON object per line
**Solution**: Parse as JSONL (one JSON per line), not single JSON

### Issue: SDK Returns ImpactDeniedError
**Cause**: Destructive operation would break formulas
**Solution** (using SDK):
```python
from excel_agent.sdk import ImpactDeniedError, AgentClient

client = AgentClient()
try:
    result = client.run("structure.xls_delete_sheet", ...)
except ImpactDeniedError as e:
    print(f"Guidance: {e.guidance}")
    print(f"Impact: {e.impact}")
    # Parse guidance, run remediation, retry
```

### Issue: Pre-commit Fails on detect-secrets
**Cause**: Existing secrets in baseline
**Solution**:
```bash
detect-secrets scan > .secrets.baseline
detect-secrets audit .secrets.baseline
git add .secrets.baseline
```

### Issue: Redis Backend Not Available
**Cause**: `redis` package not installed
**Solution**: `pip install excel-agent-tools[redis]`

---

## Lessons Learned (Phase 14)

### 1. Chunked I/O Output Format
**Discovery**: Chunked mode returns raw JSONL (row data), not JSON response envelope
**Fix**: Updated test to assert on `"values" in chunk` instead of `chunk.get("status")`

### 2. sigstore-python Package Name
**Discovery**: Package is `sigstore` in PyPI, not `sigstore-python`
**Fix**: Updated pyproject.toml to exclude from pip installs (recommend pipx)

### 3. Token Store Protocol Design
**Discovery**: In-memory set doesn't persist across processes
**Solution**: Added `TokenStore` Protocol with `InMemoryTokenStore` and `RedisTokenStore` implementations

### 4. SDK Error Handling
**Discovery**: AI agents need structured error information for recovery
**Solution**: `ImpactDeniedError` includes `guidance` and `impact` fields for programmatic remediation

### 5. Pre-commit Hook Order
**Discovery**: `detect-secrets` must run after other hooks to avoid false positives in generated files
**Solution**: Configured hook order with stages in `.pre-commit-config.yaml`

---

## Quick Reference

### Running Tests

```bash
# All tests
pytest

# With coverage
pytest --cov=excel_agent --cov-report=html

# Specific category
pytest tests/integration/test_clone_modify_workflow.py -v

# Exclude slow tests
pytest -m "not slow"
```

### Code Quality

```bash
# Format
black src/ tests/ --line-length 99

# Lint
ruff check src/ tests/

# Type check
mypy --strict src/
```

### Tool Invocation Example

```python
import json
import subprocess

result = subprocess.run(
    ["xls-read-range", "--input", "data.xlsx", "--range", "A1:C10"],
    capture_output=True,
    text=True
)

data = json.loads(result.stdout)
if data["status"] == "success":
    rows = data["data"]["values"]
```

### SDK Usage Example

```python
from excel_agent.sdk import AgentClient

client = AgentClient(secret_key="your-secret")

# Clone
clone = client.clone("template.xlsx", output_dir="./work")

# Read
data = client.read_range(clone, "A1:C10")

# Modify
client.write_range(clone, clone, "A1", [["New", "Data"]])

# Recalculate
client.recalculate(clone, clone)
```

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| openpyxl | >=3.1.5 | Core I/O |
| defusedxml | >=0.7.1 | XXE protection (mandatory) |
| formulas[excel] | >=1.3.4 | Tier 1 calculation |
| oletools | >=0.60.2 | Macro analysis |
| jsonschema | >=4.26.0 | Input validation |
| pandas | >=3.0.0 | Chunked I/O (internal) |
| redis | >=6.0.0 | Optional distributed state |

---

## Documentation Index

| File | Purpose |
|------|---------|
| `README.md` | Project overview, quick start |
| `Project_Architecture_Document.md` | Deep architecture (PAD) |
| `CLAUDE.md` | **THIS FILE** - Agent briefing |
| `docs/DESIGN.md` | Architecture blueprint |
| `docs/API.md` | Complete CLI reference (53 tools) |
| `docs/WORKFLOWS.md` | 5 production recipes with JSON |
| `docs/GOVERNANCE.md` | Token lifecycle & security |
| `docs/DEVELOPMENT.md` | Contributor guide |
| `CHANGELOG.md` | Version history |

---

## Status Summary

| Phase | Status | Deliverables |
|-------|--------|--------------|
| Phase 0 | ✅ Complete | Project scaffolding, CI/CD |
| Phase 1 | ✅ Complete | Core foundation (Agent, Lock, Serializer) |
| Phase 2 | ✅ Complete | Dependency engine, schemas |
| Phase 3 | ✅ Complete | Governance layer (Tokens, Audit) |
| Phase 4 | ✅ Complete | Governance + Read tools (13) |
| Phase 5 | ✅ Complete | Write tools (4) |
| Phase 6 | ✅ Complete | Structure tools (8) |
| Phase 7 | ✅ Complete | Cell operations (4) |
| Phase 8 | ✅ Complete | Formulas + Calculation (6) |
| Phase 9 | ✅ Complete | Macro safety tools (5) |
| Phase 10 | ✅ Complete | Object tools (5) |
| Phase 11 | ✅ Complete | Formatting tools (5) |
| Phase 12 | ✅ Complete | Export tools (3) |
| Phase 13 | ✅ Complete | E2E tests + Documentation |
| **Phase 14** | ✅ Complete | Hardening, Security, SDK, Pre-commit |

---

## For AI Coding Agents

### When Working on This Codebase

1. **NEVER** use `print()` in tools. Always return `dict` from `_run()`
2. **NEVER** catch `Exception` at tool level. Let `run_tool()` handle it
3. **ALWAYS** use `ExcelAgent` context manager for file I/O
4. **ALWAYS** validate inputs against schemas before core logic
5. **ALWAYS** handle all UI states: loading, error, empty, success
6. **ALWAYS** disable buttons during async operations
7. **ALWAYS** include `onError` handler with user feedback
8. **NEVER** commit secrets - pre-commit hooks will block
9. **ALWAYS** save before Tier 1 calculation (disk-based limitation)
10. **ALWAYS** use SDK for new integrations (simpler than raw subprocess)

### Code Style

- **Line length**: 99 characters
- **Type hints**: Strict mode enabled
- **Prefer**: `interface` for structures, `type` for unions
- **Never**: Use `any` - use `unknown` instead
- **Returns**: Early returns preferred over nested conditionals
- **Imports**: Sorted with `ruff check --select I`

### Testing Requirements

- Every tool must have unit test
- Every tool must have integration test (subprocess-based)
- Minimum coverage: 90%
- Test behavior, not implementation
- Use factory pattern for test data: `getMockX(overrides)`
- Pre-commit hooks must pass before committing

---

**Document Status**: Production-Ready
**Last Validated**: April 10, 2026
**Maintained by**: excel-agent-tools contributors
**License**: MIT
