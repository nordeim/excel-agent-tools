# 📘 Project Architecture Document (PAD)
**Project:** `excel-agent-tools` v1.0.0
**Document Type:** Single Source-of-Truth Handbook
**Audience:** New Developers, AI Coding Agents, Technical Reviewers
**Last Validated:** April 10, 2026 | **Status:** Production-Ready | Phase 14 Complete

---

## 🎯 1. Executive Summary & Design Principles

`excel-agent-tools` is a production-grade, headless Python CLI suite of **53 stateless tools** enabling AI agents to safely read, mutate, calculate, and export Excel workbooks without Microsoft Excel or COM dependencies. The architecture enforces a **Governance-First, AI-Native** paradigm: destructive operations require cryptographic approval tokens, pre-flight dependency impact reports, and clone-before-edit workflows.

### Phase 14 Additions

**Agent Orchestration SDK** (`src/excel_agent/sdk/`)
- Simplified Pythonic wrapper around CLI tools
- Automatic retry logic with exponential backoff
- JSON response parsing with structured error classification
- Integration-friendly for LangChain, AutoGen, and custom frameworks

**Distributed State Management** (`src/excel_agent/governance/`)
- Pluggable storage backends via Protocols
- Redis implementation for multi-agent deployments
- Backward compatible (in-memory by default)

**Pre-commit Security Hardening** (`.pre-commit-config.yaml`)
- Secret detection with `detect-secrets`
- Code quality automation (black, ruff, mypy)
- Prevents accidental credential commits

### Core Design Principles

| Principle | Implementation |
|:---|:---|
| **Governance-First** | HMAC-SHA256 scoped tokens with TTL, nonce, file-hash binding. Denial-with-prescriptive-guidance pattern. |
| **Formula Integrity** | `DependencyTracker` builds AST graphs via `openpyxl.Tokenizer`. Blocks mutations breaking `#REF!` chains. |
| **AI-Native Contracts** | Strict JSON `stdout`, standardized exit codes (`0–5`), stateless CLI, zero TTY assumptions. |
| **Headless & Portable** | Zero Excel dependency. Linux/macOS/Windows ready. CI-matrix validated. |
| **Clone-Before-Edit** | Source files are immutable. Atomic copies to `/work/` are the mandatory mutation surface. |
| **Pluggable Extensibility** | `MacroAnalyzer` & `AuditBackend` Protocols isolate maintenance-heavy deps (`oletools`). |
| **Distributed-Ready** | `TokenStore` & `AuditBackend` Protocols enable Redis, PostgreSQL backends for multi-agent orchestration. |

---

## 🏗 2. System Architecture Overview

```mermaid
graph TD
    Agent[🤖 AI Agent / Orchestrator]
    SDK[🎁 AgentClient SDK]
    CLI[📦 CLI Tool Layer]
    Base[🔧 _tool_base.run_tool]
    Core[⚙️ Core Hub]
    Ext[🔌 External Engines]
    FS[💾 Filesystem]

    Agent -->|Python SDK| SDK
    Agent -->|JSON/subprocess| CLI
    SDK -->|subprocess + retry| CLI
    CLI --> Base
    Base --> Core
    Base -->|JSON Response| Agent

    subgraph "Phase 14: SDK Layer"
        SDK
    end

    subgraph Core Hub
        AgentCtx[ExcelAgent\nContext Manager]
        DepTracker[DependencyTracker\nFormula Graph]
        TokenMgr[ApprovalTokenManager\nHMAC-SHA256]
        Audit[AuditTrail\nPluggable Backend]
        Macro[MacroAnalyzer\nProtocol]
        Store[TokenStore\nInMemory/Redis]
    end

    Base --> AgentCtx
    Base --> DepTracker
    Base --> TokenMgr
    Base --> Macro
    AgentCtx --> Audit

    subgraph External Engines
        OpenPyXL[openpyxl 3.1.5]
        Defused[defusedxml 0.7.1]
        Formulas[formulas 1.3.4]
        LO[LibreOffice]
        Ole[oletools 0.60.2]
        Redis[Redis 6.4.0]
    end

    Core --> OpenPyXL
    Core --> Defused
    Core --> Formulas
    Core --> LO
    Core --> Ole
    Store -.->|Optional| Redis

    Core --> FS
    OpenPyXL --> FS
    LO --> FS
```

---

## 📁 3. File Hierarchy & Component Registry

```text
excel-agent-tools/
├── 📄 pyproject.toml                    # Build, metadata, 53 entry points
├── 📄 README.md                         # Project overview
├── 📄 Project_Architecture_Document.md  # Deep architecture (THIS FILE)
├── 📄 CLAUDE.md                         # Agent briefing
├── 📄 CHANGELOG.md                      # Phase 14 additions
├── 📄 .pre-commit-config.yaml          # NEW: Security & quality hooks
│
├── 📂 src/excel_agent/
│   ├── 📂 core/                          # Foundation layer
│   │   ├── 📜 agent.py                  # ExcelAgent: Lock → Load → Hash → Verify → Save
│   │   ├── 📜 locking.py                # Cross-platform FileLock
│   │   ├── 📜 serializers.py            # RangeSerializer
│   │   ├── 📜 dependency.py             # DependencyTracker + Tarjan SCC
│   │   ├── 📜 version_hash.py          # Geometry & File hashing
│   │   ├── 📜 formula_updater.py       # Reference shifting
│   │   ├── 📜 chunked_io.py            # Streaming for >100k rows
│   │   ├── 📜 type_coercion.py         # JSON → Python types
│   │   └── 📜 style_serializer.py      # Style serialization
│   │
│   ├── 📂 governance/                    # Security & Compliance
│   │   ├── 📜 token_manager.py          # ApprovalTokenManager
│   │   │                               #   + Phase 14: NonceStore Protocol support
│   │   ├── 📜 audit_trail.py           # AuditTrail backends
│   │   ├── 📜 stores.py                # NEW: TokenStore/AuditBackend Protocols
│   │   ├── 📂 backends/                 # NEW: Pluggable storage
│   │   │   ├── 📜 __init__.py
│   │   │   └── 📜 redis.py             # NEW: RedisTokenStore + RedisAuditBackend
│   │   └── 📂 schemas/                  # JSON Schema files
│   │
│   ├── 📂 calculation/                   # Two-tier calculation engine
│   │   ├── 📜 tier1_engine.py          # In-process formulas lib
│   │   └── 📜 tier2_libreoffice.py       # LibreOffice headless
│   │
│   ├── 📂 sdk/                          # NEW: Agent Orchestration SDK
│   │   ├── 📜 __init__.py              # Exports: AgentClient, exceptions
│   │   └── 📜 client.py                # AgentClient with retry/backoff
│   │
│   ├── 📂 utils/                        # Shared utilities
│   │   ├── 📜 exit_codes.py            # ExitCode IntEnum (0-5)
│   │   ├── 📜 json_io.py               # build_response(), ExcelAgentEncoder
│   │   ├── 📜 cli_helpers.py           # argparse patterns
│   │   └── 📜 exceptions.py              # ExcelAgentError hierarchy
│   │
│   └── 📂 tools/                        # 53 CLI entry points (10 categories)
│       ├── 📜 _tool_base.py            # Base runner for all tools
│       ├── 📂 governance/               # 6 tools
│       ├── 📂 read/                     # 7 tools
│       ├── 📂 write/                    # 4 tools
│       ├── 📂 structure/                # 8 tools ⚠️
│       ├── 📂 cells/                    # 4 tools
│       ├── 📂 formulas/                 # 6 tools
│       ├── 📂 objects/                  # 5 tools
│       ├── 📂 formatting/                 # 5 tools
│       ├── 📂 macros/                    # 5 tools ⚠️⚠️
│       └── 📂 export/                   # 3 tools
│
├── 📂 tests/                            # >90% coverage
│   ├── 📄 __init__.py
│   ├── 📄 conftest.py                   # Shared fixtures
│   ├── 📂 unit/                         # 20+ test modules
│   ├── 📂 integration/                  # 10+ test modules
│   └── 📂 property/                     # Hypothesis fuzzing
│
├── 📂 docs/
│   ├── 📄 DESIGN.md                     # Architecture blueprint
│   ├── 📄 API.md                        # CLI reference
│   ├── 📄 WORKFLOWS.md                  # Production recipes
│   ├── 📄 GOVERNANCE.md                 # Token lifecycle
│   └── 📄 DEVELOPMENT.md                # Contributor guide
│
└── 📂 scripts/
    └── 📄 install_libreoffice.sh        # CI setup
```

*(⚠️ = Requires approval token)*

---

## 🆕 Phase 14: New Components

### 3.1 Agent Orchestration SDK (`src/excel_agent/sdk/`)

**Purpose**: Simplified integration for AI frameworks

**Components:**
- `AgentClient`: Main SDK client with automatic retry logic
- `ImpactDeniedError`: Structured error with guidance and impact
- `TokenRequiredError`: Authentication error
- `ToolExecutionError`: General execution failure

**Usage:**
```python
from excel_agent.sdk import AgentClient, ImpactDeniedError

client = AgentClient(secret_key="your-secret")

try:
    # Automatic retry on lock contention (exit code 3)
    result = client.run(
        "structure.xls_delete_sheet",
        input="workbook.xlsx",
        name="OldSheet",
        token="hmac-token-here",
        max_retries=3
    )
except ImpactDeniedError as e:
    # Structured error for programmatic recovery
    print(f"Denied: {e.guidance}")
    print(f"Impact: {e.impact}")
    # Run remediation per guidance, then retry
```

### 3.2 Distributed State Protocols (`src/excel_agent/governance/`)

**Purpose**: Multi-agent orchestration support

**New Files:**
- `stores.py`: `TokenStore` and `AuditBackend` Protocols
- `backends/redis.py`: Redis implementations

**Protocol Design:**
```python
from typing import Protocol

class TokenStore(Protocol):
    def add(self, nonce: str) -> None: ...
    def __contains__(self, nonce: str) -> bool: ...
    def clear(self) -> None: ...

class AuditBackend(Protocol):
    def log_event(self, event: dict) -> None: ...
    def query_events(self, **filters) -> list[dict]: ...
```

**Implementations:**
- `InMemoryTokenStore`: Default single-process
- `RedisTokenStore`: Distributed with TTL
- `JsonlAuditBackend`: Default file-based
- `RedisAuditBackend`: Redis Streams

**Usage:**
```python
from excel_agent.governance.backends.redis import RedisTokenStore
from excel_agent.governance.token_manager import ApprovalTokenManager

redis_store = RedisTokenStore("redis://localhost:6379")
manager = ApprovalTokenManager(
    secret="secret-key",
    nonce_store=redis_store  # Optional: uses in-memory by default
)
```

### 3.3 Pre-commit Configuration (`.pre-commit-config.yaml`)

**Purpose**: Automated security and quality gates

**Hooks:**
1. `trailing-whitespace`, `end-of-file-fixer`
2. `detect-private-key`
3. `detect-secrets` (Yelp)
4. `black` (formatting, line-length 99)
5. `ruff` (linting)
6. `mypy` (type checking)
7. `markdownlint`

**Usage:**
```bash
pre-commit install        # One-time setup
pre-commit run --all-files  # Manual run
git commit -m "..."       # Automatic hooks
```

---

## 🗃 4. Data Models & Schema Registry

### 4.1 Input Validation Schemas

*(Unchanged from Phase 0-13)*

### 4.2 Phase 14: New Data Structures

```mermaid
classDiagram
    class AgentClientConfig {
        +string secret_key
        +string python_executable
        +int timeout
        +float base_retry_delay
    }

    class TokenStore {
        +add(nonce: str)
        +__contains__(nonce: str) bool
        +clear()
    }

    class InMemoryTokenStore {
        -set _nonces
    }

    class RedisTokenStore {
        -Redis _redis
        -string _key
        -int _ttl_seconds
    }

    class AuditBackend {
        +log_event(event: dict)
        +query_events(**filters) list
    }

    class JsonlAuditBackend {
        -Path _log_file
    }

    class RedisAuditBackend {
        -Redis _redis
        -string _stream_key
    }

    TokenStore <|-- InMemoryTokenStore
    TokenStore <|-- RedisTokenStore
    AuditBackend <|-- JsonlAuditBackend
    AuditBackend <|-- RedisAuditBackend
```

---

## 🔄 5. Core Application Flowcharts

### 5.1 SDK Execution Flow (Phase 14)

```mermaid
sequenceDiagram
    participant User as AI Agent
    participant SDK as AgentClient
    participant Tool as CLI Tool
    participant Base as _tool_base
    participant Core as Core Hub

    User->>SDK: client.run(tool, **kwargs)
    SDK->>SDK: Build subprocess command
    loop Retry Loop (max_retries)
        SDK->>Tool: subprocess.run()
        Tool->>Base: run_tool(_run)
        Base->>Core: Execute logic
        Core-->>Base: Result or Exception
        Base-->>Tool: JSON response
        Tool-->>SDK: stdout, returncode

        alt Success (exit 0)
            SDK-->>User: Parsed JSON
        else Lock Contention (exit 3)
            SDK->>SDK: Exponential backoff
            SDK->>Tool: Retry
        else Impact Denied
            SDK->>SDK: Raise ImpactDeniedError
            SDK-->>User: Structured exception
        else Other Error
            SDK-->>User: Raise ToolExecutionError
        end
    end
```

### 5.2 Token Validation with Distributed Store

```mermaid
flowchart TD
    Start[Token Validation] --> Parse[Parse Token]
    Parse --> CheckScope{Check Scope}
    CheckScope -->|Mismatch| Error1[Exit 4: Scope Mismatch]
    CheckScope -->|OK| CheckTTL{Check TTL}
    CheckTTL -->|Expired| Error2[Exit 4: Token Expired]
    CheckTTL -->|OK| CheckHash{Check File Hash}
    CheckHash -->|Mismatch| Error3[Exit 4: Hash Mismatch]
    CheckHash -->|OK| CheckNonce{Check Nonce}

    CheckNonce -->|External Store| Redis{Redis Store}
    CheckNonce -->|Local| Memory[InMemory Set]

    Redis -->|Exists| Error4[Exit 4: Replay]
    Memory -->|Exists| Error4
    Redis -->|New| MarkRedis[Mark Used in Redis]
    Memory -->|New| MarkMemory[Mark Used in Memory]

    MarkRedis --> VerifySig[HMAC Verify]
    MarkMemory --> VerifySig
    VerifySig -->|Invalid| Error5[Exit 4: Bad Signature]
    VerifySig -->|Valid| Success[Token Validated]
```

---

## 🔐 6. Governance & Security Architecture (Updated)

| Component | Security Mechanism | Validation |
|:---|:---|:---|
| **Tokens** | `hmac.compare_digest()` (RFC 2104), `secrets.token_hex(16)` nonce, 256-bit secret | Prevents timing attacks & replay |
| **Nonce Store** | Pluggable `TokenStore` Protocol: In-memory (default) or Redis (distributed) | Supports multi-agent orchestration |
| **Scope Binding** | `scope\|file_hash\|nonce\|issued_at\|ttl` canonical string | Mathematically impossible cross-file reuse |
| **Audit Trail** | Pluggable `AuditBackend`: JSONL (default) or Redis Streams | Atomic, tamper-evident |
| **Macro Safety** | `MacroAnalyzer` Protocol, `scan_risk()` pre-condition | `oletools` isolated, source never logged |
| **XML Defense** | Mandatory `defusedxml` import | Blocks XXE & Billion Laughs |
| **Pre-commit** | `detect-secrets`, `detect-private-key` | Prevents accidental credential commits |
| **File Integrity** | Geometry hash (formulas/structure) vs File hash (bytes) | Detects silent vs structural mutations |

---

## 🛠 7. Tool Execution Pipeline & Developer Hooks

### SDK Error Handling

```python
from excel_agent.sdk import AgentClient, ImpactDeniedError

client = AgentClient(secret_key="secret")

try:
    result = client.run("structure.xls_delete_sheet", ...)
except ImpactDeniedError as e:
    # Phase 14: Structured error with guidance
    guidance = e.guidance  # Prescriptive remediation steps
    impact = e.impact      # Affected cells, formulas, sheets
    exit_code = e.exit_code

    # Parse guidance, run remediation, retry
    print(f"Guidance: {guidance}")
    print(f"Impact: {impact}")
```

---

## 📚 8. Lessons Learned (Phase 14)

### 8.1 Chunked I/O Test Fix

**Issue:** Test expected `chunk.get("status") == "success"` but chunked mode returns raw JSONL.

**Root Cause:** Chunked mode writes directly to stdout without JSON envelope.

**Fix:** Assert on `"values" in chunk` instead of status field.

**Prevention:** Add test comment explaining JSONL format.

### 8.2 Package Name Discrepancy

**Issue:** `sigstore-python` doesn't exist in PyPI (correct name is `sigstore`).

**Root Cause:** Documentation inconsistency.

**Fix:** Updated pyproject.toml to use `sigstore` (optional), recommend pipx install.

### 8.3 Distributed State Design

**Issue:** In-memory nonce set doesn't persist across processes.

**Root Cause:** Single-process assumption in original design.

**Solution:** Added `TokenStore` Protocol with duck-typing support.

**Implementation:**
```python
# Backward compatible - uses in-memory by default
manager = ApprovalTokenManager(secret="key")

# Distributed - uses Redis
manager = ApprovalTokenManager(
    secret="key",
    nonce_store=RedisTokenStore("redis://...")
)
```

### 8.4 Pre-commit Hook Ordering

**Issue:** `detect-secrets` flagged generated files in `.venv/`.

**Root Cause:** Hooks ran on all files including virtual environment.

**Fix:** Added `exclude` patterns to `.pre-commit-config.yaml`.

### 8.5 Tier 1 Calculation Clarification

**Issue:** Developers confused about when Tier 1 calculates.

**Root Cause:** `formulas` library reads from disk, not openpyxl memory.

**Documentation Update:** Added explicit warning in CLAUDE.md:

```python
# WRONG
with ExcelAgent(path) as agent:
    agent.workbook["A1"] = 42
    Tier1Calculator(path).calculate()  # Calculates OLD file

# CORRECT
with ExcelAgent(path) as agent:
    agent.workbook["A1"] = 42
# Agent exits, saves automatically
Tier1Calculator(path).calculate()  # Calculates UPDATED file
```

---

## ✅ 9. Validation & Alignment Matrix

| Master Plan Requirement | PAD Implementation | Status |
|:---|:---|:---|
| 53 CLI Tools, JSON I/O, Exit 0-5 | `_tool_base.run_tool()`, `build_response()`, `ExitCode` enum | ✅ Aligned |
| Governance-First Tokens | `ApprovalTokenManager`, HMAC, TTL, nonce, `compare_digest` | ✅ Aligned |
| Formula Integrity Pre-flight | `DependencyTracker`, Tarjan's SCC, `ImpactDeniedError` + guidance | ✅ Aligned |
| Clone-Before-Edit | `xls_clone_workbook.py`, immutable source policy | ✅ Aligned |
| Two-Tier Calculation | `formulas` 1.3.4 (Tier 1) → LibreOffice (Tier 2) | ✅ Aligned |
| Macro Safety Protocol | `MacroAnalyzer` Protocol, `oletools` backend, `scan_risk()` | ✅ Aligned |
| Headless & Server-Ready | Zero COM, `defusedxml` mandatory, CI matrix | ✅ Aligned |
| Audit Trail | Pluggable `AuditBackend`, JSONL append | ✅ Aligned |
| **Phase 14: Agent SDK** | `AgentClient` with retry/backoff | ✅ **NEW** |
| **Phase 14: Distributed State** | `TokenStore` Protocol, `RedisTokenStore` | ✅ **NEW** |
| **Phase 14: Pre-commit** | Security hooks, quality gates | ✅ **NEW** |
| Python ≥3.12 Floor | `pyproject.toml` `requires-python=">=3.12"` | ✅ Aligned |

---

## 📎 Appendix: Quick Reference Tables

### Exit Code Semantics

| Code | Meaning | Agent Recovery Action |
|:---:|:---|:---|
| `0` | Success | Parse `data`, chain to next tool |
| `1` | Validation / Impact Denial | Fix JSON input or run remediation tool |
| `2` | File Not Found | Verify path, download missing file |
| `3` | Lock Contention | Exponential backoff retry (`0.5s → 1s → 2s`) |
| `4` | Permission Denied | Request new token with correct scope |
| `5` | Internal Error | Halt workflow, alert operator, attach traceback |

### Token Scopes

`sheet:delete`, `sheet:rename`, `range:delete`, `formula:convert`, `macro:remove`, `macro:inject`, `structure:modify`

### Key Dependencies (Updated)

| Package | Version | Purpose | Phase 14 Note |
|:---|:---|:---|:---|
| `openpyxl` | `3.1.5` | Core .xlsx/.xlsm I/O | - |
| `defusedxml` | `0.7.1` | XML attack prevention | Mandatory |
| `formulas[excel]` | `1.3.4` | Tier 1 calculation | Disk-based limitation documented |
| `oletools` | `0.60.2` | VBA/XLM risk scanning | Protocol-wrapped |
| `jsonschema` | `>=4.26.0` | Input validation | Updated from 4.23 |
| `pandas` | `>=3.0.0` | Chunked I/O | Updated from 2.x |
| `redis` | `>=6.0.0` | Distributed state | NEW: Optional extra |
| `detect-secrets` | `>=1.5.0` | Pre-commit scanning | NEW: Optional extra |

---

## 🚀 Phase 14 Status

| Task | Status | File |
|:---|:---|:---|
| Agent SDK Implementation | ✅ Complete | `src/excel_agent/sdk/` |
| TokenStore Protocol | ✅ Complete | `src/excel_agent/governance/stores.py` |
| RedisTokenStore | ✅ Complete | `src/excel_agent/governance/backends/redis.py` |
| Pre-commit Configuration | ✅ Complete | `.pre-commit-config.yaml` |
| Dependency Version Update | ✅ Complete | `pyproject.toml`, `requirements*.txt` |
| Chunked I/O Test Fix | ✅ Complete | `tests/integration/test_clone_modify_workflow.py` |
| Documentation Update | ✅ Complete | `CLAUDE.md`, `CHANGELOG.md`, `README.md` |

---

**Document End.**
This PAD serves as the authoritative blueprint. Phase 14 additions are backward compatible and production-ready.
