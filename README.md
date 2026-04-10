# excel-agent-tools

> **53 governance-first, AI-native CLI tools for safe, headless Excel manipulation.**

[![CI](https://github.com/<ORGANIZATION>/<REPOSITORY>/actions/workflows/ci.yml/badge.svg)](https://github.com/<ORGANIZATION>/<REPOSITORY>/actions/workflows/ci.yml)
[![Coverage](https://codecov.io/gh/<ORGANIZATION>/<REPOSITORY>/branch/main/graph/badge.svg)](https://codecov.io/gh/<ORGANIZATION>/<REPOSITORY>)
[![PyPI](https://img.shields.io/pypi/v/excel-agent-tools.svg)](https://pypi.org/project/excel-agent-tools/)
[![Python](https://img.shields.io/pypi/pyversions/excel-agent-tools.svg)](https://pypi.org/project/excel-agent-tools/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Downloads](https://static.pepy.tech/badge/excel-agent-tools)](https://pepy.tech/project/excel-agent-tools)

---

## 🤔 Why `excel-agent-tools`?

AI agents need to manipulate spreadsheets safely. Existing automation approaches suffer from critical gaps:
- 🚫 **Require Excel/COM:** Break in headless, serverless, or Linux environments.
- 🚫 **Silent Formula Breakage:** Structural edits (`delete_row`, `rename_sheet`) silently corrupt `#REF!` chains.
- 🚫 **No Governance:** Destructive operations lack approval gates, audit trails, or rollback safety.
- 🚫 **Poor Agent UX:** Inconsistent exit codes, unstructured stdout, and no prescriptive error recovery.

`excel-agent-tools` solves this with **53 stateless, JSON-native CLI tools** designed specifically for autonomous orchestration:
- ✅ **Headless & Server-Ready:** Zero Microsoft Excel dependency. Powered by `openpyxl`, `formulas`, and optional LibreOffice.
- ✅ **Formula Integrity Preservation:** Pre-flight dependency graphs block mutations that would break references.
- ✅ **Governance-First:** HMAC-SHA256 scoped tokens, TTL, nonce tracking, and immutable audit trails.
- ✅ **AI-Native Contracts:** Strict JSON envelopes, standardized exit codes (`0–5`), and denial-with-prescriptive-guidance.
- ✅ **Agent SDK:** Pythonic wrapper with automatic retry logic for LangChain, AutoGen, and custom frameworks.

---

## 🚀 Quick Start

### 1. Installation

```bash
# Base package
pip install excel-agent-tools

# With optional dependencies
pip install excel-agent-tools[redis]     # Distributed state management
pip install excel-agent-tools[security]  # Security hardening tools

# For full-fidelity recalculation (Tier 2), install LibreOffice headless:
# Ubuntu/Debian: sudo apt-get install -y libreoffice-calc
# macOS: brew install --cask libreoffice
# Windows: choco install libreoffice-fresh
```

### 2. 3-Step Workflow: Clone → Modify → Validate

```bash
# 1. Clone source to a safe working copy (never mutate originals)
xls-clone-workbook --input financials.xlsx --output-dir ./work/

# 2. Write data to the working copy
xls-write-range --input ./work/financials_*.xlsx \
  --output ./work/financials_*.xlsx \
  --range A1 --sheet Q1 --data '["Revenue","Cost"],[50000,32000]]'

# 3. Validate integrity before proceeding
xls-validate-workbook --input ./work/financials_*.xlsx
```

### 3. Governance: Token-Protected Deletion

```bash
# Generate a scoped approval token (expires in 5 min)
TOKEN=$(xls-approve-token --scope sheet:delete --file ./work/financials.xlsx --ttl 300 | jq -r '.data.token')

# Delete sheet with pre-flight dependency check
xls-delete-sheet --input ./work/financials.xlsx \
  --output ./output/clean.xlsx \
  --name "OldSheet" --token "$TOKEN"
```

### 4. Python SDK (NEW in Phase 14)

```python
from excel_agent.sdk import AgentClient, ImpactDeniedError

# Initialize client with retry logic
client = AgentClient(secret_key="your-secret")

# Clone and modify
clone_path = client.clone("financials.xlsx", output_dir="./work")
data = client.read_range(clone_path, "A1:C10")

# Safe structural edit with automatic retry
try:
    token = client.generate_token("sheet:delete", clone_path)
    client.run("structure.xls_delete_sheet",
              input=clone_path, name="OldSheet", token=token)
except ImpactDeniedError as e:
    print(f"Guidance: {e.guidance}")
    print(f"Impact: {e.impact}")
    # Run remediation per guidance, then retry
```

---

## 🌟 Key Features

| Feature | Description |
|:---|:---|
| 🛡️ **Governance-First** | Destructive ops require scoped HMAC-SHA256 tokens with TTL, nonce, and file-hash binding. |
| 🔗 **Formula Integrity** | `DependencyTracker` builds AST graphs to block mutations that break `#REF!` chains. |
| 🤖 **AI-Native UX** | JSON stdin/stdout, standardized exit codes (`0-5`), stateless chaining, prescriptive guidance. |
| ☁️ **Headless Operation** | Runs on any server. No Excel, no COM, no GUI dependencies. |
| 🔒 **File Safety** | OS-level cross-platform locking, clone-before-edit enforcement, geometry hash verification. |
| 📊 **Two-Tier Calculation** | Tier 1 (`formulas` lib, ~50ms for 10k formulas) → Tier 2 (LibreOffice headless fallback). |
| 🦠 **Macro Safety** | `oletools`-backed risk scanning, read-only isolation, pre-scan before any `.bin` injection. |
| 📝 **Pluggable Audit** | Append-only JSONL trails by default. Swap to SIEM/webhooks via `AuditBackend` Protocol. |
| 🎁 **Agent SDK** | Pythonic wrapper with retry logic, JSON parsing, token management for AI frameworks. |
| 🌐 **Distributed-Ready** | Redis backends for `TokenStore` and `AuditBackend` in multi-agent deployments. |

---

## 🏗 Architecture Overview

### 📁 File Hierarchy & Key Components

```text
excel-agent-tools/
├── 📄 pyproject.toml              # 53 entry points, deps, tool configs
├── 📂 src/excel_agent/
│   ├── 📂 core/                    # Foundation: Agent, Lock, Serializer, Dependency
│   ├── 📂 governance/              # Security: Tokens, Audit, Schemas
│   │   ├── 📜 token_manager.py     # HMAC-SHA256 scoped tokens
│   │   ├── 📜 stores.py            # NEW: TokenStore Protocol
│   │   └── 📂 backends/            # NEW: Pluggable backends (Redis)
│   ├── 📂 calculation/             # Two-tier engine (formulas, LibreOffice)
│   ├── 📂 sdk/                     # NEW: Agent Orchestration SDK
│   │   ├── 📜 __init__.py
│   │   └── 📜 client.py            # AgentClient with retry/backoff
│   ├── 📂 utils/                   # CLI helpers, JSON I/O, exceptions
│   └── 📂 tools/                   # 53 CLI entry points (10 categories)
├── 📂 tests/                       # >90% coverage
├── 📂 docs/                        # DESIGN, API, WORKFLOWS, GOVERNANCE
└── 📄 .pre-commit-config.yaml    # NEW: Security & quality hooks
```

### 🔄 User & Application Interaction

```mermaid
sequenceDiagram
    participant User as AI Agent / Orchestrator
    participant SDK as AgentClient (NEW)
    participant CLI as CLI Tool
    participant Base as _tool_base.py
    participant Core as Core Hub

    alt Using SDK (NEW)
        User->>SDK: client.run(tool, **kwargs)
        SDK->>SDK: Retry logic, JSON parsing
        SDK->>CLI: subprocess.run()
    else Direct CLI
        User->>CLI: subprocess.run()
    end
    
    CLI->>Base: run_tool() wrapper
    Base->>Core: Execute with validation
    Core-->>Base: Result
    Base-->>CLI: JSON response
    CLI-->>User: Exit code + JSON
```

---

## 📦 Tool Catalog (53 Tools)

| Category | Count | Description | Governance |
|:---|:---:|:---|:---|
| 🔐 **Governance** | 6 | Clone, validate, approve token, version hash, lock status, dependency report | Token gen, Audit |
| 📖 **Read** | 7 | Range data, sheet names, defined names, tables, styles, formulas, metadata | Read-only |
| ✍️ **Write** | 4 | Create new, template substitution, write range, write cell | Type inference |
| 🏗 **Structure** | 8 | Add/delete/rename/move sheet, insert/delete rows & cols | ⚠️ Token + Impact Check |
| 📐 **Cells** | 4 | Merge, unmerge, delete range, batch update references | ⚠️ Token (delete) |
| 🧮 **Formulas** | 6 | Set formula, recalculate, detect errors, convert to values, copy down, define name | ⚠️ Token (convert) |
| 📊 **Objects** | 5 | Tables, charts, images, comments, data validation | Additive |
| 🎨 **Formatting** | 5 | Range styles, column width, freeze panes, conditional formatting, number formats | Additive |
| 🦠 **Macros** | 5 | Detect, inspect, validate safety, remove, inject VBA | ⚠️⚠️ Token + Pre-scan |
| 📤 **Export** | 3 | PDF (via LibreOffice), CSV, JSON | Read-only |

⚠️ = Requires `--token` and `--acknowledge-impact` if dependencies exist.

---

## 🔌 Standardized Interfaces

### JSON Response Envelope

```json
{
  "status": "success",
  "exit_code": 0,
  "timestamp": "2026-04-10T14:30:22+00:00",
  "workbook_version": "sha256:a1b2c3...",
  "data": { "sheets": ["Q1", "Q2"], "count": 2 },
  "impact": { "cells_modified": 4, "formulas_updated": 2 },
  "warnings": [],
  "guidance": null
}
```

### Exit Code Semantics

| Code | Meaning | Agent Action |
|:---:|:---|:---|
| `0` | Success | Parse `data`, proceed |
| `1` | Validation / Impact Denial | Fix input, run remediation, or retry with `--acknowledge-impact` |
| `2` | File Not Found | Verify path |
| `3` | Lock Contention | Exponential backoff (`0.5s → 1s → 2s`) |
| `4` | Permission Denied | Request new token |
| `5` | Internal Error | Alert operator |

### SDK Usage (NEW)

```python
from excel_agent.sdk import AgentClient

client = AgentClient(secret_key="secret")

# Clone workbook
clone = client.clone("data.xlsx", output_dir="./work")

# Read range
data = client.read_range(clone, "A1:C10")

# Write with automatic retry on lock contention
client.write_range(
    clone, clone, "A1",
    [["New", "Data"]],
    max_retries=3
)

# Recalculate
result = client.recalculate(clone, clone)
print(f"Calculated {result['data']['calculated_count']} formulas")
```

---

## 🛡 Governance & Safety Protocol

1. **Token Lifecycle:** Generate → Validate (scope, file-hash, TTL, nonce) → Use → Revoke.
2. **Clone-Before-Edit:** Source workbooks are immutable. All mutations happen on timestamped clones.
3. **Impact Denial & Guidance:** If a structural change breaks formulas, the tool exits with code `1` and provides exact remediation steps.
4. **Audit Trail:** Every operation appends to `.excel_agent_audit.jsonl`. Macro source code **never** enters the log.
5. **Pre-commit Security:** (NEW) `detect-secrets` and `detect-private-key` hooks prevent accidental credential commits.

---

## ☁️ Deployment

### 📦 PyPI Package

```bash
pip install excel-agent-tools

# Optional extras
pip install excel-agent-tools[redis]      # Distributed state
pip install excel-agent-tools[security]   # Security scanning
pip install excel-agent-tools[dev]        # Development tools
```

### 🐳 Docker Container

```dockerfile
FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc && \
    pip install excel-agent-tools[redis]
WORKDIR /data
```

### 🖥 Server / CI Integration

- **GitHub Actions:** Pre-installed LibreOffice matrix. Coverage gate ≥90%.
- **Pre-commit:** Install with `pre-commit install` for automated security/quality checks.
- **Production Checklist:**
  - [ ] Set `EXCEL_AGENT_SECRET` in vault/env.
  - [ ] Configure audit backend (JSONL or Redis).
  - [ ] Enable file locking (`fcntl`/`msvcrt`).
  - [ ] Set up log rotation for audit trails.
  - [ ] Install pre-commit hooks for contributors.

---

## 📋 Requirements & Installation

| Component | Requirement | Notes |
|:---|:---|:---|
| **Python** | `≥3.12` | Modern stdlib, strict typing |
| **Core I/O** | `openpyxl >=3.1.5` | Headless XML parsing |
| **XML Security** | `defusedxml >=0.7.1` | **Mandatory.** Prevents XXE attacks |
| **Calc Tier 1** | `formulas[excel] >=1.3.4` | In-process AST evaluation |
| **Macro Safety** | `oletools >=0.60.2` | VBA/XLM risk scanning |
| **Calc Tier 2** | LibreOffice Headless | Optional. Required for PDF export |
| **Distributed** | `redis >=6.0.0` | Optional. Multi-agent state management |
| **Security** | `detect-secrets >=1.5.0` | Optional. Pre-commit secret scanning |

---

## 🔒 Security Notices

- 🛡️ **XML Defense:** `defusedxml` is mandatory to prevent XXE attacks.
- 🔐 **Token Cryptography:** All tokens use `hmac.compare_digest()` (RFC 2104) to prevent timing side-channels.
- 🦠 **Macro Isolation:** `oletools` is isolated behind the `MacroAnalyzer` Protocol. Source code is **never** logged.
- 📜 **Audit Privacy:** Audit trails record metadata only. Sensitive payloads are excluded.
- 🔍 **Pre-commit Security:** (NEW) `detect-secrets` hooks prevent accidental credential commits.

---

## 📚 Documentation Index

| Document | Description |
|:---|:---|
| [`CLAUDE.md`](CLAUDE.md) | AI Coding Agent Briefing - Complete reference for AI agents |
| [`DESIGN.md`](docs/DESIGN.md) | Architecture blueprint and trade-off analysis |
| [`API.md`](docs/API.md) | Complete CLI reference for all 53 tools |
| [`WORKFLOWS.md`](docs/WORKFLOWS.md) | 5 production-ready agent recipes |
| [`GOVERNANCE.md`](docs/GOVERNANCE.md) | Token lifecycle, audit schema, safety protocols |
| [`DEVELOPMENT.md`](docs/DEVELOPMENT.md) | Contributor guide, CI setup, adding tools |
| [`CHANGELOG.md`](CHANGELOG.md) | Version history and Phase 14 additions |

---

## 🆕 What's New (Phase 14 - April 2026)

### Agent Orchestration SDK
- `AgentClient` class with automatic retry logic
- `ImpactDeniedError` with structured guidance and impact
- Convenience methods: `clone()`, `read_range()`, `write_range()`, `recalculate()`

### Distributed State Management
- `TokenStore` Protocol for pluggable nonce storage
- `RedisTokenStore` implementation for multi-agent deployments
- `AuditBackend` Protocol with Redis Streams support

### Pre-commit Security
- `detect-secrets` for preventing credential leaks
- `black`, `ruff`, `mypy` for code quality automation

### Dependency Updates
- Updated `pandas>=3.0.0` (from 2.x)
- Updated `jsonschema>=4.26.0`
- Pinned all versions in `requirements.txt`

---

## 🤝 Contributing & License

We welcome contributions that maintain the project's **governance-first, research-validated** standards.
- Read [`DEVELOPMENT.md`](docs/DEVELOPMENT.md) for setup and testing guidelines.
- All PRs require passing CI matrix (`black`, `ruff`, `mypy --strict`, `pytest --cov=90`).
- Pre-commit hooks must pass: `pre-commit run --all-files`

**License:** [MIT](LICENSE)  
© 2026 `excel-agent-tools` contributors.
