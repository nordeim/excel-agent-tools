# excel-agent-tools

> **53 governance-first CLI tools for AI agents to safely read, write, and manipulate Excel workbooks — headless, JSON-native, zero Excel dependency.**

[![CI](https://github.com/user/excel-agent-tools/actions/workflows/ci.yml/badge.svg)](https://github.com/user/excel-agent-tools/actions/workflows/ci.yml)
[![Coverage](https://codecov.io/gh/user/excel-agent-tools/branch/main/graph/badge.svg)](https://codecov.io/gh/user/excel-agent-tools)
[![PyPI](https://img.shields.io/pypi/v/excel-agent-tools.svg)](https://pypi.org/project/excel-agent-tools/)
[![Python](https://img.shields.io/pypi/pyversions/excel-agent-tools.svg)](https://pypi.org/project/excel-agent-tools/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## Why excel-agent-tools?

AI agents need to manipulate spreadsheets safely. Existing tools either require Excel to
be running, lack governance controls, or can't handle formula dependencies. This project
provides **53 stateless CLI tools** that:

- **Never require Excel** — powered by openpyxl + formulas library for headless operation
- **Protect formula integrity** — dependency graph analysis before every destructive operation
- **Enforce governance** — HMAC-SHA256 scoped approval tokens for dangerous mutations
- **Speak JSON** — every tool reads JSON from stdin and writes JSON to stdout
- **Log everything** — pluggable audit trail for compliance and debugging

## Design Philosophy

| Principle | Implementation |
|:---|:---|
| **Governance-First** | Destructive operations require scoped HMAC-SHA256 tokens with TTL, nonce, and file-hash binding |
| **Formula Integrity** | Pre-flight dependency impact reports before any mutation that could break references |
| **Clone-Before-Edit** | Source files are never modified in-place; atomic copy to `/work/` first |
| **AI-Native** | JSON stdin/stdout, standardized exit codes (0-5), stateless CLI tools for agent chaining |
| **Headless** | No Excel, no COM, no GUI — runs on any server with Python ≥3.12 |

## Quick Start

### Installation

```bash
pip install excel-agent-tools
```

For full-fidelity recalculation (Tier 2), install LibreOffice headless:

```bash
# Ubuntu/Debian
sudo apt-get install -y libreoffice-calc

# macOS (Homebrew)
brew install --cask libreoffice

# Windows (Chocolatey)
choco install libreoffice-fresh
```

### 3-Step Workflow: Clone → Modify → Validate

```bash
# 1. Clone the source workbook to a safe working copy
xls-clone-workbook --input financials.xlsx --output-dir ./work/

# 2. Write data to the working copy
xls-write-range --input ./work/financials_20260408_abc123.xlsx \
  --output ./work/financials_20260408_abc123.xlsx \
  --range A1 --sheet Sheet1 \
  --data '[["Revenue", "Q1", "Q2"], ["Product A", 50000, 62000]]'

# 3. Validate the workbook (broken refs, circular deps, errors)
xls-validate-workbook --input ./work/financials_20260408_abc123.xlsx
```

### Governance: Token-Protected Deletion

```bash
# Generate a scoped approval token
xls-approve-token --scope sheet:delete \
  --file ./work/financials.xlsx --ttl 300

# Delete sheet with token (includes pre-flight dependency check)
xls-delete-sheet --input ./work/financials.xlsx \
  --output ./output/financials.xlsx \
  --name "OldSheet" --token "<token_from_above>"
```

## Tool Categories

| Category | Tools | Description |
|:---|:---|:---|
| **Governance** (6) | clone, validate, token, hash, lock, dependency | Safety infrastructure |
| **Read** (7) | range, sheets, names, tables, style, formula, metadata | Zero-mutation introspection |
| **Write** (4) | create, template, write-range, write-cell | Data insertion with type inference |
| **Structure** (8) | add/delete/rename/move sheet, insert/delete rows/cols | Token-gated mutations |
| **Cells** (4) | merge, unmerge, delete-range, update-refs | Cell-level operations |
| **Formulas** (6) | set, recalculate, detect-errors, convert, copy-down, define-name | Two-tier calc engine |
| **Objects** (5) | table, chart, image, comment, data-validation | Visualization & annotation |
| **Formatting** (5) | format-range, column-width, freeze, conditional, number-format | Style & layout |
| **Macros** (5) | has, inspect, validate-safety, remove, inject | oletools-backed VBA analysis |
| **Export** (3) | PDF, CSV, JSON | Interoperability |

## JSON Response Format

Every tool produces a standardized JSON envelope:

```json
{
  "status": "success",
  "exit_code": 0,
  "timestamp": "2026-04-08T14:30:22Z",
  "workbook_version": "sha256:abc123...",
  "data": {},
  "impact": {"cells_modified": 0, "formulas_updated": 0},
  "warnings": []
}
```

### Exit Codes

| Code | Meaning |
|:---|:---|
| `0` | Success |
| `1` | Validation error (bad input, schema violation, impact denial) |
| `2` | File not found |
| `3` | Lock contention (file locked by another process) |
| `4` | Permission denied (invalid/expired/wrong-scope token) |
| `5` | Internal error (bug, corrupt file, concurrent modification) |

## Documentation

- **[DESIGN.md](docs/DESIGN.md)** — Architecture blueprint and design decisions
- **[API.md](docs/API.md)** — CLI reference for all 53 tools
- **[WORKFLOWS.md](docs/WORKFLOWS.md)** — Common agent workflow recipes
- **[GOVERNANCE.md](docs/GOVERNANCE.md)** — Token scopes, audit trail, safety protocols
- **[DEVELOPMENT.md](docs/DEVELOPMENT.md)** — Contributing guide and code standards

## Requirements

- **Python ≥3.12**
- **openpyxl ≥3.1.5** — Excel I/O (headless)
- **defusedxml ≥0.7.1** — XML attack prevention (mandatory)
- **formulas[excel] ≥1.3.0** — Formula calculation engine (Tier 1)
- **oletools ≥0.60** — VBA macro analysis
- **LibreOffice Headless** — Full-fidelity recalculation (Tier 2, optional)

## Security

- **defusedxml is mandatory** — openpyxl does not guard against quadratic blowup or
  billion laughs XML attacks by default
- **HMAC-SHA256 tokens** use `hmac.compare_digest()` for constant-time comparison
- **Tokens are scoped, time-limited, single-use** with nonce tracking and file-hash binding
- **oletools** wrapped behind `MacroAnalyzer` Protocol for swappable backends
- **Audit trail** logs every destructive operation to `.excel_agent_audit.jsonl`

## License

[MIT](LICENSE)
