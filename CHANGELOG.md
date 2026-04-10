# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-04-10

### Added

#### Phase 14: Hardening, Security & Release Preparation

- **Agent Orchestration SDK** (`src/excel_agent/sdk/`)
  - `AgentClient` class for simplified AI agent integration
  - Automatic retry logic with exponential backoff
  - JSON response parsing with error classification
  - Token generation helper
  - Convenience methods: `clone()`, `read_range()`, `write_range()`, `recalculate()`
  - Custom exceptions: `ImpactDeniedError`, `TokenRequiredError`, `ToolExecutionError`

- **Pre-commit Configuration** (`.pre-commit-config.yaml`)
  - Secret detection with `detect-secrets`
  - Code formatting with `black`
  - Linting with `ruff`
  - Type checking with `mypy`
  - Markdown linting

- **Distributed State Protocols** (`src/excel_agent/governance/`)
  - `TokenStore` Protocol for pluggable nonce storage
  - `AuditBackend` Protocol for pluggable audit logging
  - `InMemoryTokenStore` (default single-process)
  - `RedisTokenStore` for distributed deployments
  - `RedisAuditBackend` for Redis Streams audit logging

- **Token Manager Enhancement**
  - Support for external nonce stores via constructor injection
  - Backward compatible (uses in-memory set by default)

- **Optional Dependencies**
  - `redis` extra for distributed state management
  - `security` extra for hardening tools (cyclonedx, detect-secrets, safety)

### Fixed

- **Chunked I/O Test** (`tests/integration/test_clone_modify_workflow.py`)
  - Fixed test expectation to match chunked JSONL output format

### Changed

- **Dependency Updates**
  - Updated `pandas>=2.1.0,<3.0.0` to `pandas>=3.0.0`
  - Updated `jsonschema>=4.23.0,<5.0.0` to `jsonschema>=4.26.0`
  - Pinned all versions in `requirements.txt` and `requirements-dev.txt`

## [1.0.0-rc1] - 2026-04-09

### Added

- **Initial Release**
  - 53 CLI tools across 10 categories
  - Core foundation: `ExcelAgent`, `FileLock`, `RangeSerializer`, `VersionHash`
  - Dependency engine: `DependencyTracker` with Tarjan's SCC
  - Governance layer: `ApprovalTokenManager`, `AuditTrail`
  - Two-tier calculation: Tier 1 (formulas library), Tier 2 (LibreOffice)
  - Macro safety: `MacroAnalyzer` Protocol with `oletools` backend
  - Complete documentation: DESIGN.md, API.md, WORKFLOWS.md, GOVERNANCE.md, DEVELOPMENT.md

## Categories

### Governance (6 tools)
- `xls-clone-workbook`
- `xls-validate-workbook`
- `xls-approve-token`
- `xls-version-hash`
- `xls-lock-status`
- `xls-dependency-report`

### Read (7 tools)
- `xls-read-range`
- `xls-get-sheet-names`
- `xls-get-workbook-metadata`
- `xls-get-defined-names`
- `xls-get-table-info`
- `xls-get-cell-style`
- `xls-get-formula`

### Write (4 tools)
- `xls-create-new`
- `xls-create-from-template`
- `xls-write-range`
- `xls-write-cell`

### Structure (8 tools)
- `xls-add-sheet`
- `xls-delete-sheet` ⚠️
- `xls-rename-sheet` ⚠️
- `xls-insert-rows`
- `xls-delete-rows` ⚠️
- `xls-insert-columns`
- `xls-delete-columns` ⚠️
- `xls-move-sheet`

### Cells (4 tools)
- `xls-merge-cells`
- `xls-unmerge-cells`
- `xls-delete-range` ⚠️
- `xls-update-references`

### Formulas (6 tools)
- `xls-set-formula`
- `xls-recalculate`
- `xls-detect-errors`
- `xls-convert-to-values` ⚠️
- `xls-copy-formula-down`
- `xls-define-name`

### Objects (5 tools)
- `xls-add-table`
- `xls-add-chart`
- `xls-add-image`
- `xls-add-comment`
- `xls-set-data-validation`

### Formatting (5 tools)
- `xls-format-range`
- `xls-set-column-width`
- `xls-freeze-panes`
- `xls-apply-conditional-formatting`
- `xls-set-number-format`

### Macros (5 tools)
- `xls-has-macros`
- `xls-inspect-macros`
- `xls-validate-macro-safety`
- `xls-remove-macros` ⚠️⚠️
- `xls-inject-vba-project` ⚠️

### Export (3 tools)
- `xls-export-pdf`
- `xls-export-csv`
- `xls-export-json`

⚠️ = Token required | ⚠️⚠️ = Double-token required
