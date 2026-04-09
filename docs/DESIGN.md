# Architecture Blueprint: excel-agent-tools

**Version:** 1.0.0  
**Last Updated:** April 2026

---

## Table of Contents

1. [Design Philosophy](#design-philosophy)
2. [Layered Architecture](#layered-architecture)
3. [Core Components](#core-components)
4. [Security Model](#security-model)
5. [Calculation Engine](#calculation-engine)
6. [Data Flow](#data-flow)
7. [Error Handling](#error-handling)
8. [Performance Considerations](#performance-considerations)

---

## Design Philosophy

### Governance-First
Every destructive operation requires explicit approval through HMAC-SHA256 scoped tokens. The system analyzes impact before allowing mutations, preventing agents from inadvertently breaking formula chains.

### AI-Native
All tools communicate via JSON stdin/stdout with standardized exit codes (0-5). Stateless design enables seamless tool chaining in agent orchestration frameworks.

### Headless Operation
Zero dependency on Microsoft Excel or COM interfaces. Two-tier calculation engine provides full functionality without GUI automation.

### Safety by Default
- Clone-before-edit workflow enforced
- Dependency analysis before structural changes
- Audit trail for all destructive operations
- Token scoping prevents privilege escalation

---

## Layered Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    AI Agent Orchestrator                         │
│              (Claude, GPT, Custom Agent Framework)               │
└───────────────────────┬───────────────────────────────────────────┘
                        │ JSON / Exit Codes
┌───────────────────────▼───────────────────────────────────────────┐
│                      CLI Tool Layer (53 Tools)                     │
│  ┌──────────┬──────────┬──────────┬──────────┬──────────┐       │
│  │Governance│   Read   │  Write   │ Structure│  Cells   │       │
│  │   (6)    │   (7)    │   (4)    │   (8)    │   (4)    │       │
│  ├──────────┼──────────┼──────────┼──────────┼──────────┤       │
│  │ Formulas │ Objects  │ Formatting│ Macros  │  Export  │       │
│  │   (6)    │   (5)    │   (5)    │   (5)    │   (3)    │       │
│  └──────────┴──────────┴──────────┴──────────┴──────────┘       │
└───────────────────────┬───────────────────────────────────────────┘
                        │ Protocol
┌───────────────────────▼───────────────────────────────────────────┐
│                    Core Hub Layer                                  │
│  ┌─────────────────┬─────────────────┬──────────────────┐       │
│  │   ExcelAgent    │ DependencyTrack │  TokenManager    │       │
│  │   (Context Mgr) │    (Graph)      │   (HMAC-SHA256)  │       │
│  ├─────────────────┼─────────────────┼──────────────────┤       │
│  │   FileLock      │   RangeSerial   │   AuditTrail     │       │
│  │  (OS-level)     │   (A1/R1C1)     │   (JSONL)        │       │
│  ├─────────────────┼─────────────────┼──────────────────┤       │
│  │   VersionHash   │   MacroHandler  │   ChunkedIO      │       │
│  │  (Geometry)     │   (oletools)    │   (Streaming)    │       │
│  └─────────────────┴─────────────────┴──────────────────┘       │
└───────────────────────┬───────────────────────────────────────────┘
                        │ Libraries
┌───────────────────────▼───────────────────────────────────────────┐
│                     Library Layer                                    │
│  ┌──────────┬──────────┬──────────┬──────────┬──────────┐       │
│  │openpyxl  │ formulas │ oletools │defusedxml│ jsonschema│       │
│  │(I/O)     │(Tier 1)  │(Macros)  │(Security)│(Schemas)  │       │
│  └──────────┴──────────┴──────────┴──────────┴──────────┘       │
└─────────────────────────────────────────────────────────────────┘
```

### Component Boundaries

| Layer | Responsibility | Isolation Guarantee |
|-------|---------------|---------------------|
| **CLI Tools** | Argparse, JSON I/O, Exit codes | No direct file I/O |
| **Core Hub** | Business logic, safety enforcement | Transactional operations |
| **Libraries** | File parsing, formula execution | Sandboxed (no eval) |

---

## Core Components

### ExcelAgent: Context Manager

```python
with ExcelAgent(path, mode="rw") as agent:
    # Lock → Load → Hash → (Modify) → Verify → Save → Unlock
    agent.workbook["Sheet1"]["A1"] = 42
    # Hash verification on exit
```

**Lifecycle:**
1. **Lock**: OS-level file lock via `fcntl` (Unix) / `msvcrt` (Windows)
2. **Load**: `openpyxl.load_workbook(keep_vba=True, data_only=False)`
3. **Hash**: Compute SHA-256 of workbook geometry (structure + formulas)
4. **Modify**: Agent performs mutations
5. **Verify**: Re-read file, compare hash (concurrent modification detection)
6. **Save**: Atomic write via `.save()`
7. **Unlock**: Release file lock

**Key Properties:**
- Thread-safe: One agent per workbook per process
- Reentrant: Nested contexts on same file use same lock
- Fail-safe: Lock released even on exception

### DependencyTracker: Formula Graph

```python
tracker = DependencyTracker(workbook)
tracker.build_graph()  # Lazy construction

# Impact analysis before deletion
report = tracker.impact_report("Sheet1!A5:A10", action="delete")
if report.broken_references > 0:
    # Return denial with guidance
    return ImpactDeniedError(report)
```

**Graph Construction:**
- Tokenize formulas with `openpyxl.utils.Tokenizer`
- Extract OPERAND tokens of subtype RANGE
- Build forward graph: `{cell: {dependents}}`
- Build reverse graph: `{cell: {precedents}}`
- Detect circular refs with Tarjan's SCC algorithm

**Performance:**
- 10-sheet, 1000-formula workbook: <5 seconds
- Memory: O(V + E) where V=cells, E=dependencies

### ApprovalTokenManager: HMAC-SHA256

```python
manager = ApprovalTokenManager()
token = manager.generate_token(
    scope="sheet:delete",
    target_file_hash="sha256:abc123...",
    ttl_seconds=300
)

# Validation
manager.validate_token(
    token_str,
    expected_scope="sheet:delete",
    expected_file_hash="sha256:abc123..."
)
```

**Token Structure:**
```json
{
  "scope": "sheet:delete",
  "target_file_hash": "sha256:...",
  "nonce": "uuid4",
  "issued_at": 1712585600,
  "ttl_seconds": 300,
  "signature": "hmac-sha256(...)"
}
```

**Validation Order:**
1. Parse JSON
2. Verify scope matches
3. Verify file hash matches
4. Verify not expired (`issued_at + ttl > now`)
5. Verify nonce not revoked
6. HMAC-SHA256 comparison via `hmac.compare_digest()` (timing-safe)
7. Mark nonce as used

### AuditTrail: Pluggable Logging

```python
trail = AuditTrail(backend=JsonlAuditBackend())

trail.log_operation(
    tool="xls_delete_sheet",
    scope="sheet:delete",
    resource="Sheet1",
    action="delete",
    outcome="success",
    token_used=True,
    file_hash="sha256:...",
)
```

**JSONL Format:**
```json
{"timestamp":"2026-04-08T14:30:22Z","tool":"xls_delete_sheet","scope":"sheet:delete","resource":"Sheet1","action":"delete","outcome":"success","token_used":true,"file_hash":"sha256:abc...","pid":12345,"details":{}}
```

**Privacy Guards:**
- VBA source code never logged
- Formula content excluded from audit
- Only structural changes recorded

---

## Security Model

### Token Scopes

| Scope | Risk Level | Requires Token | Operations |
|-------|-----------|----------------|------------|
| `sheet:delete` | High | Yes | Remove entire sheet |
| `sheet:rename` | Medium | Yes | Rename with ref update |
| `range:delete` | High | Yes | Delete rows/columns |
| `formula:convert` | High | Yes | Formulas → values |
| `macro:remove` | Critical | Yes+ | Strip VBA project |
| `macro:inject` | Critical | Yes | Add VBA project |
| `structure:modify` | High | Yes | Batch structural changes |

### Safety Protocols

#### 1. Clone-Before-Edit
```
Agent: "Modify financials.xlsx"
System: Clone to /work/financials_20260409T143022_abc123.xlsx
Agent: Work on clone
System: Audit log records clone_path
```

#### 2. Impact Analysis
```
Agent: "Delete Sheet1"
System: Scan dependencies...
System: Found 7 formulas referencing Sheet1
System: DENIED with guidance: "Run xls_update_references..."
Agent: Run xls_update_references --updates [...]
Agent: Delete Sheet1 --acknowledge-impact
System: APPROVED
```

#### 3. Concurrent Modification Detection
```
Agent: Opens file (hash=sha256:abc123)
[External process modifies file]
Agent: Save
System: Recompute hash → sha256:xyz789
System: DENIED - ConcurrentModificationError
```

### Path Traversal Prevention

```python
def validate_path(path: Path) -> Path:
    """Rejects ../ traversal and symlinks."""
    resolved = path.resolve()
    if not resolved.is_relative_to(ALLOWED_BASE):
        raise ValidationError("Path traversal detected")
    if resolved.is_symlink():
        raise ValidationError("Symlinks not allowed")
    return resolved
```

---

## Calculation Engine

### Tier 1: In-Process (formulas library)

```python
from formulas import ExcelModel

xl_model = ExcelModel().loads(path).finish()
xl_model.calculate()
xl_model.write(dirpath=output_dir)
```

**Capabilities:**
- SUM, AVERAGE, IF, VLOOKUP, INDEX, MATCH
- ~300 Excel functions supported
- Circular reference handling with `circular=True`
- ~50ms for 10k formulas

**Limitations:**
- Some complex functions unsupported
- External links not resolved
- Array formulas limited

### Tier 2: LibreOffice Headless

```python
subprocess.run([
    "soffice", "--headless", "--calc",
    "--convert-to", "xlsx",
    "--outdir", output_dir,
    input_file
], timeout=60)
```

**Capabilities:**
- Full Excel function support
- Pivot tables
- Power Query
- External links

**Performance:**
- 1-5s startup overhead
- Best for complex recalculations

### Auto-Fallback Strategy

```python
def recalculate(path: Path, *, tier: int | None = None) -> CalculationResult:
    if tier == 1 or tier is None:
        try:
            return tier1_calculate(path)
        except UnsupportedFunctionError:
            pass  # Fall through to Tier 2
    
    if tier == 2 or tier is None:
        if Tier2Calculator.is_available():
            return tier2_calculate(path)
        else:
            raise RuntimeError("LibreOffice unavailable for Tier 2")
```

---

## Data Flow

### Standard Write Operation

```
Agent Request
    ↓
[xls_write_range.py]
    ↓
Validate JSON input → Schema validation
    ↓
ExcelAgent.__enter__()
    ├── FileLock.acquire() [blocking, timeout=30s]
    ├── openpyxl.load_workbook()
    ├── compute_workbook_hash()
    ↓
Write data to range
    ↓
DependencyTracker.find_dependents()
    ├── Update affected formulas
    └── Return impact report
    ↓
ExcelAgent.__exit__()
    ├── verify_no_concurrent_modification()
    ├── workbook.save()
    ├── FileLock.release()
    ↓
AuditTrail.log_operation()
    ↓
JSON Response → Agent
```

### Token-Gated Destructive Operation

```
Agent Request (no token)
    ↓
[xls_delete_sheet.py]
    ↓
ExcelAgent.__enter__()
    ↓
DependencyTracker.impact_report()
    ├── Find dependents of target sheet
    ├── Count broken references
    └── Build ImpactReport
    ↓
ImpactReport.broken_references > 0?
    ├── YES → ImpactDeniedError
    │           └── Return JSON:
    │               {"status": "denied",
    │                "guidance": "Run xls_update_references...",
    │                "impact": {...}}
    │
    └── NO → Continue to token validation
                ↓
            PermissionDeniedError
                └── Return JSON:
                    {"status": "denied",
                     "error": "Token required for sheet:delete"}

Agent Request (with token)
    ↓
[xls_delete_sheet.py --token <token>]
    ↓
ApprovalTokenManager.validate_token()
    ├── Check scope matches "sheet:delete"
    ├── Check file hash matches
    ├── Check not expired
    ├── HMAC-SHA256 verification
    └── Mark nonce used
    ↓
Delete sheet
    ↓
AuditTrail.log_operation(token_used=True)
    ↓
JSON Response: {"status": "success", ...}
```

---

## Error Handling

### Exception Hierarchy

```
ExcelAgentError (base)
├── ValidationError          → Exit 1
│   └── SchemaError
│   └── RangeError
├── FileNotFoundError        → Exit 2
├── LockContentionError      → Exit 3
├── PermissionDeniedError    → Exit 4
│   └── TokenExpiredError
│   └── TokenScopeError
│   └── TokenReplayError
├── ConcurrentModificationError → Exit 5
├── ImpactDeniedError        → Exit 1 (with guidance)
└── InternalError            → Exit 5
```

### JSON Error Response

```json
{
  "status": "error",
  "exit_code": 4,
  "error": "Permission denied: Token expired",
  "timestamp": "2026-04-08T14:30:22Z",
  "guidance": "Generate new token with: xls_approve_token --scope ..."
}
```

---

## Performance Considerations

### Benchmarks

| Operation | Target | Implementation |
|-----------|--------|----------------|
| Read 500k rows | <3s | Chunked streaming (10k rows/chunk) |
| Write 500k rows | <5s | Write-only mode, batch writes |
| Dependency graph (10s/1000f) | <5s | Tokenizer caching, lazy build |
| Tier 1 recalc (1000f) | <500ms | In-process, no COM |
| File lock acquire | <100ms | Exponential backoff, 0.1s start |

### Memory Optimization

```python
# Chunked reading for large files
def read_range_chunked(sheet, min_row, max_row, chunk_size=10_000):
    for i in range(min_row, max_row + 1, chunk_size):
        chunk_end = min(i + chunk_size - 1, max_row)
        rows = list(sheet.iter_rows(
            min_row=i, max_row=chunk_end,
            min_col=min_col, max_col=max_col
        ))
        yield rows
        # Rows garbage collected after yield
```

### Caching Strategy

| Data | Cache Strategy | TTL |
|------|---------------|-----|
| JSON Schemas | Module-level dict | Infinite (immutable) |
| Dependency Graph | Instance-level | Session (until rebuild) |
| Token Nonce Set | Memory set | Session (until revocation) |
| Workbook Hash | Instance variable | Context lifetime |

---

## Integration Points

### Agent Orchestration Framework

```python
# Python subprocess example
import subprocess

result = subprocess.run(
    ["xls_read_range", "--input", "data.xlsx", "--range", "A1:C10"],
    capture_output=True,
    text=True
)

data = json.loads(result.stdout)
if data["status"] == "success":
    rows = data["data"]["values"]
```

### CI/CD Pipeline

```yaml
# GitHub Actions
- name: Test Excel Agent Tools
  run: |
    pytest tests/integration/test_clone_modify_workflow.py
    pytest tests/integration/test_formula_dependency_workflow.py
  env:
    EXCEL_AGENT_SECRET: ${{ secrets.EXCEL_AGENT_SECRET }}
```

---

## Future Considerations

### Potential Enhancements

1. **Distributed Locking**: Redis/etcd for multi-node agent clusters
2. **Async I/O**: `aiofiles` for concurrent large file operations
3. **WebSocket Bridge**: Real-time streaming for browser-based agents
4. **Formula Compilation**: JIT to Python via `formulas` export
5. **Cloud Storage**: S3/GCS backends for workbook storage

### Research Areas

- LLM-native formula repair suggestions
- Automatic template extraction from exemplar files
- Predictive impact analysis using ML on formula patterns

---

**Document maintained by:** excel-agent-tools maintainers  
**License:** MIT
