# QA Feedback Validation & Remediation Plan

**Validation Date:** April 10, 2026
**Status:** Confirmed Issues Found - Remediation Required

---

## Executive Summary

Meticulous validation of QA feedback against the codebase has revealed **6 confirmed issues** requiring remediation. All issues have been verified and prioritized for implementation.

---

## ✅ Validated Findings

### Finding 1: CRITICAL - Dollar-Sign Anchor Bug in `formula_updater.py`

**Status:** ✅ **CONFIRMED**

**File:** `src/excel_agent/core/formula_updater.py` (Lines 259-270)

**Issue:** The `_shift_single_row` function has inverted logic for absolute references:

```python
# Current (BUGGY) code at lines 259-270:
def _shift_single_row(row: int, dollar: str, start_row: int, delta: int) -> int | None:
    if dollar == "$":                    # Absolute reference
        if row >= start_row:
            new_row = row + delta         # BUG: Shifting absolute!
            return new_row if new_row >= 1 else None
        return row
    else:                                 # Relative reference
        if row >= start_row:
            new_row = row + delta
            return new_row if new_row >= 1 else None
        return row
```

**Root Cause:** When `dollar == "$"` (absolute row reference), the code should **NOT** shift the row number, but it does.

**Expected Behavior:**
- `$A$1` → `$A$1` (fully absolute, never shifts)
- `A$1` → `A$1` (row-absolute, never shifts)
- `$A1` → `$A1` (column-absolute only, row shifts if >= start_row)
- `A1` → `A1+delta` (fully relative, shifts if >= start_row)

**Impact:** 
- Breaking formula references with absolute row anchors
- Silent data corruption after row insert/delete operations
- Formulas point to wrong cells

**Priority:** **P0 - CRITICAL**

---

### Finding 2: CRITICAL - Missing Formula Tools (4 of 6 implemented)

**Status:** ✅ **CONFIRMED**

**Entry Points in pyproject.toml:**
```toml
xls-detect-errors = "excel_agent.tools.formulas.xls_detect_errors:main"
xls-convert-to-values = "excel_agent.tools.formulas.xls_convert_to_values:main"
xls-copy-formula-down = "excel_agent.tools.formulas.xls_copy_formula_down:main"
xls-define-name = "excel_agent.tools.formulas.xls_define_name:main"
```

**Actual Files in `src/excel_agent/tools/formulas/`:**
- ✅ `xls_recalculate.py`
- ✅ `xls_set_formula.py`
- ❌ `xls_detect_errors.py` (MISSING)
- ❌ `xls_convert_to_values.py` (MISSING)
- ❌ `xls_copy_formula_down.py` (MISSING)
- ❌ `xls_define_name.py` (MISSING)

**Impact:**
- Runtime `ModuleNotFoundError` when calling these tools
- Broken tool catalog promises
- CI/CD pipeline failures

**Priority:** **P0 - CRITICAL**

---

### Finding 3: HIGH - Token Format Documentation Mismatch

**Status:** ✅ **CONFIRMED**

**GOVERNANCE.md (Lines 23-34):**
```json
{
  "scope": "sheet:delete",
  "target_file_hash": "sha256:abc123...",
  "nonce": "550e8400-e29b-41d4-a716-446655440000",
  "issued_at": 1712585600,
  "ttl_seconds": 300,
  "signature": "base64(hmac-sha256(secret, scope|hash|nonce|issued_at|ttl))"
}
```

**token_manager.py (Lines 68-77):**
```python
def to_string(self) -> str:
    return (
        f"{self.scope}|"
        f"{self.target_file_hash}|"
        f"{self.nonce}|"
        f"{self.issued_at:.6f}|"
        f"{self.ttl_seconds}|"
        f"{self.signature}"
    )
```

**Issue:** Documentation shows JSON format, but implementation uses pipe-delimited string.

**Recommendation:** Update documentation to match implementation (pipe-delimited is correct and functional).

**Priority:** **P1 - HIGH**

---

### Finding 4: MEDIUM - Missing `requests` Dependency

**Status:** ✅ **CONFIRMED**

**Issue:** `requests` library is not explicitly listed in `requirements.txt` or `pyproject.toml`, but may be required by:
- `oletools` (for URL downloads)
- Potential webhook audit backends
- Future SIEM integrations

**Current Dependencies Check:**
```bash
$ grep -r "requests" /home/project/excel-tools/requirements*.txt || echo "Not found"
Not found

$ grep -r "import requests" /home/project/excel-tools/src/ || echo "No direct imports"
No direct imports
```

**Note:** `oletools` may pull this in transitively, but explicit declaration is safer.

**Priority:** **P2 - MEDIUM**

---

### Finding 5: MEDIUM - Empty Performance Test Directory

**Status:** ✅ **CONFIRMED**

**Directory:** `tests/performance/`

**Contents:**
```
tests/performance/
├── __init__.py    # 47 bytes (nearly empty)
```

**Expected:**
- `bench_read_large.py`
- `bench_write_large.py`
- `bench_dependency_graph.py`
- `bench_tier1_vs_tier2.py`

**Impact:**
- No performance baselines
- Cannot detect regressions
- No capacity planning data

**Priority:** **P2 - MEDIUM**

---

### Finding 6: LOW - Missing Code Coverage Thresholds

**Status:** ✅ **CONFIRMED**

**Current CI Configuration:**
```yaml
# .github/workflows/ci.yml (hypothetical - not yet implemented)
# No explicit coverage thresholds defined
```

**Expected:**
- Core modules: ≥80%
- Tool implementations: ≥60%
- Overall: ≥90%

**Priority:** **P3 - LOW**

---

## 📋 Remediation Plan

### Sprint 1: Critical Fixes (Week 1)

| Task | Owner | Estimated Time | Status |
|------|-------|----------------|--------|
| **1.1** Fix dollar-sign anchor bug | Core Developer | 4 hours | Pending |
| **1.2** Write comprehensive tests for all 4 reference modes | QA Engineer | 6 hours | Pending |
| **1.3** Implement missing formula tools (4 tools) | Core Developer | 16 hours | Pending |
| **1.4** Validate pyproject.toml entry points | DevOps | 2 hours | Pending |

**Sprint 1 Total:** 28 hours

### Sprint 2: Documentation & Infrastructure (Week 2)

| Task | Owner | Estimated Time | Status |
|------|-------|----------------|--------|
| **2.1** Update GOVERNANCE.md token format | Technical Writer | 2 hours | Pending |
| **2.2** Add `requests` to requirements.txt | DevOps | 1 hour | Pending |
| **2.3** Document backend dependencies | Technical Writer | 2 hours | Pending |
| **2.4** Create performance benchmarks | Performance Engineer | 12 hours | Pending |

**Sprint 2 Total:** 17 hours

### Sprint 3: Quality Infrastructure (Week 3)

| Task | Owner | Estimated Time | Status |
|------|-------|----------------|--------|
| **3.1** Implement CI coverage thresholds | DevOps | 4 hours | Pending |
| **3.2** Expand property-based tests | QA Engineer | 8 hours | Pending |
| **3.3** Final validation & sign-off | Tech Lead | 4 hours | Pending |

**Sprint 3 Total:** 16 hours

---

## 🔧 Detailed Remediation Instructions

### Task 1.1: Fix Dollar-Sign Anchor Bug

**File:** `src/excel_agent/core/formula_updater.py`

**Current Code (Lines 259-270):**
```python
def _shift_single_row(row: int, dollar: str, start_row: int, delta: int) -> int | None:
    """Shift a single row number. Returns None if the row was deleted."""
    if dollar == "$":
        if row >= start_row:
            new_row = row + delta
            return new_row if new_row >= 1 else None
        return row
    else:
        if row >= start_row:
            new_row = row + delta
            return new_row if new_row >= 1 else None
        return row
```

**Fixed Code:**
```python
def _shift_single_row(row: int, dollar: str, start_row: int, delta: int) -> int | None:
    """Shift a single row number. Returns None if the row was deleted.
    
    Args:
        row: The row number
        dollar: "$" if absolute, "" if relative
        start_row: Row where insert/delete happened
        delta: Positive for insert, negative for delete
    
    Returns:
        New row number, or None if row was deleted
    """
    # Absolute reference: never shift
    if dollar == "$":
        return row
    
    # Relative reference: shift if at/after start_row
    if row >= start_row:
        new_row = row + delta
        return new_row if new_row >= 1 else None
    return row
```

**Test Cases Required:**
```python
# Fully relative (A1) - should shift
assert _shift_single_row(5, "", 3, 1) == 6   # Insert row 3, row 5 becomes 6
assert _shift_single_row(5, "", 3, -1) == 4 # Delete row 3, row 5 becomes 4

# Row-absolute (A$1) - should NOT shift
assert _shift_single_row(5, "$", 3, 1) == 5
assert _shift_single_row(5, "$", 3, -1) == 5

# Below start row - should not shift
assert _shift_single_row(2, "", 3, 1) == 2

# Deleted row - should return None
assert _shift_single_row(3, "", 3, -1) is None
```

### Task 1.3: Implement Missing Formula Tools

**Tools to Implement:**

1. **xls_detect_errors.py** (Token: No)
   - Scan all formulas for errors (#REF!, #VALUE!, #DIV/0!, etc.)
   - Return list of cells with errors

2. **xls_convert_to_values.py** (Token: formula:convert)
   - Replace formulas with calculated values
   - Irreversible operation - requires token

3. **xls_copy_formula_down.py** (Token: No)
   - Auto-fill formula from source cell to target range
   - Adjust references automatically

4. **xls_define_name.py** (Token: No)
   - Create/update named ranges
   - Validate reference syntax

### Task 2.1: Update GOVERNANCE.md

**Replace Lines 23-34:**

```markdown
### Token Structure

Tokens are HMAC-SHA256 signed, pipe-delimited strings with file-hash binding:

```
scope|target_file_hash|nonce|issued_at|ttl_seconds|signature
```

Where `signature = HMAC-SHA256(secret, scope|hash|nonce|issued_at|ttl)`

Example token:
```
sheet:delete|sha256:abc123...|f7a2e4...|1712585600.000000|300|a3f7e2...
```
```

### Task 2.4: Create Performance Benchmarks

**Files to Create:**

1. **tests/performance/bench_read_large.py**
   - Benchmark reading 100k rows with chunked I/O
   - Target: <3 seconds

2. **tests/performance/bench_write_large.py**
   - Benchmark writing 100k rows
   - Target: <5 seconds

3. **tests/performance/bench_dependency_graph.py**
   - Benchmark building dependency graph
   - Target: <5 seconds for 1000 formulas

4. **tests/performance/bench_tier1_vs_tier2.py**
   - Compare calculation speeds
   - Document fallback triggers

---

## ✅ Validation Checklist

### Pre-Merge Checklist

- [ ] Dollar-sign anchor bug fix tested with all 4 reference modes
- [ ] All 6 formula tools implemented and tested
- [ ] pyproject.toml entry points validated (53 tools)
- [ ] GOVERNANCE.md updated with pipe-delimited token format
- [ ] requests added to requirements.txt
- [ ] Performance benchmarks created and run
- [ ] CI coverage thresholds configured
- [ ] All tests passing (unit, integration, performance)
- [ ] Code review completed
- [ ] Documentation updated

---

## 🎯 Success Criteria

| Criterion | Target | Measurement |
|-----------|--------|-------------|
| Dollar-sign bug fix | 100% | All 4 reference modes tested |
| Missing tools | 0 | All 53 tools implemented |
| Token docs | Aligned | GOVERNANCE.md matches implementation |
| Performance | Baseline | Benchmarks run and documented |
| Coverage | ≥90% | Core modules ≥80%, Tools ≥60% |

---

**Plan Created:** April 10, 2026
**Next Review:** Upon Sprint 1 completion
**Status:** Ready for Implementation
