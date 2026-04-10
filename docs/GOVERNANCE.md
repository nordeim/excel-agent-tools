# Governance & Security: excel-agent-tools

**Version:** 1.0.0  
**Classification:** Security-Critical

This document defines the token lifecycle, audit schema, and safety protocols for excel-agent-tools.

---

## Table of Contents

1. [Token System](#token-system)
2. [Audit Trail](#audit-trail)
3. [Safety Protocols](#safety-protocols)
4. [Compliance](#compliance)

---

## Token System

### Token Structure

Tokens are HMAC-SHA256 signed, **pipe-delimited strings** with file-hash binding:

```
scope|target_file_hash|nonce|issued_at|ttl_seconds|signature
```

Where:
- `scope`: Operation scope (e.g., "sheet:delete")
- `target_file_hash`: SHA-256 hash of the target workbook
- `nonce`: Random 32-character hex string (single-use)
- `issued_at`: Unix timestamp with 6 decimal places
- `ttl_seconds`: Token lifetime in seconds
- `signature`: HMAC-SHA256 of the canonical string

The canonical string for signing is:
```
scope|target_file_hash|nonce|issued_at|ttl_seconds
```

Example token:
```
sheet:delete|sha256:abc123...|f7a2e4b8...|1712585600.123456|300|a3f7e2d9...
```

**Note:** Tokens are serialized as pipe-delimited strings, not JSON, for compactness and ease of CLI passing.

### Scopes

| Scope | Risk | Operations |
|-------|------|------------|
| `sheet:delete` | High | Remove entire sheet |
| `sheet:rename` | Medium | Rename with ref update |
| `range:delete` | High | Delete rows/columns |
| `formula:convert` | High | Formulas to values |
| `macro:remove` | Critical | Strip VBA (2 tokens) |
| `macro:inject` | Critical | Add VBA project |
| `structure:modify` | High | Batch mutations |

### Validation Order

1. Deserialize JSON
2. Verify scope matches expected
3. Verify file hash matches target
4. Verify not expired (issued_at + ttl > now)
5. Verify nonce not revoked
6. HMAC-SHA256 via `hmac.compare_digest()`
7. Mark nonce as used (single-use)

### CLI Usage

```bash
# Generate token
xls-approve-token --scope sheet:delete --file workbook.xlsx --ttl 300

# Use token
xls-delete-sheet --input workbook.xlsx --name Sheet1 \
  --token "eyJzY29wZSI6..." --acknowledge-impact
```

---

## Audit Trail

### Format

Append-only JSONL:

```json
{"timestamp":"2026-04-08T14:30:22Z","tool":"xls_delete_sheet","scope":"sheet:delete","resource":"Sheet1","action":"delete","outcome":"success","token_used":true,"file_hash":"sha256:abc...","pid":12345,"details":{}}
```

### Privacy

**Never logged:**
- VBA source code
- Formula content
- Token signatures
- Credentials

**Always logged:**
- Tool name
- Scope and nonce
- File hash
- Timestamp
- Impact metrics

---

## Safety Protocols

### Clone-Before-Edit

Always work on clones:

```bash
xls-clone-workbook --input original.xlsx --output-dir ./work/
# Work on clone, never original
```

### Impact Analysis

Pre-flight check for destructive ops:

```bash
xls-dependency-report --input workbook.xlsx
# Returns broken_references count
```

### Concurrent Modification

Hash-based detection prevents external changes:

```
1. Open file (hash=abc123)
2. [External process modifies]
3. Save attempt
4. System detects hash mismatch (xyz789)
5. DENIED with ConcurrentModificationError
```

---

## Compliance

### Checklist

- [ ] Secret key generated and secured
- [ ] Audit backend configured
- [ ] Token TTL set per risk tolerance
- [ ] Emergency revocation documented

### Rotation

- **Token secrets:** Annually
- **Audit review:** Monthly
- **Scope audit:** Quarterly

---

**Maintained by:** Security Team  
**License:** MIT
