## Executive verdict (claim-by-claim)

| Claim (from your description / repo tagline) | What I found in the repo | Verdict |
|---|---|---|
| “Headless Python CLI suite” | `pyproject.toml` defines **53 console scripts** (`xls-*`) implemented as Python modules under `src/excel_agent/tools/*`.  | **Supported** |
| “53 stateless tools” | The “53 tools” count is explicitly encoded in `[project.scripts]` (not just marketing text). “Stateless” is used in docs to mean “CLI processes with JSON I/O (no daemon / no TTY assumptions)”.  | **Supported (with nuance)** |
| “Safely read, mutate, calculate, and export Excel workbooks” | Tool entry points exist for reading ranges/metadata, writing, structural edits, formula recalculation, and exporting to CSV/JSON/PDF.  | **Supported at the tool-surface level** |
| “Without Microsoft Excel or COM dependencies” | Dependencies are `openpyxl`, `formulas`, etc.; no Excel/COM automation libs are required. Optional **LibreOffice headless** is used for a “Tier 2” recalculation path and PDF export.  | **Supported** |
| “Enforces governance via cryptographic tokens” | There is a real HMAC-SHA256 token system (scope + workbook hash binding + TTL + nonce + constant-time signature check).  | **Supported** |
| “Dependency-aware pre-flight checks” | There is a real dependency graph builder using `openpyxl`’s formula tokenizer, plus impact reporting; destructive tools (example: delete sheet) block by default if references would break and return prescriptive guidance.  | **Supported** |
| “Clone-before-edit workflows” | There is an explicit `xls-clone-workbook` tool, and the docs describe clone-before-edit as a core principle. However, common CLI args describe `--output` defaulting to overwrite input **with `--force`**, so cloning is strongly encouraged but not strictly mandatory in all flows.  | **Partially supported (policy + tooling, not absolute enforcement)** |
| “Immutable audit trails” | Audit logging exists (default `.excel_agent_audit.jsonl`). I did **not** see evidence (in what’s exposed in code/docs I reviewed) of cryptographic tamper-evidence (hash chaining/signing). “Immutable” looks like an architectural intent/discipline rather than a cryptographic guarantee.  | **Partially supported** |
| “Production-grade” | Repo has CI workflows, structured tests (unit/integration/performance/property), strict typing/lint config, standardized exit codes + JSON envelopes. That said, “production-grade” is ultimately a judgment call; the architecture doc includes self-reported readiness metrics that should be independently verified by running CI in your environment.  | **Mostly supported (but ‘production-grade’ depends on your bar)** |

---

## What’s actually implemented (deep dive, anchored to code)

### 1) The “53 tools” claim is real, not hand-wavy
The strongest possible verification here is that the count is enforced in packaging:

- `pyproject.toml`’s `[project.scripts]` enumerates **exactly 53** `xls-*` command entry points, grouped by category (Governance/Read/Write/Structure/Cells/Formulas/Objects/Formatting/Macros/Export).   
- The README mirrors this with a “Tool Catalog (53 Tools)” section and a category table.   

This means the “53” claim is *operationally true* at install time (i.e., what ends up on `$PATH`).

---

### 2) The CLI contract is designed for agents (JSON-only stdout + deterministic exit codes)
A key “AI-native” property is: tools behave like pure functions from `(args, stdin)` → `(json stdout, exit code)`.

- The shared runner `run_tool()` in `_tool_base.py` enforces standardized error handling and prints JSON responses (including structured “impact” and “guidance” fields when relevant), and maps exceptions to standardized exit codes.   
- The architecture doc explicitly calls out “Strict JSON stdout” and exit codes `0–5` as a design principle.   

This is a real differentiator vs. many “Excel scripts” that intermix logs + data.

---

### 3) Governance tokens: cryptographically scoped approvals (HMAC-SHA256) exist in code
The repo’s governance claims are unusually concrete.

#### Token properties implemented
`token_manager.py` documents and implements:

- **Scoped tokens** (e.g., `sheet:delete`)   
- **Workbook binding** via a file hash (prevents reusing a token on a different workbook)   
- **TTL** with defaults and a max window   
- **Nonce** + replay detection (single-use)   
- **Constant-time signature compare** via `hmac.compare_digest`   

#### Example: `xls-delete-sheet` really enforces the gate
`xls_delete_sheet.py` shows the whole governance pattern in one place:

- Computes workbook hash, requires a token, validates it for expected scope + file hash.   
- Runs a dependency pre-flight; if broken references would occur and impact isn’t acknowledged, it throws an `ImpactDeniedError` with an impact report + guidance telling the agent what to do next (e.g., run `xls-update-references` or re-run with `--acknowledge-impact`).   
- Logs an audit event after success.   

So: this isn’t a README-only promise—the enforcement is coded into at least this destructive operation.

---

### 4) Dependency-aware pre-flight checks: there’s a real formula graph, not a placeholder
The core safety claim is “structural edits can silently break formulas; we prevent that.”

In `core/dependency.py`, the repo builds a formula reference graph by:

- Parsing formulas using `openpyxl`’s `Tokenizer` and collecting range references.   
- Computing strongly connected components with an explicit **Tarjan SCC** implementation (iterative) to detect cycles.   
- Wrapping this as a `DependencyTracker` whose docstring explicitly frames it as “pre-flight impact analysis.”   

Then tools can do what `xls-delete-sheet` does: generate an impact report *before* mutating and deny-by-default if it would break references.   

This is exactly the kind of “agent-safe guardrail” that matters in autonomous orchestration.

---

### 5) “Calculate formulas” without Excel: implemented as a two-tier engine (Python-first, LibreOffice fallback)
Excel formula recalculation is the hardest part of “Excel-without-Excel.”

The repo’s approach is explicit:

- `xls-recalculate` tries **Tier 1** (the Python `formulas` library) and falls back to **Tier 2** (LibreOffice headless) if there are unsupported functions or errors.   
- Tier 1 engine loads a workbook into `formulas.ExcelModel()` and calls `calculate()`.   
- Tier 2 uses `soffice --headless ... --convert-to xlsx` to force a recalc via LibreOffice and write back out.   
- The README explicitly positions LibreOffice as an optional dependency for “full-fidelity recalculation (Tier 2).”   

**Important nuance for your “Excel equivalent” goal:** this is credible engineering, but it is not proof of “Excel parity.” It’s a pragmatic hybrid: “fast Python where possible; outsource edge cases to a spreadsheet engine.”

---

### 6) File safety beyond formulas: locking + concurrency patterns exist
For multi-agent or multi-process safety, workbook locking matters.

- `core/locking.py` implements a sidecar lock file (`.<filename>.lock`) with timeout + retry/backoff and contention detection.   
- The exception taxonomy includes lock contention and concurrent modification concepts mapped to standardized exit codes.   

This supports the “safe in automation” story (especially in CI runners / worker pools).

---

### 7) Audit trail exists, but “immutable” should be read carefully
- `audit_trail.py` defines `DEFAULT_AUDIT_FILE = ".excel_agent_audit.jsonl"` and supports reading events back.   
- `xls-delete-sheet` logs operations after success.   

**But:** “immutable” in a strict governance/compliance sense usually means tamper-evident logs (hash chains, signatures, remote append-only storage, WORM). I did not see direct evidence of that guarantee in the portions inspected—so treat it as “append-only by convention” unless you add tamper-evidence.

---

## “Headless / no Excel / no COM” — validated
Two independent confirmations:

1) README explicitly claims “Zero Microsoft Excel dependency… optional LibreOffice.”   
2) Packaging dependencies list `openpyxl`, `formulas[excel]`, `oletools`, etc., and requires Python ≥ 3.12. No COM automation dependencies are present.   

So the repo is genuinely cross-platform in principle (Windows/macOS/Linux), with LibreOffice as the only external “engine-like” dependency.   

---

## “Clone-before-edit workflows” — what’s true vs. what’s aspirational
What’s clearly true:

- There is a dedicated cloning tool (`xls-clone-workbook`) and an SDK usage example that clones before edits.   

What’s not strictly proven as “enforced everywhere”:

- The shared CLI helper text says `--output` defaults to overwriting input, but “requires --force for safety.” That implies in-place edits are possible as an escape hatch, even if discouraged.   

So: the repo provides a **clone-first happy path**, but it’s not a hard “you can’t edit originals” rule across the entire surface area (at least not at the argument-contract layer).

---

## Production-grade signals (and what I’d still verify by running it)

### Strong signals present in-repo
- A real test suite layout: `tests/unit`, `tests/integration`, `tests/performance`, `tests/property`, plus fixtures.   
- GitHub Actions workflows exist.   
- Strict typing/linting/tooling is configured (mypy strict, ruff, etc.) and modern dependency pin floors.   
- Safety mechanisms are implemented in code (tokens, pre-flight dependency analysis, locking, standardized error envelope).   

### What “production-grade” still depends on (not fully provable by static reading)
- Whether CI is green on all supported OS/Python versions in practice (you’d confirm by checking recent workflow runs). (Workflows exist, but static presence ≠ passing).   
- Real-world Excel compatibility (especially around pivot tables, external links, slicers, PowerQuery, volatile functions, locale-specific formats, etc.). The tool surface suggests broad coverage, but Excel is enormous.   

---

## How this maps to the “Excel equivalent for AI agents” vision

This repo looks like a **tool layer / substrate** for an “AI-native spreadsheet system,” not a complete Excel replacement by itself:

- It focuses on *workbook manipulation* (read/write/structure/objects) and *governance* (tokens, audit, pre-flight), which is exactly what agents need.   
- It does **not** (based on the tool catalog visible) claim pivot-table authoring/refresh, full chart parity, PowerQuery equivalence, or a full calculation engine written from scratch—rather it uses a hybrid “formulas library + LibreOffice fallback.” That’s a sane approach, but it’s not “Excel reimplemented.”   

If the “latest project” goal is truly “an Excel equivalent for agents,” this repo is a solid foundation for:
- deterministic tool contracts,
- safe mutation policies,
- dependency impact governance,
- headless execution.

But the remaining hardest gaps for “Excel equivalence” are typically:
- full-fidelity calculation semantics,
- pivot table engine behavior,
- chart rendering fidelity,
- compatibility with the long tail of Excel features (named formulas, dynamic arrays edge cases, external refs, etc.).

---

## Bottom line
Your quoted capability claims are **substantially valid** as descriptions of what’s implemented—especially the *53-tool CLI surface*, *headless/no-Excel approach*, *token governance*, and *dependency-aware pre-flight checks*.   

The main places where the repo’s phrasing is stronger than what I can strictly verify from the inspected code are:
- **“immutable” audit trail** (logging exists; cryptographic tamper-evidence isn’t clearly shown).   
- **“clone-before-edit enforced”** (clone tooling exists and is promoted; the shared CLI contract still allows overwrite-with-force).   

If you want to treat this as infrastructure for AI-friendly Office replacements, the repo’s design decisions (JSON-only stdout, denial-with-guidance, scoped approvals, pre-flight dependency graphs) are exactly aligned with “CLI access is water to fish” for agents. 

