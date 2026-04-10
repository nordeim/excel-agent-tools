## What to build: a “real-office” fixture pack that stresses the exact tool surface

The repo exposes **53 CLI entry points** (the `xls-*` commands) across governance/read/write/structure/cells/formulas/objects/formatting/macros/export.   
Your dummy data should be designed so that *every category* is exercised in a way that resembles actual office work: ingest data → clean/tabulate → compute → visualize → publish/export → govern risky changes.

Below is a fixture pack that’s intentionally “office-realistic” but also contains **booby traps** (structured references, external links, circular refs, macro risk patterns) to expose gaps.

---

## Fixture Pack (dummy workbooks + macro binaries)

### Fixture 1 — `OfficeOps_Expenses_KPI.xlsx` (the main “real life” workbook)
Purpose: emulate a common analyst workflow (expenses + budget + KPI dashboard) while hitting formulas, formatting, objects, validations, and dependency risk.

**Sheets**
1) `Lists`
- `A1:A10` Categories: Travel, Software, Meals, Shipping, Contractor, OfficeSupplies, Marketing, Cloud, Training, Other
- `B1:B6` Departments: Sales, Finance, Ops, Eng, HR, Exec
- Named ranges:
  - `Categories` → `Lists!$A$1:$A$10`
  - `Departments` → `Lists!$B$1:$B$6`
  - `TaxRate` → `Lists!$D$2` (value e.g. `0.0825`)
  - `ReportMonth` → `Lists!$D$3` (date e.g. `2026-03-01`)

2) `Raw_Expenses`
- Header row (row 1): `Date, Dept, Vendor, Category, Amount, Currency, FX, AmountUSD, Notes, ReceiptURL`
- Add **data validation** on:
  - `Dept` column: list = `=Departments`
  - `Category` column: list = `=Categories`
- Formulas (start row 2):
  - `FX` (G2): `=IF(F2="USD",1, XLOOKUP(F2, FXRates[Currency], FXRates[Rate]))`
  - `AmountUSD` (H2): `=E2*G2`
  - Add one deliberate error row: `=VLOOKUP("missing",Lists!A:B,2,FALSE)` to yield `#N/A` for `xls-detect-errors`
- Convert `A1:H{n}` into an **Excel Table** named `Expenses` (later via `xls-add-table`).   
- Include tricky strings in `Notes` (commas, quotes, newlines) to test CSV/JSON export quoting behavior.

3) `FXRates`
- Table `FXRates` with columns `Currency, Rate`
- Keep small, deterministic values (EUR=1.10, GBP=1.30, JPY=0.009)

4) `Summary`
- KPI cells:
  - Total spend: `=SUM(Expenses[AmountUSD])`  (structured reference)
  - Spend by department (a small matrix using `SUMIFS`)
  - Month filter: `=EOMONTH(ReportMonth,0)` etc.
- Add **conditional formatting**: highlight department totals > threshold (e.g. $10,000)

5) `Dashboard`
- Reserve space for charts (`G2` etc.)
- Add a merged title cell and freeze panes for “report feel” (tests merge/unmerge + freeze panes).

Why this fixture is valuable:
- It forces you to test **structured references** and named ranges—areas where many formula parsers and dependency checkers are weaker than with plain `Sheet!A1` references.
- It creates a plausible “exportable report artifact” suitable for PDF/JSON.

---

### Fixture 2 — `EdgeCases_Formulas_and_Links.xlsx`
Purpose: force calculation + dependency analysis edge cases.

Include:
- `Circular` sheet:
  - `A1 = B1 + 1`
  - `B1 = A1 + 1`
  - Used to test `xls-recalculate --circular` behavior. 
- `DynamicArrays` sheet:
  - `=UNIQUE(...)`, `=FILTER(...)`, `=LET(...)`, `=TEXTSPLIT(...)` (functions likely to trip Tier 1 and force Tier 2 fallback if LibreOffice exists). 
- `ExternalLinks` sheet:
  - A formula like `='[OtherBook.xlsx]Sheet1'!A1`
  - This tests whether dependency reporting / validation gracefully handles external references.

---

### Fixture 3 — Macro binaries + macro-enabled workbook
The repo’s macro tooling is **inspection + risk scoring + remove + inject**; it is *not* a macro runtime. 

Build:
1) `vbaProject_safe.bin`
- Contains a benign module, e.g. formatting numbers / cleaning cells (no auto-exec, no shell, no network)

2) `vbaProject_risky.bin`
- Contains patterns explicitly scanned for (examples include `AutoOpen`, `Shell`, `URLDownloadToFile`, and obfuscation via `Chr(`). 

3) `MacroTarget.xlsx`
- A normal `.xlsx` that you’ll inject `vbaProject_safe.bin` into to produce `MacroTarget_with_macros.xlsm` using `xls-inject-vba-project`. 

**How to produce the `.bin` files (practical approach)**
- Create two `.xlsm` files in Excel (or LibreOffice, if it can embed VBA as expected), each containing the desired VBA code.
- Rename `.xlsm` → `.zip`, extract `xl/vbaProject.bin`.
- Store those `.bin` files under `tests/fixtures/macros/`.

This is the cleanest “simulated but realistic” macro artifact because it uses real Office file internals.

---

## How to generate the workbook fixtures (deterministic + scalable)

Create a script `scripts/generate_fixtures.py` that:
- Uses `openpyxl` to create the sheets + headers + baseline formulas + formatting.
- Uses a **fixed random seed** to generate N rows of expenses (e.g. 200 rows for functional tests, 50k for perf tests).
- Writes values only; do **not** rely on Excel to compute anything at generation time (you’ll test recalculation using the tools).

You can keep the generator minimal: the CLI suite already provides tools to add tables/charts/CF/validations—so you can intentionally create some features via CLI in the E2E tests (that’s closer to “agent usage”).

---

## The E2E test plan (CLI-driven) — designed to expose “fit-for-use” gaps

### Test harness conventions
- Use `pytest` + `subprocess.run(...)`
- Parse stdout JSON and assert:
  - `exit_code`
  - `status`
  - presence/shape of `data`, `impact`, `guidance`
- Use a temp work dir per test (`tmp_path`)

This matches the project’s “strict JSON stdout + exit codes” contract. 

Also set a stable secret for tokens:
- In tests, explicitly pass a secret to the SDK or set an env var if supported by your wrapper; otherwise, token behavior may be nondeterministic (the token manager generates a random secret if none is provided). 

---

# Suite A — Smoke: “Do the tools run end-to-end at all?”

### A1) `--help` for all 53 commands
Goal: catch packaging/entrypoint issues quickly. (This also detects typo’d arguments and missing imports.)

Expected: exit code 0 for each.

### A2) Minimal read operations on `OfficeOps_Expenses_KPI.xlsx`
Commands:
- `xls-get-sheet-names --input ...`
- `xls-read-range --input ... --sheet Raw_Expenses --range A1:H5`
- `xls-get-defined-names --input ...`
- `xls-get-workbook-metadata --input ...`

Expected:
- `exit_code=0`
- JSON is parseable
- data contains expected fields

This validates the agent-friendly “read surface” first. 

---

# Suite B — Core “office workflow”: ingest → tabulate → compute → visualize → export

### B1) Clone-before-edit workflow
Command:
- `xls-clone-workbook --input OfficeOps_Expenses_KPI.xlsx --output-dir ./work`

Expected:
- Clone created with timestamped/unique name (whatever the tool outputs)
- Source remains unchanged

(Repo documents clone-first as the intended workflow, and common args warn that in-place overwrite requires `--force`.) 

### B2) Write new expense rows (simulate “agent gathered data”)
On the clone:
- `xls-write-range --input clone.xlsx --output clone.xlsx --sheet Raw_Expenses --range A2 --data '[[...],[...],...]'`

Expected:
- `impact.cells_modified == rows * cols`
- Values appear in workbook when re-opened with openpyxl

`xls-write-range` uses type inference (including treating strings beginning with `=` as formulas). 

### B3) Convert the data area into a Table
- `xls-add-table --input clone.xlsx --output clone.xlsx --sheet Raw_Expenses --range A1:J201 --name Expenses`

Expected:
- Table exists; table name validated; no overlap errors. 

### B4) Add a chart to the Dashboard
- `xls-add-chart --input clone.xlsx --output clone.xlsx --sheet Summary --type bar --data-range B1:E7 --categories-range A2:A7 --position G2 --title "Spend by Dept"`

Expected:
- Chart created; audit event logged. 

### B5) Recalculate
Run two variants:
1) Tier 1 forced:
- `xls-recalculate --input clone.xlsx --output clone_calc_t1.xlsx --tier 1`

2) Auto mode (Tier 1 then fallback to Tier 2 if needed):
- `xls-recalculate --input clone.xlsx --output clone_calc_auto.xlsx`

Expected:
- `engine` field indicates tier used
- If `unsupported_functions` non-empty, auto should attempt Tier 2 when LibreOffice is available. 

**Critical “fit-for-use” check:** Export tools like `xls-export-json` load with `data_only=True` (cached values). If your recalc path doesn’t write cached results in a way openpyxl can read, exports may be stale.   
So after recalculation, immediately run export and validate the totals match expected values.

### B6) Export JSON + CSV
- `xls-export-json --input clone_calc_auto.xlsx --sheet Raw_Expenses --range A1:J201 --orient records --pretty`
- `xls-export-csv --input clone_calc_auto.xlsx --sheet Raw_Expenses --range A1:J201` (verify quoting)

Expected:
- Record counts correct
- Date fields converted to ISO in JSON
- CSV quoting preserves commas/newlines in Notes

JSON export explicitly documents orientations and type conversion. 

### B7) Export PDF (LibreOffice required)
- `xls-export-pdf --input clone_calc_auto.xlsx --output out/report.pdf`

Expected:
- File exists, size > 0
- Optional: parse PDF with `pypdf` and assert at least 1 page

PDF export is LibreOffice-based. 

---

# Suite C — Governance + “safe mutation” (tokens, impact denial, remediation)

This suite is where you’ll discover whether the project is truly “governance-first” in practice (not just in README language). Tokens are HMAC-SHA256, scoped, TTL-bound, file-hash-bound, and single-use via nonce tracking. 

### C1) Token properties (scope/file-hash/TTL/replay)
Steps:
1) Generate token:
- `xls-approve-token --scope sheet:delete --file clone.xlsx --ttl 60` 

2) Use token successfully in a gated operation (sheet delete, row delete, etc.).
3) Re-use the same token (should fail as “replay detected”).
4) Modify the workbook, then try the old token (should fail as “file_hash_mismatch”).
5) Wait 61 seconds and try again (should fail as “expired”).

Expected:
- Exit codes map to “permission denied” for invalid/expired tokens per the exception design. 

### C2) Dependency impact denial on structural edits
Design the workbook so `Summary` references `Raw_Expenses`, then attempt:
- `xls-delete-sheet --name Raw_Expenses ...` without `--acknowledge-impact`

Expected:
- Denied with exit code 1, and includes an `impact` report + `guidance` telling you to run `xls-update-references` or acknowledge impact. 

### C3) Remediation: update references
After an impact denial, run:
- `xls-update-references --updates '[{"old":"Raw_Expenses!A2","new":"Raw_Expenses!A3"}]' ...`

Expected:
- `formulas_updated > 0` and details show sample changed formulas/defined names. 

**Gap-hunting note:** `xls-update-references` targets Tokenizer “range” operands and does string replacement. That’s strong for classic A1 refs, but it may miss:
- structured references (`Expenses[AmountUSD]`)
- formula text inside charts / conditional formatting rules
- external link tokens  
Your fixture uses structured refs specifically to test this. 

---

# Suite D — Formula tool correctness (including a likely “convert-to-values” gap)

### D1) Set formula + copy down
- `xls-set-formula --sheet Raw_Expenses --cell H2 --formula "=E2*G2"`
- `xls-copy-formula-down --sheet Raw_Expenses --source H2 --target-range H2:H201`

Expected:
- Formulas appear in all target cells (verify by opening workbook and checking `cell.data_type == 'f'` or via `xls-get-formula`).

### D2) Detect errors
- `xls-detect-errors --sheet Raw_Expenses --range A1:J201`

Expected:
- It finds the deliberate `#N/A` row. (This is your regression sentinel.)

### D3) Convert-to-values “truth test”
Claimed behavior: replace formulas with calculated values (irreversible, token-gated).   
But the current implementation text (as written) indicates it **clears formulas** rather than computing/writing calculated values. 

E2E steps:
1) Recalc workbook to produce correct values (`xls-recalculate`)
2) Generate token scope `formula:convert`
3) Run `xls-convert-to-values --range ... --token TOKEN`
4) Re-open output and check:
   - cells formerly containing formulas now contain numeric values
   - totals remain correct

Expected for “fit-for-use”:
- The numeric results persist and exports show correct totals

If results become blank/None or stale, you’ve discovered a major real-world gap: agents often need to “freeze” a report before sharing/exporting.

---

# Suite E — Macro workflows: detect/inspect/risk/strip/inject

### E1) Detect + inspect on macro-enabled file
- `xls-has-macros --input MacroTarget_with_macros.xlsm`
- `xls-inspect-macros --input ... --code-preview-length 200`

Expected:
- Lists modules; shows preview; signature fields present (even if unsigned). 

### E2) Risk scoring (safe vs risky)
- `xls-validate-macro-safety --input MacroTarget_with_macros.xlsm`
- `xls-validate-macro-safety --input RiskyWorkbook.xlsm`

Expected:
- Risky macro triggers warning/high due to regex patterns (`AutoOpen`, `Shell`, network IOCs, obfuscation). 

### E3) Remove macros (double token)
- Generate two tokens scope `macro:remove`
- `xls-remove-macros --input risky.xlsm --token TOKEN1 --token2 TOKEN2 --output risky_clean.xlsx`

Expected:
- Output has no `xl/vbaProject.bin` inside zip; tool warns irreversible. 

### E4) Inject macros (token + mandatory scan)
- Generate token scope `macro:inject`
- `xls-inject-vba-project --input MacroTarget.xlsx --vba-bin vbaProject_safe.bin --token TOKEN`

Expected:
- Success; audit trail contains event but not macro source. 

Also test risky injection:
- `xls-inject-vba-project ... vbaProject_risky.bin ...` without `--force`

Expected:
- Denied (exit code 1) when risk is high/critical. 

---

# Suite F — Concurrency + lock behavior (multi-agent realism)

`ExcelAgent` acquires an exclusive lock, computes file hashes, and refuses to save on concurrent modification. 

### F1) Lock contention
- Start a process that opens the workbook with a long-running `ExcelAgent` context (or hold a lock file if the lock is file-based).
- In parallel, run `xls-write-cell` (or any rw tool) on the same file.

Expected:
- Exit code 3 (lock contention) with retry guidance.

### F2) Concurrent modification detection
- Enter a write tool that holds workbook open (harder with CLI), then externally modify the file on disk, then let tool attempt save.

Expected:
- It refuses to save and exits with internal/concurrent modification error.

This is important for real orchestrators where agents can collide.

---

## “Gap discovery” checklist (what this plan is specifically trying to catch)

1) **Structured references & named ranges** not correctly handled by:
- dependency impact analysis
- reference updater
- error detector

2) **Calculation fidelity** gaps:
- Tier 1 coverage limitations
- Tier 2 dependency on LibreOffice availability / timeouts
- whether recalculated values are visible to `data_only=True` export tools

3) **Governance correctness**:
- token scope mismatch
- replay detection
- TTL enforcement
- file-hash binding after edits

4) **Macro safety realism**:
- false positives/negatives in regex-based scanning
- injection denial behavior for risky bins
- assurance that logs do not leak macro source

5) **Report finalization**:
- convert-to-values works (or fails) in a way that would break real “publish a report” workflows

---

## Deliverables you’ll have after running this plan

- A deterministic `tests/fixtures/` folder that represents real office complexity
- A CLI-driven E2E suite that:
  - validates “happy path” (agent can produce a dashboard + exports)
  - validates “governed path” (agent can’t do risky ops without approvals)
  - produces concrete failure artifacts (JSON responses, audit logs, output workbooks, PDFs)
- A clear “fit-for-use scorecard” showing which claims hold under realistic workloads

If you want to push this even harder, the next step is a “golden oracle” comparison: for the same workbook, compute results in **LibreOffice** and compare exported JSON numeric outputs cell-by-cell against Tier 1, to quantify calculation parity (not just pass/fail). 
