## Phased remediation plan (validated against the current repo code)

This plan closes the three gaps you called out by (1) **making output semantics unambiguous and safe**, (2) **delivering pivot tables in a way that is actually achievable headlessly**, and (3) **making macros vs `.xlsx` an explicit, enforced contract**. I’m grounding each phase in what the code *currently* does.

---

# Phase 0 — Stop-the-bleeding safety patch (1–3 days)

## 0.1 Fix the “output-path / input-mutation” foot-gun everywhere it exists
### What’s broken today (validated)
- `ExcelAgent.__exit__()` **always saves to the path it opened**: `self._wb.save(str(self._path))`.   
- Multiple mutating tools open `ExcelAgent(input_path, mode="rw")` *and then* do `wb.save(output_path)` when output differs. Example:  
  - `xls_write_range`: opens on `input_path`, and if output differs calls `wb.save(str(output_path))`.   
  - `xls_set_formula`: same pattern.   
  - `xls_delete_sheet`: same pattern.   
**Net effect:** if `--output` is different, the tool still exits the context manager and saves back onto `input_path` (because the agent opened the input), which can mutate your source workbook unexpectedly.   

### Remediation (minimal, high-confidence)
- **Rule:** if `--output` is provided and differs from `--input`, then the tool must *edit the output file*, not the input file.
- Implement a shared helper (new module), e.g. `excel_agent/core/io_targets.py`:

  1) Resolve `input_path`, `output_path`.  
  2) If `output_path != input_path`:
     - copy `input_path → output_path` (atomic copy like `xls_clone_workbook` already does with `shutil.copy2`).   
     - open `ExcelAgent(output_path, mode="rw")` and perform the mutation.  
     - **never** open `ExcelAgent(input_path)` for a “save-as” operation.
  3) If `output_path == input_path`:
     - allow in-place edit **only** with an explicit safety flag (see 0.2).

- In all tools that use `ExcelAgent`, **remove any `wb.save(output_path)` calls** and rely on `ExcelAgent.__exit__()` saving exactly once, to the correct file (the one it opened).   

### Acceptance tests (must add now)
- For each of 3 representative tools (`xls_write_range`, `xls_set_formula`, `xls_delete_sheet`):
  - create `input.xlsx` with sentinel value (and compute hash),
  - run tool with `--output out.xlsx`,
  - assert `input.xlsx` hash unchanged,
  - assert `out.xlsx` contains the mutation.   

## 0.2 Enforce the CLI safety contract for in-place edits
### What’s inconsistent today (validated)
- `add_common_args()` advertises: “default overwrite input — requires --force for safety”.   
- But many tools compute `output_path = validate_output_path(args.output or args.input, ...)` and then proceed without checking `--force`, so in-place overwrites happen silently. Example in `xls_delete_sheet`.   

### Remediation
- Add a shared resolver, e.g. `resolve_output_or_inplace(args)` in `cli_helpers.py` that:
  - If `args.output is None` (implies in-place), require `args.force == True` or exit with validation error.
  - If `args.output == args.input`, also require `--force`.
- Update all mutating tools to use this function rather than re-implementing “args.output or args.input” ad hoc.   

## 0.3 Fix token validation call sites that are currently incompatible with the token manager
This isn’t one of your three gaps, but it directly impacts “safe governance” while you refactor output behavior.

### What’s broken today (validated)
- `ApprovalTokenManager.validate_token()` takes `(token_string, scope, file_path)` and checks the file hash internally.   
- `xls_delete_sheet` calls it with `expected_scope=` and `expected_file_hash=` keyword args, which do not exist on the manager method shown in `token_manager.py`.   

### Remediation
- Update `xls_delete_sheet` (and any similar tools) to:
  - `mgr.validate_token(args.token, scope="sheet:delete", file_path=<path-being-edited>)`.   
- Add a unit test that executes the tool’s `_run()` logic up to token validation to prevent regression.

---

# Phase 1 — Unified “edit target” semantics across all tools (1–2 weeks)

This phase makes the system coherent: **every tool knows what file it is actually modifying**, and governance/audit locks onto that same target.

## 1.1 Introduce a single “edit session” abstraction and migrate tools
Create something like:

- `excel_agent/core/edit_session.py`
  - `prepare_edit_target(input_path, output_path, *, force_inplace) -> Path`
    - if `output_path` is provided and differs: copy input→output and return output
    - else: require `force_inplace`, return input
  - `open_edit_agent(edit_path) -> ExcelAgent`

Then mechanically update all mutating tools:

- Pattern becomes:
  1) resolve `edit_path`
  2) `with ExcelAgent(edit_path, mode="rw") as agent: ...`

This directly eliminates the double-save and “wrong file saved on exit” issues caused by `ExcelAgent.__exit__()` saving to its opened path.   

## 1.2 Normalize macro preservation behavior (crucial for Phase 3)
### What’s risky today (validated)
Some tools modify workbooks by calling `openpyxl.load_workbook(str(input_path))` directly (not via `ExcelAgent`) and then `wb.save(output_path)`. Example: `xls_add_chart`.   

Since `ExcelAgent` explicitly auto-detects macro preservation from extension (`.xlsm` / `.xltm`) via `self._keep_vba = self._path.suffix.lower() in _VBA_EXTENSIONS`, bypassing it risks inconsistent macro retention policies across the tool suite.   

### Remediation
- For every mutating tool that currently uses raw `load_workbook()`, migrate it to use the same edit-session + `ExcelAgent` path.
- For “create from template”, explicitly decide if `.xltm → .xlsm` should preserve macros (see Phase 3). Today it loads templates with `load_workbook(str(template_path))` without any macro-preservation handling.   

## 1.3 Tighten `validate_output_path()` to validate suffix + policy
Today `validate_output_path()` only checks parent directory existence/creation, not extension or overwrite policy.   

Add:
- `validate_output_suffix(path, allowed={.xlsx,.xlsm,.xltx,.xltm})`
- “macro contract” checks (Phase 3)
- optional `--overwrite-output` vs “fail if output exists” (recommended default: fail unless explicit overwrite)

---

# Phase 2 — Close the “lack of pivot table creation” gap (pragmatic + reliable) (2–4 weeks)

## Why “true pivot creation” is hard in this stack
OpenPyXL’s own documentation states it **preserves** pivot tables in existing files but is *not intended* for client code to create pivot tables from scratch.   

So the most reliable way to deliver pivot tables headlessly is:

1) **use a human-authored Excel template that already contains the pivot table(s)**, and  
2) have the agent only update the source data + ensure pivots refresh when opened in Excel.

That meets your stated workflow: agent does the tedious layout + data work; human opens in real Excel for viewing/refresh/customization.

## 2.1 Ship a “pivot-ready template workflow” as first-class
Deliverables:
- `docs/pivots.md`: canonical workflow
- `templates/` sample files:
  - a template with:
    - a “Data” sheet with an Excel Table (ListObject) named `DataTable`
    - a “Pivot” sheet with pivot table + pivot chart built against that table
    - optional slicers
- Provide a recommended sequence:
  1) `xls-create-from-template`
  2) `xls-write-range` into the table area
  3) (new) `xls-resize-table` (see below)
  4) (new) `xls-set-pivot-refresh-on-open`
  5) handoff to Excel user (they open → pivots refresh)

## 2.2 Add two missing tools that make pivot templates actually work in agent pipelines

### Tool A: `xls-resize-table` (or `xls-append-to-table`)
Pivot templates are only robust if the source table expands correctly when new data arrives. The repo already has `xls_add_table` but no “resize/append” primitive in the objects tool list.   

- `xls-resize-table --name DataTable --new-range Data!A1:H5000`
- or `xls-append-to-table --name DataTable --rows-json [...]` (and it auto-resizes)

### Tool B: `xls-set-pivot-refresh-on-open`
Even with a template pivot, you want Excel to refresh on open. This tool can be implemented as **OOXML zip patching**, similar in spirit to how `xls_inject_vba_project` already edits the workbook zip structure (it writes `xl/vbaProject.bin` into the archive).   

Implementation approach:
- open `.xlsx/.xlsm` as zip
- locate `xl/pivotTables/pivotTable*.xml`
- set `refreshOnLoad="1"` (and any related cache refresh flags you standardize on)
- write out to the edit target (using Phase 1 edit-session semantics)

This avoids depending on OpenPyXL to *create* pivots while still enabling “agent-updated data → Excel refreshes pivots”.

## 2.3 Acceptance criteria for “pivot gap closed”
- Starting from a provided template containing a pivot table:
  - agent fills new data,
  - agent resizes the source table,
  - agent marks pivots refresh-on-open,
  - human opens in Microsoft Excel and sees correct updated pivot with a single refresh (ideally automatic).

---

# Phase 3 — Close the “.xlsx deliverable vs macros needed” mismatch (2–6 weeks)

This phase turns an implicit, confusing situation into an explicit supported contract.

## 3.1 Make file format a deliberate, enforced choice
### What’s true technically (validated)
- `ExcelAgent` only auto-preserves VBA projects for `.xlsm` / `.xltm` extensions.   
- Macro injection defaults to producing an `.xlsm` output when the input is `.xlsx` (string replace).   

### Remediation
Introduce a repo-wide policy:

- If macros are required, the **deliverable must be `.xlsm`**.
- If the deliverable must be `.xlsx`, then macros must be absent, and tools that inject macros must reject `.xlsx` output.

Concretely:
- Update `xls_inject_vba_project`:
  - If `output_path.suffix != ".xlsm"` (or `.xltm` for templates), fail validation with guidance.
  - Also require `--force` if overwriting existing output (Phase 1 overwrite policy).   

## 3.2 Add a “strip macros” tool to support dual outputs
Many orgs want:
- an internal `.xlsm` for automation, and
- a macro-free `.xlsx` for distribution or easier opening under strict policies.

Add:
- `xls-strip-macros --input report.xlsm --output report.xlsx`

Implementation can mirror `xls_inject_vba_project`’s zip-based editing approach (but removing `xl/vbaProject.bin` + cleaning relationships/content types).   

## 3.3 Fix template creation to support macro-enabled templates properly
`xls_create_from_template` currently:
- accepts `.xltm` templates,
- loads them with `load_workbook(str(template_path))`,
- saves to an output path described as “(.xlsx)”.   

To align with the macro contract:
- If `template` is `.xltm` and output is `.xlsm`, the tool must preserve VBA (either by using `ExcelAgent` on a copied output, or by calling `load_workbook(..., keep_vba=True)`—but standardize on `ExcelAgent` so behavior is consistent with the rest of the suite).   

## 3.4 Document “macros are not required for charts; pivots are best via templates”
- Charts can be created without macros (repo already supports chart creation).   
- Pivot tables should be delivered via templates + refresh-on-open (Phase 2), because OpenPyXL creation isn’t a supported design target per its docs.   

---

# How these phases close your three gaps

## Gap 1: Output-path / input-mutation foot-gun
Closed by:
- editing the *edit target* (output copy) rather than the source input,
- removing “manual save-as inside ExcelAgent context” patterns,
- enforcing `--force` for in-place edits,
- adding regression tests.   

## Gap 2: Lack of pivot table creation
Closed (reliably) by:
- making pivot tables a **template-authored artifact**,
- adding table-resize/append + pivot-refresh-on-open tools,
- explicitly not promising “create pivots from scratch via OpenPyXL” given upstream docs.   

## Gap 3: Format mismatch between `.xlsx` deliverable and macros
Closed by:
- enforcing extension rules at tool boundaries (macro injection requires `.xlsm`),
- adding `xls-strip-macros` for macro-free deliverables,
- updating template creation so `.xltm → .xlsm` is a first-class supported path.   

---

## One additional “must-fix” discovered while validating (governance correctness)
While doing this refactor, fix the token-validation mismatch in tools like `xls_delete_sheet`, which currently calls `ApprovalTokenManager.validate_token()` with parameters that don’t match the manager’s implemented signature.   

Even if not directly part of the three gaps, it’s tightly coupled to “safe mutation” and will otherwise break destructive operations during remediation.

--- 

If you want to execute this as a tracked implementation plan, the cleanest work breakdown is: **(A) introduce edit-session + CLI enforcement + tests**, then **(B) migrate tools category-by-category (write/structure/cells/formulas/objects/formatting)**, then **(C) pivot-template toolchain**, then **(D) macro contract + strip tool + template macro preservation**.
