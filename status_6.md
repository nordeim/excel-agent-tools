The structure directory doesn't exist yet. Let me create the Phase 6 implementation plan:
Phase 6: Structural Mutation Tools - Implementation Plan
Based on the Phase 6 specification, I need to create:
Core Module (1 file)
- src/excel_agent/core/formula_updater.py - Formula reference updating engine
CLI Tools (8 files)
1. src/excel_agent/tools/structure/xls_add_sheet.py - Add sheet at position
2. src/excel_agent/tools/structure/xls_delete_sheet.py - Delete sheet with token + dependency check
3. src/excel_agent/tools/structure/xls_rename_sheet.py - Rename sheet + auto-update cross-sheet refs
4. src/excel_agent/tools/structure/xls_insert_rows.py - Insert rows with style inheritance
5. src/excel_agent/tools/structure/xls_delete_rows.py - Delete rows with token + impact report
6. src/excel_agent/tools/structure/xls_insert_columns.py - Insert columns
7. src/excel_agent/tools/structure/xls_delete_columns.py - Delete columns with token + impact report
8. src/excel_agent/tools/structure/xls_move_sheet.py - Reorder sheet position
Package Init (1 file)
- src/excel_agent/tools/structure/__init__.py
Test Files (2 files)
- tests/unit/test_formula_updater.py - Unit tests for formula_updater
- tests/integration/test_structure_tools.py - Integration tests
Total: 11 files to create
Let
I notice that Phase 6 requires creating the structure tools directory and implementing 11 files. Let me create a comprehensive implementation plan:
Phase 6: Structural Mutation Tools - Implementation Plan
Executive Summary
Implement 8 CLI tools for structural mutations (add/delete/rename/move sheets, insert/delete rows/columns) with the critical formula_updater.py module that handles formula reference adjustments since openpyxl does NOT manage formula dependencies on structural mutations.
Key Technical Challenge
Openpyxl explicitly states: "Openpyxl does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted" — this is considered out of scope. Therefore, we must implement our own formula reference updating using the Tokenizer.
Files to Create (11 total)
Category	File Path	Purpose
Core Module	src/excel_agent/core/formula_updater.py	Formula reference updating engine
Package Init	src/excel_agent/tools/structure/__init__.py	Structure package initialization
CLI Tool	src/excel_agent/tools/structure/xls_add_sheet.py	Add sheet at position
CLI Tool	src/excel_agent/tools/structure/xls_delete_sheet.py	Delete sheet (token + dependency check)
CLI Tool	src/excel_agent/tools/structure/xls_rename_sheet.py	Rename + auto-update cross-sheet refs
CLI Tool	src/excel_agent/tools/structure/xls_insert_rows.py	Insert rows with formula adjustment
CLI Tool	src/excel_agent/tools/structure/xls_delete_rows.py	Delete rows (token + impact report)
CLI Tool	src/excel_agent/tools/structure/xls_insert_columns.py	Insert columns with formula adjustment
CLI Tool	src/excel_agent/tools/structure/xls_delete_columns.py	Delete columns (token + impact report)
CLI Tool	src/excel_agent/tools/structure/xls_move_sheet.py	Reorder sheet position
Unit Tests	tests/unit/test_formula_updater.py	Formula updater unit tests
Integration Tests	tests/integration/test_structure_tools.py	Structure tools integration tests
Critical Implementation Notes
1. Formula Updater Architecture:
   - Uses openpyxl's Tokenizer to identify OPERAND/RANGE tokens
   - Sheet rename: Simple string replacement in formulas
   - Row insert/delete: Shift row numbers in cell references
   - Column insert/delete: Shift column letters in references
   - Returns #REF! for references pointing to deleted locations
2. Token-Gated Operations (require sheet:delete, sheet:rename, range:delete tokens):
   - xls_delete_sheet
   - xls_rename_sheet
   - xls_delete_rows
   - xls_delete_columns
3. Dependency Tracking:
   - Pre-flight impact reports via DependencyTracker
   - ImpactDeniedError raised if formulas would break
   - --acknowledge-impact flag to proceed anyway
4. Reference Adjustment Behavior:
   - Only shift references at/after the insertion/deletion point
   - Absolute references ($A$1) are NOT shifted (matching Excel behavior)
   - Cross-sheet references are updated across the entire workbook
Would you like me to proceed with implementing Phase 6?

Now I have a complete understanding of Phase 6 requirements. This is the most architecturally complex phase because:
1. Openpyxl limitation: Openpyxl does NOT manage formula dependencies on structural mutations (insert/delete rows/columns, sheet rename)
2. We must implement formula reference updating ourselves: This requires tokenizing formulas, identifying cell references, and shifting them
3. Critical file: formula_updater.py: This is the core module that handles all formula reference adjustments
Let me create a comprehensive implementation plan and then implement all the files.
Phase 6: Structural Mutation Tools - Implementation Plan
Executive Summary
Implement 8 CLI tools for structural mutations (add/delete/rename/move sheets, insert/delete rows/columns) with the critical formula_updater.py module that handles formula reference adjustments since openpyxl does NOT manage formula dependencies on structural mutations.
Key Technical Challenge
Openpyxl explicitly states: "Openpyxl does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted" — this is considered out of scope. Therefore, we must implement our own formula reference updating using the Tokenizer.
Files to Create (11 total)
| Category | File Path | Purpose |
|----------|-----------|---------|
| Core Module | src/excel_agent/core/formula_updater.py | Formula reference updating engine |
| Package Init | src/excel_agent/tools/structure/__init__.py | Structure package initialization |
| CLI Tool | src/excel_agent/tools/structure/xls_add_sheet.py | Add sheet at position |
| CLI Tool | src/excel_agent/tools/structure/xls_delete_sheet.py | Delete sheet (token + dependency check) |
| CLI Tool | src/excel_agent/tools/structure/xls_rename_sheet.py | Rename + auto-update cross-sheet refs |
| CLI Tool | src/excel_agent/tools/structure/xls_insert_rows.py | Insert rows with formula adjustment |
| CLI Tool | src/excel_agent/tools/structure/xls_delete_rows.py | Delete rows (token + impact report) |
| CLI Tool | src/excel_agent/tools/structure/xls_insert_columns.py | Insert columns with formula adjustment |
| CLI Tool | src/excel_agent/tools/structure/xls_delete_columns.py | Delete columns (token + impact report) |
| CLI Tool | src/excel_agent/tools/structure/xls_move_sheet.py | Reorder sheet position |
| Unit Tests | tests/unit/test_formula_updater.py | Formula updater unit tests |
| Integration Tests | tests/integration/test_structure_tools.py | Structure tools integration tests |
Critical Implementation Notes
1. Formula Updater Architecture:
   - Uses openpyxl's Tokenizer to identify OPERAND/RANGE tokens
   - Sheet rename: Simple string replacement in formulas
   - Row insert/delete: Shift row numbers in cell references
   - Column insert/delete: Shift column letters in references
   - Returns #REF! for references pointing to deleted locations
2. Token-Gated Operations (require sheet:delete, sheet:rename, range:delete tokens):
   - xls_delete_sheet
   - xls_rename_sheet
   - xls_delete_rows
   - xls_delete_columns
3. Dependency Tracking:
   - Pre-flight impact reports via DependencyTracker
   - ImpactDeniedError raised if formulas would break
   - --acknowledge-impact flag to proceed anyway
4. Reference Adjustment Behavior:
   - Only shift references at/after the insertion/deletion point
   - Absolute references ($A$1) are NOT shifted (matching Excel behavior)
   - Cross-sheet references are updated across the entire workbook
Let
Now I have a complete understanding of Phase 6. This is the most architecturally complex phase because openpyxl does NOT manage formula dependencies on structural mutations. Let
# Create structure tools directories
$ mkdir -p /home/project/excel-tools/src/excel_agent/tools/structure /home/project/excel-tools/tests/integration

---

Now let me provide a summary of Phase 6 implementation completion:
---
Phase 6: Structural Mutation Tools — COMPLETE
Summary
Successfully implemented all Phase 6 deliverables including the critical formula_updater.py module that handles formula reference adjustments when structural mutations occur (since openpyxl does NOT manage formula dependencies on structural changes).
Files Created (11 files)
| Category | File | Status |
|----------|------|--------|
| Core Module | src/excel_agent/core/formula_updater.py | ✅ Complete |
| Package Init | src/excel_agent/tools/structure/__init__.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_add_sheet.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_delete_sheet.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_rename_sheet.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_insert_rows.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_delete_rows.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_insert_columns.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_delete_columns.py | ✅ Complete |
| CLI Tool | src/excel_agent/tools/structure/xls_move_sheet.py | ✅ Complete |
| Unit Tests | tests/unit/test_formula_updater.py | ✅ 9 tests pass |
| Integration Tests | tests/integration/test_structure_tools.py | ✅ 7 tests pass |
Key Features Implemented
1. formula_updater.py — Central formula reference updating engine:
   - rename_sheet_in_formulas(): Updates cross-sheet references after sheet rename
   - adjust_row_references(): Shifts row numbers after row insert/delete
   - adjust_col_references(): Shifts column letters after column insert/delete
   - Generates #REF! for deleted references (matching Excel behavior)
   - Handles both absolute ($A$1) and relative references
   - Only shifts references at/after the insertion/deletion point
2. Token-Gated Operations (require approval tokens):
   - xls_delete_sheet — requires sheet:delete token + impact check
   - xls_rename_sheet — requires sheet:rename token
   - xls_delete_rows — requires range:delete token + impact check
   - xls_delete_columns — requires range:delete token + impact check
3. Impact Analysis & Audit Trail:
   - All destructive operations use DependencyTracker.impact_report()
   - ImpactDeniedError raised with prescriptive guidance when formulas would break
   - --acknowledge-impact flag to proceed despite warnings
   - Audit trail logging for all token-gated operations
Test Results
- Phase 6 Tests: 16 passed (9 unit + 7 integration)
- Full Test Suite: 214 passed (all tests in project)
- Formula Updater Tests: 9/9 passed
- Structure Tools Tests: 7/7 passed
Key Design Decisions
1. Critical Finding: Openpyxl explicitly states "does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted" — our formula_updater.py fills this gap.
2. Tokenizer-based Rewriting: Uses openpyxl's Tokenizer to identify OPERAND/RANGE tokens and adjust references individually.
3. Absolute vs Relative Handling: Absolute references ($A$1) with $ anchors are NOT shifted, matching Excel's native behavior.
4. Formula Update BEFORE Title Change: In xls_rename_sheet, formulas are updated before ws.title is changed, ensuring the old sheet name can still be found in formulas.
5. Cross-Sheet Impact: Formula adjustments propagate across ALL sheets in the workbook, not just the modified sheet.
Next Phase
Phase 7 implements the Cell Operations (4 tools): merge cells, unmerge cells, delete range
Now let me provide a final summary of Phase 6:
---
Phase 6: Structural Mutation Tools — COMPLETE ✅
Executive Summary
Successfully implemented 11 files including the critical formula_updater.py module that handles formula reference adjustments when structural mutations occur. This is essential because openpyxl does NOT manage formula dependencies on structural changes — this is explicitly documented as "out of scope" for openpyxl.
Implementation Summary
Component	Files	Status
Core Module	formula_updater.py	✅ Complete
Structure Tools	8 CLI tools	✅ Complete
Tests	Unit + Integration	✅ Complete
Test Results
- Full Test Suite: 214 passed (100%)
- New Phase 6 Tests: 16 passed (9 unit + 7 integration)
- Linting: All files pass black, minor ruff issues remaining (non-critical)
Key Capabilities
1. Formula Reference Updating:
   - Sheet rename → updates all cross-sheet references
   - Row insert/delete → shifts row references
   - Column insert/delete → shifts column references
   - Generates #REF! for deleted references
2. Governance:
   - Token-gated destructive operations
   - Pre-flight impact analysis
   - Prescriptive denial messages
   - Audit trail logging
3. Excel-Compatible Behavior:
   - Absolute references ($A$1) not shifted
   - Only references at/after modification point shifted
   - Cross-sheet references updated across entire workbook
Phase 6 Exit Criteria — ALL MET ✅
All 23 exit criteria from @Phase-6.md have been implemented and verified:
- ✅ Sheet rename updates cross-sheet refs
- ✅ Row/column insert/delete shifts references
- ✅ References before insertion point NOT shifted
- ✅ Deleted refs become #REF!
- ✅ All token-gated operations with impact checks
- ✅ Audit trail logging
- ✅ All tools tested and working
Manual CLI Verification
All tools are functional:
# Test add sheet
xls-add-sheet --input test.xlsx --output out.xlsx --name "NewSheet" --position 0
# Test move sheet
xls-move-sheet --input test.xlsx --output out.xlsx --name "NewSheet" --position 2
# Test insert rows (with formula adjustment)
xls-insert-rows --input test.xlsx --output out.xlsx --sheet Sheet1 --before-row 5 --count 3
Phase 6 is complete and ready for Phase 7!

