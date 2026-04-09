I now have Phase 8 specification. Let me analyze it and create a plan:
Phase 8: Formulas & Calculation Engine
Goal: Implement 6 formula tools and 2 calculation engines (Tier 1 + Tier 2)
Tools to Create (6):
1. xls_set_formula.py - Set formula with syntax validation
2. xls_recalculate.py - Two-tier recalculation (formulas library → LibreOffice)
3. xls_detect_errors.py - Scan for #REF!, #VALUE!, #DIV/0!, etc.
4. xls_convert_to_values.py - Replace formulas with values (⚠️ token required)
5. xls_copy_formula_down.py - Auto-fill formula with reference adjustment
6. xls_define_name.py - Create/update named ranges
Calculation Engines (3 files):
1. src/excel_agent/calculation/__init__.py
2. src/excel_agent/calculation/tier1_engine.py - formulas library integration
3. src/excel_agent/calculation/tier2_libreoffice.py - LibreOffice headless
4. src/excel_agent/calculation/error_detector.py - Error scanning
Tests (3 files):
1. tests/unit/test_tier1_engine.py
2. tests/unit/test_formula_tools.py 
3. tests/integration/test_calculation.py
Key Technical Considerations:
1. Tier 1 limitation: formulas library calculates from file on disk, not in-memory
2. Auto-fallback: Tier 1 → Tier 2 when unsupported functions encountered
3. Per-process LibreOffice profile: Prevents lock conflicts
4. Translator class: For formula copying with reference adjustment
5. Dual load for convert-to-values: data_only=True for cached values
6. Syntax validation: Tokenizer-based validation for formulas
I'll methodically work through this phase, breaking down the implementation into systematic steps. The focus is on creating robust, efficient formula and calculation tools that leverage both in-memory and
I'll now implement Phase 8: Formulas & Calculation Engine. Let me start with the **Phase 1:
# Todos
Create calculation directory structure
Create calculation/__init__.py
Create tier1_engine.py - formulas library integration
Create tier2_libreoffice.py - LibreOffice headless wrapper
Create error_detector.py - error scanning
Create formulas directory structure
Create xls_set_formula.py - set formula with validation
Create xls_recalculate.py - two-tier recalculation
Create xls_detect_errors.py - error scanning tool
Create xls_convert_to_values.py - token-gated conversion
Create xls_copy_formula_down.py - formula auto-fill
Create xls_define_name.py - named ranges
Create formulas/__init__.py
Create tests/unit/test_tier1_engine.py
Create tests/unit/test_formula_tools.py
Create tests/integration/test_calculation.py
Create scripts/recalc.py
Run linting (black, ruff, mypy)
Run tests and verify all pass
Thinking: Let
# Create Phase 8 directories
$ mkdir -p /home/project/excel-tools/src/excel_agent/calculation /home/project/excel-tools/src/excel_agent/tools/formulas /home/project/excel-tools/tests/unit /home/project/excel-tools/tests/integration /home/project/excel-tools/scripts

I've implemented the core calculation engine and several formula tools. However, I need to pause and correct a pattern I've been following. Looking at the xls_recalculate.py file I just wrote, I've created a _run() function that returns a dictionary and then manually calls build_response. 
Wait... I should be using the run_tool wrapper from _tool_base.py. The wrapper is designed to handle exceptions, print the JSON output, and exit with the correct code. By manually calling build_response and then presumably printing it, I'm duplicating logic that the framework already provides. More importantly, I'm not using the run_tool decorator, which means the tool runner won't have access to the metadata it needs to properly execute the command.
Let me check the imports. Yes, I've imported run_tool but I'm not actually using it as the entry point for my _run function. The correct pattern, as seen in other tools like xls_merge_cells.py, is to define _run, call run_tool(_run) in the main() function, and let the wrapper handle the response serialization and exit codes.
Actually, I should have caught this sooner. It's a recurring theme in my recent file creations. If I don't fix this now, the tool discovery mechanism might fail, or the output might not be formatted according to the expected JSON schema, which would break integration with the agent. 
I'll now fix xls_recalculate.py by removing the manual build_response call and adjusting the flow so that run_tool manages the execution and output. This will bring the tool into compliance
Let me fix
# Wrote src/excel_agent/tools/formulas/xls_recalculate.py

