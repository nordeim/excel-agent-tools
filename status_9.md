I have already successfully implemented the first 8 phases of the excel-agent-tools project, reaching a total of 45 out of 53 planned tools. The project is currently in a very stable and well-tested state, with 250 passing tests and zero linting errors.
Current Status
- Completed: Phases 0 through 8 (Project Scaffolding, Core Foundation, Dependency Engine, Governance, Write Tools, Structure Tools, Cell Operations, Formulas & Calculation Engine).
- Tools Completed: 45 out of 53.
- Test Status: All 250 existing tests pass; linting is clean.
- Next Step: Phase 9 - Macro Safety Tools.
Phase 9 Analysis
Goal: Implement 5 tools for macro safety analysis: xls_has_macros, xls_inspect_macros, xls_validate_macro_safety, xls_remove_macros, and xls_inject_vba_project.
Core Technology Stack:
- oletools (v0.60.2): The primary engine for VBA extraction and initial analysis.
- olefile (v0.46): Required for low-level VBA binary manipulation.
- Implementation Pattern: A MacroAnalyzer Protocol will be defined in core/macro_handler.py to provide a clean abstraction over the oletools library.
Security & Technical Considerations:
1. VBA Binary Format: VBA macros are stored as compressed binary streams within the .xlsm file. Direct injection or removal requires careful handling of these streams to avoid corrupting the workbook.
2. Security Risks: The analysis must specifically scan for AutoOpen/AutoExec triggers, Shell function calls, external network indicators (IOCs), and suspicious DDE links.
3. Governance: Operations that modify macro content (xls_remove_macros, xls_inject_vba_project) are highly sensitive and will require token-based approval, similar to the security model used for formula conversion in previous phases.
Deliverables:
I will now proceed with creating the necessary Python modules for this phase, including:
1. src/excel_agent/core/macro_handler.py: The MacroAnalyzer class utilizing oletools.
2. src/excel_agent/tools/macros/: Directory containing the 5 CLI tools.
3. Unit & Integration Tests: To verify functionality and security scanning logic.
Let's proceed with the implementation of
I'll now implement Phase 9: Macro Safety Tools. Let me first check the existing structure and then implement all 5
# Todos
Create macro_handler.py - MacroAnalyzer Protocol
Create macros directory structure
Create xls_has_macros.py - VBA presence check
Create xls_inspect_macros.py - List modules + signatures
Create xls_validate_macro_safety.py - Risk scan
Create xls_remove_macros.py - Strip VBA (double-token)
Create xls_inject_vba_project.py - Inject .bin (token-gated)
Create macros/__init__.py
Create tests/unit/test_macro_handler.py
Create tests/integration/test_macro_tools.py
Run linting and tests
Thinking: Let me check the current project structure and start implementing Phase 9: Macro Safety Tools. This phase implements 5 tools for macro detection, inspection, safety validation, removal, and injection - all powered by
# Ensure directories exist
$ mkdir -p /home/project/excel-tools/src/excel_agent/core /home/project/excel-tools/src/excel_agent/tools/macros /home/project/excel-tools/tests/unit /home/project/excel-tools/tests/integration

