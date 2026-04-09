"""Formula and calculation tools for excel-agent-tools.

This module provides 6 CLI tools for formula manipulation:
- xls_set_formula: Set formula in a cell with syntax validation
- xls_recalculate: Two-tier recalculation (formulas → LibreOffice)
- xls_detect_errors: Scan for #REF!, #VALUE!, #DIV/0!, etc.
- xls_convert_to_values: Replace formulas with values (irreversible, token-gated)
- xls_copy_formula_down: Auto-fill formula with reference adjustment
- xls_define_name: Create/update named ranges

These tools use the formulas library (Tier 1) and LibreOffice (Tier 2)
for calculation, plus openpyxl's Translator for formula copying.
"""

from __future__ import annotations

__all__ = []
