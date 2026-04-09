"""Formatting tools for excel-agent-tools.

This module provides 5 CLI tools for formatting Excel workbooks:
- xls_format_range: Apply fonts, fills, borders, and alignment
- xls_set_column_width: Set fixed or auto-fit column widths
- xls_freeze_panes: Freeze rows/columns for scrolling
- xls_apply_conditional_formatting: Add ColorScale, DataBar, IconSet rules
- xls_set_number_format: Apply currency, percentage, date formats

All operations are additive (non-destructive) and do not require governance tokens.
"""

from __future__ import annotations

__all__ = []
