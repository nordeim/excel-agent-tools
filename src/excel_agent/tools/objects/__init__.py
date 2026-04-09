"""Object manipulation tools for excel-agent-tools.

This module provides 5 CLI tools for adding objects to Excel workbooks:
- xls_add_table: Convert range to Excel Table (ListObject)
- xls_add_chart: Create Bar, Line, Pie, Scatter charts
- xls_add_image: Insert images with aspect ratio preservation
- xls_add_comment: Add threaded comments to cells
- xls_set_data_validation: Configure dropdown lists and constraints

All operations are additive (non-destructive) and do not require governance tokens.
"""

from __future__ import annotations

__all__ = []
