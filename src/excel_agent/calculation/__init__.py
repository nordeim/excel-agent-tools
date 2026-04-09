"""
Calculation engines for excel-agent-tools.

Two-tier architecture:
- Tier 1: formulas library (in-process, 90.1% function coverage)
- Tier 2: LibreOffice headless (full-fidelity, external process)
"""

from __future__ import annotations

__all__: list[str] = []
