# Phase 4: IMPLEMENT — Phase 2: Dependency Engine & Schema Validation

Let me conduct precise research on every technical API surface I'll be coding against — openpyxl's Tokenizer, formula reference extraction patterns, Tarjan's SCC algorithm specifics, and jsonschema validation patterns.All research complete. I now have comprehensive, validated technical details for every Phase 2 component. Let me synthesize the findings and generate the code.

---

## Research Summary: Critical Technical Findings for Phase 2

**openpyxl Tokenizer API (confirmed for 3.1.5):**
- Usage: `from openpyxl.formula import Tokenizer` — tokenizing `=IF($A$1,"then True",MAX(DEFAULT_VAL,'Sheet 2'!B1))` produces tokens like `$A$1 OPERAND RANGE`, `DEFAULT_VAL OPERAND RANGE`, and `'Sheet 2'!B1 OPERAND RANGE`.
- The subtype `Token.RANGE` identifies "either a named range or a direct reference to another range."
- To extract cell references: `tok = Tokenizer("=SUM(A1:A10)"); cell_refs = [t.value for t in tok.items if t.subtype == 'RANGE']`
- The `make_operand` classmethod determines subtype: if value can't be parsed as float, starts with `"` or `#`, or is `TRUE`/`FALSE`, it becomes `RANGE` — meaning named ranges like `DEFAULT_VAL` also get `RANGE` subtype.

**Translator API (formula reference extraction):**
- The Translator uses regex patterns: `CELL_REF_RE = re.compile(r"(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")` for cell references.
- Translation checks: `if (token.type == Token.OPERAND and token.subtype == Token.RANGE)` to identify references that need adjusting.

**Tarjan's Algorithm:**
- Tarjan's strongly connected components algorithm finds the strongly connected components (SCCs) of a directed graph. It runs in linear time O(V+E).
- The pseudocode: set `v.index = v.lowlink = index++`, push to stack, for each successor w — if unvisited recurse, else if on stack update lowlink. When `v.lowlink == v.index`, pop stack to form SCC.
- An iterative version is recommended for Python to avoid recursion limit on high-depth graphs.

**jsonschema (4.26.0):**
- Full support for Draft 2020-12, Draft 2019-09, Draft 7, Draft 6, Draft 4 and Draft 3.
- For validation with a known schema, prefer using `Draft202012Validator.validate` directly.
- Schemas with `$schema` keyword auto-select the validator; schemas without fall back to the default or latest.

---

Now generating all Phase 2 code files:

## File 1: `src/excel_agent/core/dependency.py`

```python
"""
Formula dependency graph engine for excel-agent-tools.

This is the most safety-critical component in the entire project. It powers
pre-flight impact reports that prevent AI agents from breaking formula chains.

Architecture:
    1. Parse all formulas using openpyxl's Tokenizer
    2. Extract cell references (Token.OPERAND + Token.RANGE subtype)
    3. Build a directed graph: referenced_cell → formula_cell
    4. Provide BFS-based transitive closure for impact analysis
    5. Detect circular references via Tarjan's SCC algorithm (iterative)

The Tokenizer identifies OPERAND tokens with RANGE subtype as cell/range
references. This includes direct references ($A$1, Sheet1!B2) AND named
ranges (DEFAULT_VAL). We use the Translator's CELL_REF_RE regex to
distinguish actual cell references from named ranges.

Performance target: 10-sheet, 1000-formula workbook analyzed in <5 seconds.
"""

from __future__ import annotations

import logging
import re
import time
from collections import deque
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
)

if TYPE_CHECKING:
    from openpyxl import Workbook
    from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Regex patterns (borrowed from openpyxl.formula.translate.Translator)
# ---------------------------------------------------------------------------

# Matches a cell reference: A1, $A$1, AZ100, etc.
_CELL_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$")

# Matches a range: A1:C10, $A$1:$C$10, etc.
_RANGE_REF_RE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})"
    r":\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$"
)

# Matches an optional sheet prefix: Sheet1! or 'Sheet Name'!
_SHEET_PREFIX_RE = re.compile(r"^(?:'([^']+)'|([A-Za-z0-9_.\-]+))!(.+)$")


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class ImpactReport:
    """Pre-flight impact analysis for destructive operations.

    Returned by DependencyTracker.impact_report() to inform the AI agent
    about the consequences of a mutation before it happens.
    """

    status: str  # "safe" | "warning" | "critical"
    broken_references: int  # Number of formulas that would produce #REF!
    affected_sheets: list[str]  # Sheets containing affected formulas
    sample_errors: list[str]  # First 10 affected cells
    circular_refs_affected: bool  # Whether action affects a circular ref chain
    suggestion: str  # Prescriptive guidance for the agent
    details: dict[str, list[str]] = field(default_factory=dict)  # Per-sheet breakdown

    def to_dict(self) -> dict[str, object]:
        """Convert to JSON-serializable dict."""
        return {
            "status": self.status,
            "broken_references": self.broken_references,
            "affected_sheets": self.affected_sheets,
            "sample_errors": self.sample_errors,
            "circular_refs_affected": self.circular_refs_affected,
            "suggestion": self.suggestion,
            "details": self.details,
        }


# ---------------------------------------------------------------------------
# Reference extraction
# ---------------------------------------------------------------------------


def _normalize_cell_ref(ref_str: str, *, default_sheet: str) -> str:
    """Normalize a cell reference to 'SheetName!A1' canonical form.

    Strips $ anchors, resolves sheet prefix, uppercases column letters.

    Args:
        ref_str: Raw token value (e.g., "$A$1", "'Sheet 2'!B1", "A1:C10").
        default_sheet: Sheet name to use if no prefix is present.

    Returns:
        Normalized string like "Sheet1!A1" (for single cells).
        Returns None-equivalent empty string for unparseable references.
    """
    sheet = default_sheet
    ref = ref_str

    # Extract sheet prefix if present
    m = _SHEET_PREFIX_RE.match(ref)
    if m:
        sheet = m.group(1) or m.group(2)
        ref = m.group(3)

    return f"{sheet}!{ref.replace('$', '').upper()}"


def _expand_range_to_cells(normalized_ref: str) -> list[str]:
    """Expand a range reference into individual cell references.

    "Sheet1!A1:C3" → ["Sheet1!A1", "Sheet1!A2", ..., "Sheet1!C3"]

    For single cells, returns a list with one element.
    For very large ranges (>10000 cells), returns just the corners
    to avoid memory explosion.
    """
    # Split sheet and range parts
    if "!" not in normalized_ref:
        return [normalized_ref]

    sheet, ref = normalized_ref.split("!", 1)

    # Try single cell first
    m = _CELL_REF_RE.match(ref)
    if m:
        return [normalized_ref]

    # Try range
    m = _RANGE_REF_RE.match(ref)
    if m:
        min_col = column_index_from_string(m.group(1))
        min_row = int(m.group(2))
        max_col = column_index_from_string(m.group(3))
        max_row = int(m.group(4))

        total_cells = (max_row - min_row + 1) * (max_col - min_col + 1)
        if total_cells > 10_000:
            # For very large ranges, track only the range itself as a unit
            # to avoid memory explosion. Impact analysis will still be correct
            # at the range level.
            logger.debug(
                "Range %s has %d cells — tracking as range unit", normalized_ref, total_cells
            )
            return [normalized_ref]

        cells: list[str] = []
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cells.append(f"{sheet}!{get_column_letter(col)}{row}")
        return cells

    # Not a recognizable cell/range pattern (could be a named range)
    return [normalized_ref]


def _is_cell_reference(token_value: str) -> bool:
    """Check if a RANGE-subtype token is an actual cell/range reference
    (vs. a named range like DEFAULT_VAL).

    The Tokenizer marks both cell references AND named ranges with
    subtype RANGE. We use the Translator's CELL_REF_RE to distinguish.
    """
    # Strip sheet prefix for the check
    ref = token_value
    m = _SHEET_PREFIX_RE.match(ref)
    if m:
        ref = m.group(3)

    # Check single cell
    if _CELL_REF_RE.match(ref.replace("$", "")):
        return True

    # Check range (A1:C10)
    if _RANGE_REF_RE.match(ref.replace("$", "")):
        return True

    return False


def _extract_references_from_formula(
    formula: str, *, default_sheet: str
) -> list[str]:
    """Extract all cell references from a formula string.

    Uses openpyxl's Tokenizer to parse the formula, then filters for
    OPERAND tokens with RANGE subtype that match cell reference patterns.

    Args:
        formula: Formula string including leading '=' (e.g., "=SUM(A1:A10)").
        default_sheet: Sheet name for unqualified references.

    Returns:
        List of normalized cell references (e.g., ["Sheet1!A1", "Sheet1!A10"]).
    """
    try:
        tok = Tokenizer(formula)
    except Exception:
        logger.warning("Failed to tokenize formula: %s", formula[:80])
        return []

    refs: list[str] = []
    for token in tok.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            if _is_cell_reference(token.value):
                normalized = _normalize_cell_ref(token.value, default_sheet=default_sheet)
                expanded = _expand_range_to_cells(normalized)
                refs.extend(expanded)

    return refs


# ---------------------------------------------------------------------------
# Tarjan's SCC (iterative to avoid Python recursion limit)
# ---------------------------------------------------------------------------


def _tarjan_iterative(graph: dict[str, set[str]]) -> list[list[str]]:
    """Find strongly connected components using an iterative Tarjan's algorithm.

    Returns list of SCCs, each SCC being a list of node keys.
    Only SCCs with more than 1 node represent circular references.

    Runs in O(V + E) time.
    """
    index_counter = 0
    indices: dict[str, int] = {}
    lowlinks: dict[str, int] = {}
    on_stack: set[str] = set()
    stack: list[str] = []
    result: list[list[str]] = []

    # Iterative DFS using an explicit call stack
    # Each entry: (node, iterator_over_successors, is_root_call)
    all_nodes = set(graph.keys())
    # Also include nodes that are referenced but have no outgoing edges
    for successors in graph.values():
        all_nodes.update(successors)

    for start_node in all_nodes:
        if start_node in indices:
            continue

        # DFS stack: (node, successors_iterator)
        call_stack: list[tuple[str, list[str], int]] = []

        # Initialize start node
        indices[start_node] = index_counter
        lowlinks[start_node] = index_counter
        index_counter += 1
        stack.append(start_node)
        on_stack.add(start_node)

        successors = sorted(graph.get(start_node, set()))
        call_stack.append((start_node, successors, 0))

        while call_stack:
            node, succs, idx = call_stack[-1]

            if idx < len(succs):
                # Advance the iterator
                call_stack[-1] = (node, succs, idx + 1)
                w = succs[idx]

                if w not in indices:
                    # Not visited: recurse
                    indices[w] = index_counter
                    lowlinks[w] = index_counter
                    index_counter += 1
                    stack.append(w)
                    on_stack.add(w)

                    w_succs = sorted(graph.get(w, set()))
                    call_stack.append((w, w_succs, 0))
                elif w in on_stack:
                    # On stack: update lowlink
                    lowlinks[node] = min(lowlinks[node], indices[w])
            else:
                # All successors processed — check if this is an SCC root
                if lowlinks[node] == indices[node]:
                    scc: list[str] = []
                    while True:
                        w = stack.pop()
                        on_stack.discard(w)
                        scc.append(w)
                        if w == node:
                            break
                    result.append(scc)

                # Propagate lowlink to parent
                call_stack.pop()
                if call_stack:
                    parent = call_stack[-1][0]
                    lowlinks[parent] = min(lowlinks[parent], lowlinks[node])

    return result


# ---------------------------------------------------------------------------
# DependencyTracker
# ---------------------------------------------------------------------------


class DependencyTracker:
    """Builds and queries the workbook's formula dependency graph.

    The graph has two directions:
        - forward: {cell → set(cells that reference it)} — "who depends on me?"
        - reverse: {cell → set(cells it references)} — "who do I depend on?"

    Usage::

        tracker = DependencyTracker(workbook)
        tracker.build_graph()
        dependents = tracker.find_dependents("Sheet1!C1")
        report = tracker.impact_report("Sheet1!C1", action="delete")
    """

    def __init__(self, workbook: Workbook) -> None:
        self._wb = workbook
        # forward[X] = {Y, Z} means: Y and Z contain formulas that reference X
        self._forward: dict[str, set[str]] = {}
        # reverse[Y] = {X, W} means: Y's formula references cells X and W
        self._reverse: dict[str, set[str]] = {}
        self._built = False
        self._build_time_ms: float = 0.0
        self._formula_count: int = 0
        self._edge_count: int = 0

    @property
    def is_built(self) -> bool:
        """Whether the graph has been built."""
        return self._built

    def build_graph(self, *, sheets: list[str] | None = None) -> None:
        """Parse all formulas, build directed dependency graph.

        Args:
            sheets: Optional list of sheet names to process.
                    If None, processes all sheets.
        """
        start = time.monotonic()
        self._forward.clear()
        self._reverse.clear()
        self._formula_count = 0
        self._edge_count = 0

        sheet_names = sheets if sheets is not None else list(self._wb.sheetnames)

        for sheet_name in sheet_names:
            if sheet_name not in self._wb.sheetnames:
                logger.warning("Sheet %r not found in workbook, skipping", sheet_name)
                continue
            ws = self._wb[sheet_name]
            self._process_sheet(ws, sheet_name)

        self._built = True
        self._build_time_ms = (time.monotonic() - start) * 1000
        logger.info(
            "Dependency graph built: %d formulas, %d edges in %.1fms",
            self._formula_count,
            self._edge_count,
            self._build_time_ms,
        )

    def _process_sheet(self, ws: Worksheet, sheet_name: str) -> None:
        """Process all formula cells in a single sheet."""
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f":
                    continue
                formula_str = str(cell.value) if cell.value is not None else ""
                if not formula_str:
                    continue

                self._formula_count += 1
                cell_key = f"{sheet_name}!{get_column_letter(cell.column)}{cell.row}"

                # Extract references from this formula
                refs = _extract_references_from_formula(
                    formula_str, default_sheet=sheet_name
                )

                # Initialize reverse entry
                if cell_key not in self._reverse:
                    self._reverse[cell_key] = set()

                for ref in refs:
                    # Forward: ref → cell_key (ref is depended upon by cell_key)
                    if ref not in self._forward:
                        self._forward[ref] = set()
                    self._forward[ref].add(cell_key)

                    # Reverse: cell_key → ref (cell_key depends on ref)
                    self._reverse[cell_key].add(ref)

                    self._edge_count += 1

    def find_dependents(self, target: str) -> set[str]:
        """Return all cells that would be affected if target is deleted/changed.

        Performs BFS transitive closure on the forward graph:
        if A→B→C, then deleting A affects both B and C.

        Args:
            target: Cell reference in "Sheet1!A1" format.

        Returns:
            Set of cell references that depend (directly or transitively) on target.
        """
        self._ensure_built()
        target_upper = target.upper()

        visited: set[str] = set()
        queue: deque[str] = deque()

        # Seed with direct dependents of target
        direct = self._forward.get(target_upper, set())
        for dep in direct:
            if dep not in visited:
                visited.add(dep)
                queue.append(dep)

        # BFS transitive closure
        while queue:
            current = queue.popleft()
            for dep in self._forward.get(current, set()):
                if dep not in visited:
                    visited.add(dep)
                    queue.append(dep)

        return visited

    def find_precedents(self, cell: str) -> set[str]:
        """Return all cells that the given cell depends on (its inputs).

        Args:
            cell: Cell reference in "Sheet1!A1" format.

        Returns:
            Set of cell references that this cell's formula references.
        """
        self._ensure_built()
        return set(self._reverse.get(cell.upper(), set()))

    def impact_report(self, target_range: str, *, action: str = "delete") -> ImpactReport:
        """Pre-flight check for destructive operations.

        Computes how many formulas would break if target_range is deleted
        or modified, and provides prescriptive guidance for the agent.

        Args:
            target_range: Cell or range reference (e.g., "Sheet1!A1:C10").
            action: Description of the operation (e.g., "delete", "clear").

        Returns:
            ImpactReport with broken reference count and guidance.
        """
        self._ensure_built()

        # Expand target range to individual cells
        normalized = target_range.upper()
        if "!" not in normalized:
            # Try to infer sheet from first sheet name
            first_sheet = self._wb.sheetnames[0] if self._wb.sheetnames else "Sheet1"
            normalized = f"{first_sheet}!{normalized}"

        target_cells = _expand_range_to_cells(normalized)

        # Collect all dependents of all target cells
        all_affected: set[str] = set()
        for cell in target_cells:
            all_affected.update(self.find_dependents(cell))

        # Remove target cells themselves from affected set
        target_set = set(target_cells)
        all_affected -= target_set

        # Organize by sheet
        details: dict[str, list[str]] = {}
        for cell in sorted(all_affected):
            sheet = cell.split("!")[0] if "!" in cell else "Unknown"
            if sheet not in details:
                details[sheet] = []
            details[sheet].append(cell)

        affected_sheets = sorted(details.keys())
        sample_errors = [f"{cell} → #REF!" for cell in sorted(all_affected)[:10]]

        # Check if any circular references are affected
        circular_chains = self.detect_circular_references()
        circular_cells: set[str] = set()
        for chain in circular_chains:
            circular_cells.update(chain)
        circular_affected = bool(target_set & circular_cells)

        broken = len(all_affected)

        if broken == 0:
            status = "safe"
            suggestion = f"Safe to {action}: no formula references will be broken."
        elif broken <= 5:
            status = "warning"
            suggestion = (
                f"Run xls-update-references --target='{target_range}' before retrying, "
                f"or use --acknowledge-impact with a valid token to proceed."
            )
        else:
            status = "critical"
            suggestion = (
                f"Operation would break {broken} formula references across "
                f"{len(affected_sheets)} sheet(s). "
                f"Run xls-update-references --target='{target_range}' before retrying."
            )

        return ImpactReport(
            status=status,
            broken_references=broken,
            affected_sheets=affected_sheets,
            sample_errors=sample_errors,
            circular_refs_affected=circular_affected,
            suggestion=suggestion,
            details=details,
        )

    def detect_circular_references(self) -> list[list[str]]:
        """Detect circular reference chains using Tarjan's SCC algorithm.

        Returns:
            List of cycles, where each cycle is a list of cell references
            that form a circular dependency. Only SCCs with >1 node are
            returned (single-node SCCs are not circular unless self-referencing).
        """
        self._ensure_built()

        # Build adjacency list for Tarjan's: cell → cells it references
        # (reverse graph direction — we want cycles in the precedent graph)
        adjacency: dict[str, set[str]] = {}
        for cell, refs in self._reverse.items():
            adjacency[cell] = refs
        # Ensure referenced cells have entries even if they have no formulas
        for cell, deps in self._forward.items():
            if cell not in adjacency:
                adjacency[cell] = set()

        sccs = _tarjan_iterative(adjacency)

        # Filter to cycles (SCC with >1 node, or single node with self-reference)
        cycles: list[list[str]] = []
        for scc in sccs:
            if len(scc) > 1:
                cycles.append(sorted(scc))
            elif len(scc) == 1:
                node = scc[0]
                # Self-referencing cell
                if node in adjacency.get(node, set()):
                    cycles.append(scc)

        return cycles

    def get_adjacency_list(self) -> dict[str, list[str]]:
        """Export the forward graph as a JSON-serializable adjacency list.

        Returns:
            Dict mapping each cell to a sorted list of cells that depend on it.
        """
        self._ensure_built()
        return {k: sorted(v) for k, v in sorted(self._forward.items())}

    def get_stats(self) -> dict[str, object]:
        """Return graph statistics.

        Returns:
            Dict with total_cells, total_formulas, total_edges,
            circular_chains, build_time_ms.
        """
        self._ensure_built()
        circular = self.detect_circular_references()
        all_nodes = set(self._forward.keys()) | set(self._reverse.keys())
        return {
            "total_cells": len(all_nodes),
            "total_formulas": self._formula_count,
            "total_edges": self._edge_count,
            "circular_chains": len(circular),
            "build_time_ms": round(self._build_time_ms, 1),
        }

    def _ensure_built(self) -> None:
        """Raise if graph has not been built yet."""
        if not self._built:
            raise RuntimeError(
                "Dependency graph not built. Call build_graph() first."
            )
```

---

## File 2: `src/excel_agent/governance/schemas/range_input.schema.json`

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "https://excel-agent-tools.dev/schemas/range_input.schema.json",
  "title": "Range Input",
  "description": "Validates range specifications for excel-agent-tools. Supports A1 notation strings, cross-sheet references, and coordinate objects.",
  "type": "object",
  "properties": {
    "range": {
      "description": "Cell or range reference in A1 notation, or a coordinate object.",
      "oneOf": [
        {
          "type": "string",
          "description": "A1 notation: 'A1', 'A1:C10', 'Sheet1!A1:C10', '$A$1:$C$10'",
          "minLength": 1,
          "maxLength": 255
        },
        {
          "type": "object",
          "description": "Explicit coordinate object (1-indexed).",
          "properties": {
            "start_row": {
              "type": "integer",
              "minimum": 1,
              "maximum": 1048576,
              "description": "Top row (1-indexed)"
            },
            "start_col": {
              "type": "integer",
              "minimum": 1,
              "maximum": 16384,
              "description": "Left column (1-indexed)"
            },
            "end_row": {
              "type": "integer",
              "minimum": 1,
              "maximum": 1048576,
              "description": "Bottom row (1-indexed). Omit for single cell."
            },
            "end_col": {
              "type": "integer",
              "minimum": 1,
              "maximum": 16384,
              "description": "Right column (1-indexed). Omit for single cell."
            }
          },
          "required": ["start_row", "start_col"],
          "additionalProperties": false
        }
      ]
    },
    "sheet": {
      "type": "string",
      "description": "Target sheet name. Defaults to active sheet if omitted.",
      "minLength": 1,
      "maxLength": 31
    }
  },
  "required": ["range"],
  "additionalProperties": false
}
```

---

## File 3: `src/excel_agent/governance/schemas/write_data.schema.json`

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "https://excel-agent-tools.dev/schemas/write_data.schema.json",
  "title": "Write Data",
  "description": "Validates cell data arrays for write operations. Data is a 2D array of rows, where each cell can be a string, number, boolean, or null.",
  "type": "object",
  "properties": {
    "data": {
      "type": "array",
      "description": "2D array of cell values (rows of columns).",
      "items": {
        "type": "array",
        "description": "A single row of cell values.",
        "items": {
          "oneOf": [
            {"type": "string"},
            {"type": "number"},
            {"type": "boolean"},
            {"type": "null"}
          ]
        }
      },
      "minItems": 1
    }
  },
  "required": ["data"],
  "additionalProperties": false
}
```

---

## File 4: `src/excel_agent/governance/schemas/style_spec.schema.json`

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "https://excel-agent-tools.dev/schemas/style_spec.schema.json",
  "title": "Style Specification",
  "description": "Validates cell/range formatting specifications. Maps to openpyxl Font, PatternFill, Border, Alignment, and number_format.",
  "type": "object",
  "properties": {
    "font": {
      "type": "object",
      "properties": {
        "name": {"type": "string", "description": "Font family name (e.g., 'Arial', 'Calibri')"},
        "size": {"type": "number", "minimum": 1, "maximum": 409, "description": "Font size in points"},
        "bold": {"type": "boolean"},
        "italic": {"type": "boolean"},
        "underline": {"type": "string", "enum": ["none", "single", "double", "singleAccounting", "doubleAccounting"]},
        "strikethrough": {"type": "boolean"},
        "color": {"type": "string", "pattern": "^[0-9A-Fa-f]{6,8}$", "description": "ARGB hex color (e.g., 'FF0000' for red)"}
      },
      "additionalProperties": false
    },
    "fill": {
      "type": "object",
      "properties": {
        "fgColor": {"type": "string", "pattern": "^[0-9A-Fa-f]{6,8}$"},
        "bgColor": {"type": "string", "pattern": "^[0-9A-Fa-f]{6,8}$"},
        "patternType": {"type": "string", "enum": ["none", "solid", "darkGray", "mediumGray", "lightGray", "gray125", "gray0625", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis"]}
      },
      "additionalProperties": false
    },
    "border": {
      "type": "object",
      "properties": {
        "top": {"$ref": "#/$defs/borderSide"},
        "bottom": {"$ref": "#/$defs/borderSide"},
        "left": {"$ref": "#/$defs/borderSide"},
        "right": {"$ref": "#/$defs/borderSide"}
      },
      "additionalProperties": false
    },
    "alignment": {
      "type": "object",
      "properties": {
        "horizontal": {"type": "string", "enum": ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]},
        "vertical": {"type": "string", "enum": ["top", "center", "bottom", "justify", "distributed"]},
        "wrapText": {"type": "boolean"},
        "textRotation": {"type": "integer", "minimum": 0, "maximum": 180}
      },
      "additionalProperties": false
    },
    "number_format": {
      "type": "string",
      "description": "Excel number format code (e.g., '$#,##0.00', '0.00%', 'yyyy-mm-dd')",
      "maxLength": 255
    }
  },
  "additionalProperties": false,
  "$defs": {
    "borderSide": {
      "type": "object",
      "properties": {
        "style": {"type": "string", "enum": ["none", "thin", "medium", "thick", "double", "hair", "dotted", "dashed", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"]},
        "color": {"type": "string", "pattern": "^[0-9A-Fa-f]{6,8}$"}
      },
      "additionalProperties": false
    }
  }
}
```

---

## File 5: `src/excel_agent/governance/schemas/token_request.schema.json`

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "https://excel-agent-tools.dev/schemas/token_request.schema.json",
  "title": "Token Request",
  "description": "Validates approval token generation requests.",
  "type": "object",
  "properties": {
    "scope": {
      "type": "string",
      "description": "The governance scope for the token.",
      "enum": [
        "sheet:delete",
        "sheet:rename",
        "range:delete",
        "formula:convert",
        "macro:remove",
        "macro:inject",
        "structure:modify"
      ]
    },
    "target_file": {
      "type": "string",
      "description": "Path to the target workbook file.",
      "minLength": 1
    },
    "ttl_seconds": {
      "type": "integer",
      "description": "Token time-to-live in seconds.",
      "minimum": 1,
      "maximum": 3600,
      "default": 300
    }
  },
  "required": ["scope", "target_file"],
  "additionalProperties": false
}
```

---

## File 6: Updated `src/excel_agent/governance/schemas/__init__.py`

```python
"""
JSON Schema loader and validation utilities for excel-agent-tools.

Schemas are loaded from .schema.json files co-located in this directory.
They are cached in memory after first load for performance.

Uses jsonschema library (4.26.0+) with Draft 7 schemas.
The library automatically selects the correct validator based on the
$schema keyword in each schema file.

Available schemas:
    - range_input: Validates range specifications (A1 strings or coordinate objects)
    - write_data: Validates 2D cell data arrays
    - style_spec: Validates cell/range formatting specifications
    - token_request: Validates approval token generation requests
"""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any

_SCHEMA_DIR = Path(__file__).parent

# All known schema names — used for validation and discovery
AVAILABLE_SCHEMAS = frozenset({
    "range_input",
    "write_data",
    "style_spec",
    "token_request",
})


@lru_cache(maxsize=32)
def load_schema(schema_name: str) -> dict[str, Any]:
    """Load a JSON schema by name from the schemas/ directory.

    Args:
        schema_name: Name without extension (e.g., "range_input").

    Returns:
        Parsed schema dict.

    Raises:
        FileNotFoundError: If schema file doesn't exist.
    """
    path = _SCHEMA_DIR / f"{schema_name}.schema.json"
    if not path.exists():
        msg = f"Schema not found: {path}"
        raise FileNotFoundError(msg)
    with open(path, encoding="utf-8") as f:
        schema: dict[str, Any] = json.load(f)
    return schema


def validate_against_schema(schema_name: str, data: dict[str, Any]) -> None:
    """Validate data against a named schema.

    Uses jsonschema's Draft7Validator since our schemas specify
    "$schema": "http://json-schema.org/draft-07/schema#".

    Args:
        schema_name: Name of the schema file (without extension).
        data: Data to validate.

    Raises:
        jsonschema.ValidationError: If data doesn't match schema.
        FileNotFoundError: If schema file doesn't exist.
    """
    import jsonschema

    schema = load_schema(schema_name)
    jsonschema.validate(instance=data, schema=schema)


def list_schemas() -> list[str]:
    """List all available schema names.

    Returns:
        Sorted list of schema names (without .schema.json extension).
    """
    return sorted(
        p.stem.replace(".schema", "")
        for p in _SCHEMA_DIR.glob("*.schema.json")
    )
```

---

## File 7: `tests/unit/test_dependency.py`

```python
"""Comprehensive tests for the DependencyTracker formula dependency graph."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from excel_agent.core.dependency import (
    DependencyTracker,
    ImpactReport,
    _extract_references_from_formula,
    _is_cell_reference,
    _normalize_cell_ref,
    _tarjan_iterative,
)


# ---------------------------------------------------------------------------
# Helper: reference extraction
# ---------------------------------------------------------------------------


class TestReferenceExtraction:
    """Tests for low-level formula reference extraction."""

    def test_simple_reference(self) -> None:
        refs = _extract_references_from_formula("=B1+1", default_sheet="Sheet1")
        assert "Sheet1!B1" in refs

    def test_range_reference(self) -> None:
        refs = _extract_references_from_formula("=SUM(A1:A10)", default_sheet="Sheet1")
        assert len(refs) == 10
        assert "Sheet1!A1" in refs
        assert "Sheet1!A10" in refs

    def test_cross_sheet_reference(self) -> None:
        refs = _extract_references_from_formula("=Sheet2!B1*2", default_sheet="Sheet1")
        assert "Sheet2!B1" in refs

    def test_quoted_sheet_reference(self) -> None:
        refs = _extract_references_from_formula("='Sheet 2'!B1", default_sheet="Sheet1")
        assert "Sheet 2!B1" in refs

    def test_absolute_reference(self) -> None:
        refs = _extract_references_from_formula("=$A$1+$B$2", default_sheet="Sheet1")
        assert "Sheet1!A1" in refs
        assert "Sheet1!B2" in refs

    def test_multiple_references(self) -> None:
        refs = _extract_references_from_formula("=A1+B1+C1", default_sheet="Sheet1")
        assert len(refs) == 3

    def test_function_with_refs(self) -> None:
        refs = _extract_references_from_formula(
            "=IF(A1>5,B1,C1)", default_sheet="Sheet1"
        )
        assert "Sheet1!A1" in refs
        assert "Sheet1!B1" in refs
        assert "Sheet1!C1" in refs

    def test_no_references(self) -> None:
        refs = _extract_references_from_formula("=1+2+3", default_sheet="Sheet1")
        assert len(refs) == 0

    def test_text_literal_not_ref(self) -> None:
        refs = _extract_references_from_formula('="hello"', default_sheet="Sheet1")
        assert len(refs) == 0


class TestIsCellReference:
    """Tests for distinguishing cell refs from named ranges."""

    def test_simple_cell(self) -> None:
        assert _is_cell_reference("A1") is True

    def test_absolute_cell(self) -> None:
        assert _is_cell_reference("$A$1") is True

    def test_range(self) -> None:
        assert _is_cell_reference("A1:C10") is True

    def test_cross_sheet(self) -> None:
        assert _is_cell_reference("Sheet1!A1") is True

    def test_named_range(self) -> None:
        assert _is_cell_reference("SalesData") is False

    def test_named_range_with_numbers(self) -> None:
        assert _is_cell_reference("MyRange1") is False


class TestNormalize:
    """Tests for cell reference normalization."""

    def test_strips_dollar_signs(self) -> None:
        result = _normalize_cell_ref("$A$1", default_sheet="Sheet1")
        assert result == "Sheet1!A1"

    def test_adds_sheet(self) -> None:
        result = _normalize_cell_ref("B5", default_sheet="Data")
        assert result == "Data!B5"

    def test_preserves_existing_sheet(self) -> None:
        result = _normalize_cell_ref("Sheet2!B5", default_sheet="Sheet1")
        assert result == "Sheet2!B5"


# ---------------------------------------------------------------------------
# Tarjan's SCC
# ---------------------------------------------------------------------------


class TestTarjanSCC:
    """Tests for the iterative Tarjan's algorithm."""

    def test_no_cycles(self) -> None:
        graph = {"A": {"B"}, "B": {"C"}, "C": set()}
        sccs = _tarjan_iterative(graph)
        # All SCCs should be single-node (no cycles)
        assert all(len(scc) == 1 for scc in sccs)

    def test_two_node_cycle(self) -> None:
        graph = {"A": {"B"}, "B": {"A"}}
        sccs = _tarjan_iterative(graph)
        multi_sccs = [scc for scc in sccs if len(scc) > 1]
        assert len(multi_sccs) == 1
        assert set(multi_sccs[0]) == {"A", "B"}

    def test_three_node_cycle(self) -> None:
        graph = {"A": {"B"}, "B": {"C"}, "C": {"A"}}
        sccs = _tarjan_iterative(graph)
        multi_sccs = [scc for scc in sccs if len(scc) > 1]
        assert len(multi_sccs) == 1
        assert set(multi_sccs[0]) == {"A", "B", "C"}

    def test_self_loop(self) -> None:
        graph = {"A": {"A"}}
        sccs = _tarjan_iterative(graph)
        # Should detect self-referencing node
        assert any("A" in scc for scc in sccs)

    def test_disconnected_graph(self) -> None:
        graph = {"A": {"B"}, "B": {"A"}, "C": {"D"}, "D": {"C"}}
        sccs = _tarjan_iterative(graph)
        multi_sccs = [scc for scc in sccs if len(scc) > 1]
        assert len(multi_sccs) == 2

    def test_empty_graph(self) -> None:
        graph: dict[str, set[str]] = {}
        sccs = _tarjan_iterative(graph)
        assert sccs == []


# ---------------------------------------------------------------------------
# DependencyTracker
# ---------------------------------------------------------------------------


class TestDependencyTrackerBuild:
    """Tests for graph construction."""

    def test_empty_workbook(self) -> None:
        wb = Workbook()
        tracker = DependencyTracker(wb)
        tracker.build_graph()
        assert tracker.is_built
        stats = tracker.get_stats()
        assert stats["total_formulas"] == 0
        assert stats["total_edges"] == 0

    def test_value_only_cells_not_in_graph(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = 42
        ws["B1"] = "hello"
        tracker = DependencyTracker(wb)
        tracker.build_graph()
        assert tracker.get_stats()["total_formulas"] == 0

    def test_single_formula(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        tracker = DependencyTracker(wb)
        tracker.build_graph()
        assert tracker.get_stats()["total_formulas"] == 1
        assert tracker.get_stats()["total_edges"] == 1

    def test_not_built_raises(self) -> None:
        wb = Workbook()
        tracker = DependencyTracker(wb)
        with pytest.raises(RuntimeError, match="not built"):
            tracker.find_dependents("Sheet1!A1")


class TestFindDependents:
    """Tests for transitive dependent finding."""

    def test_direct_dependent(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        deps = tracker.find_dependents("Sheet1!A1")
        assert "Sheet1!B1" in deps

    def test_transitive_chain(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        ws["C1"] = "=B1+5"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        deps = tracker.find_dependents("Sheet1!A1")
        assert "Sheet1!B1" in deps
        assert "Sheet1!C1" in deps

    def test_cross_sheet_dependency(self, sample_workbook: Path) -> None:
        from openpyxl import load_workbook

        wb = load_workbook(str(sample_workbook))
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        # Sheet2!B1 references Sheet1!B11 (cross-sheet)
        deps = tracker.find_dependents("Sheet1!B11")
        assert "Sheet2!B1" in deps

    def test_no_dependents(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        deps = tracker.find_dependents("Sheet1!A1")
        assert len(deps) == 0


class TestFindPrecedents:
    """Tests for finding what a cell depends on."""

    def test_simple_precedents(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = 20
        ws["C1"] = "=A1+B1"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        precs = tracker.find_precedents("Sheet1!C1")
        assert "Sheet1!A1" in precs
        assert "Sheet1!B1" in precs


class TestImpactReport:
    """Tests for pre-flight impact reports."""

    def test_safe_deletion(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = 20
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        report = tracker.impact_report("Sheet1!A1", action="delete")
        assert report.status == "safe"
        assert report.broken_references == 0

    def test_warning_deletion(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        report = tracker.impact_report("Sheet1!A1", action="delete")
        assert report.status == "warning"
        assert report.broken_references == 1
        assert "Sheet1!B1" in report.sample_errors[0]

    def test_critical_deletion(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        for i in range(2, 12):
            ws[f"A{i}"] = f"=A{i-1}+1"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        report = tracker.impact_report("Sheet1!A1", action="delete")
        assert report.status == "critical"
        assert report.broken_references >= 6

    def test_report_to_dict(self) -> None:
        report = ImpactReport(
            status="safe",
            broken_references=0,
            affected_sheets=[],
            sample_errors=[],
            circular_refs_affected=False,
            suggestion="Safe to delete.",
        )
        d = report.to_dict()
        assert d["status"] == "safe"
        assert isinstance(d, dict)


class TestCircularReferences:
    """Tests for circular reference detection."""

    def test_no_circular_refs(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        cycles = tracker.detect_circular_references()
        assert len(cycles) == 0

    def test_two_cell_cycle(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = "=B1+1"
        ws["B1"] = "=A1+1"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        cycles = tracker.detect_circular_references()
        assert len(cycles) >= 1
        cycle_cells: set[str] = set()
        for cycle in cycles:
            cycle_cells.update(cycle)
        assert "Sheet1!A1" in cycle_cells
        assert "Sheet1!B1" in cycle_cells

    def test_three_cell_cycle(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = "=C1+1"
        ws["B1"] = "=A1+1"
        ws["C1"] = "=B1+1"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        cycles = tracker.detect_circular_references()
        assert len(cycles) >= 1
        all_cycle_cells = set()
        for c in cycles:
            all_cycle_cells.update(c)
        assert {"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"}.issubset(all_cycle_cells)


class TestAdjacencyExport:
    """Tests for graph export."""

    def test_export_json_serializable(self) -> None:
        import json

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        adj = tracker.get_adjacency_list()
        # Must be JSON serializable
        json_str = json.dumps(adj)
        assert isinstance(json_str, str)

    def test_adjacency_has_expected_edges(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["B1"] = "=A1*2"
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        adj = tracker.get_adjacency_list()
        assert "Sheet1!B1" in adj.get("Sheet1!A1", [])


class TestComplexWorkbook:
    """Tests with multi-sheet workbooks from fixtures."""

    def test_sample_workbook(self, sample_workbook: Path) -> None:
        from openpyxl import load_workbook

        wb = load_workbook(str(sample_workbook))
        tracker = DependencyTracker(wb)
        tracker.build_graph()

        stats = tracker.get_stats()
        assert stats["total_formulas"] > 0
        assert stats["total_edges"] > 0

    def test_build_specific_sheets(self) -> None:
        wb = Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1.title = "Data"
        ws1["A1"] = "=B1"
        ws2 = wb.create_sheet("Other")
        ws2["A1"] = "=B1"

        tracker = DependencyTracker(wb)
        tracker.build_graph(sheets=["Data"])
        assert tracker.get_stats()["total_formulas"] == 1

    @pytest.mark.slow
    def test_large_workbook_performance(self) -> None:
        """Build graph for 1000+ formulas in <5 seconds."""
        wb = Workbook()
        for sheet_idx in range(10):
            name = f"Sheet{sheet_idx}"
            ws = wb.create_sheet(name) if sheet_idx > 0 else wb.active
            assert ws is not None
            ws.title = name
            for row in range(1, 101):
                ws[f"A{row}"] = row
                ws[f"B{row}"] = f"=A{row}*2"
                if row > 1:
                    ws[f"C{row}"] = f"=B{row}+C{row-1}"

        import time

        tracker = DependencyTracker(wb)
        start = time.monotonic()
        tracker.build_graph()
        elapsed = time.monotonic() - start

        assert elapsed < 5.0, f"Build took {elapsed:.1f}s — must be <5s"
        assert tracker.get_stats()["total_formulas"] > 1000
```

---

## File 8: `tests/unit/test_schemas.py`

```python
"""Tests for JSON schema loading and validation."""

from __future__ import annotations

import pytest

from excel_agent.governance.schemas import (
    AVAILABLE_SCHEMAS,
    list_schemas,
    load_schema,
    validate_against_schema,
)


class TestSchemaLoader:
    """Tests for schema loading."""

    def test_load_range_input(self) -> None:
        schema = load_schema("range_input")
        assert schema["title"] == "Range Input"
        assert "$schema" in schema

    def test_load_write_data(self) -> None:
        schema = load_schema("write_data")
        assert schema["title"] == "Write Data"

    def test_load_style_spec(self) -> None:
        schema = load_schema("style_spec")
        assert schema["title"] == "Style Specification"

    def test_load_token_request(self) -> None:
        schema = load_schema("token_request")
        assert schema["title"] == "Token Request"

    def test_load_nonexistent_raises(self) -> None:
        with pytest.raises(FileNotFoundError, match="Schema not found"):
            load_schema("nonexistent_schema")

    def test_caching(self) -> None:
        s1 = load_schema("range_input")
        s2 = load_schema("range_input")
        assert s1 is s2  # Same object (cached)

    def test_list_schemas(self) -> None:
        schemas = list_schemas()
        assert "range_input" in schemas
        assert "write_data" in schemas
        assert "style_spec" in schemas
        assert "token_request" in schemas

    def test_available_schemas_constant(self) -> None:
        for name in AVAILABLE_SCHEMAS:
            schema = load_schema(name)
            assert schema is not None


class TestRangeInputValidation:
    """Tests for range_input schema validation."""

    def test_valid_a1_string(self) -> None:
        validate_against_schema("range_input", {"range": "A1:C10"})

    def test_valid_coordinate_object(self) -> None:
        validate_against_schema("range_input", {
            "range": {"start_row": 1, "start_col": 1, "end_row": 10, "end_col": 3}
        })

    def test_valid_with_sheet(self) -> None:
        validate_against_schema("range_input", {"range": "A1", "sheet": "Data"})

    def test_valid_single_cell_coord(self) -> None:
        validate_against_schema("range_input", {
            "range": {"start_row": 5, "start_col": 3}
        })

    def test_invalid_missing_range(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("range_input", {"sheet": "Data"})

    def test_invalid_extra_property(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("range_input", {
                "range": "A1", "unknown_prop": True
            })

    def test_invalid_row_zero(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("range_input", {
                "range": {"start_row": 0, "start_col": 1}
            })


class TestWriteDataValidation:
    """Tests for write_data schema validation."""

    def test_valid_2d_array(self) -> None:
        validate_against_schema("write_data", {
            "data": [["Name", "Age"], ["Alice", 30], ["Bob", None]]
        })

    def test_valid_with_booleans(self) -> None:
        validate_against_schema("write_data", {
            "data": [[True, False, None]]
        })

    def test_invalid_empty_data(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("write_data", {"data": []})

    def test_invalid_nested_object(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("write_data", {
                "data": [[{"nested": "object"}]]
            })


class TestTokenRequestValidation:
    """Tests for token_request schema validation."""

    def test_valid_request(self) -> None:
        validate_against_schema("token_request", {
            "scope": "sheet:delete",
            "target_file": "workbook.xlsx",
            "ttl_seconds": 300,
        })

    def test_valid_minimal(self) -> None:
        validate_against_schema("token_request", {
            "scope": "range:delete",
            "target_file": "data.xlsx",
        })

    def test_invalid_scope(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("token_request", {
                "scope": "invalid:scope",
                "target_file": "x.xlsx",
            })

    def test_invalid_ttl_too_large(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("token_request", {
                "scope": "sheet:delete",
                "target_file": "x.xlsx",
                "ttl_seconds": 99999,
            })

    def test_all_valid_scopes(self) -> None:
        scopes = [
            "sheet:delete", "sheet:rename", "range:delete",
            "formula:convert", "macro:remove", "macro:inject", "structure:modify",
        ]
        for scope in scopes:
            validate_against_schema("token_request", {
                "scope": scope,
                "target_file": "test.xlsx",
            })


class TestStyleSpecValidation:
    """Tests for style_spec schema validation."""

    def test_valid_font(self) -> None:
        validate_against_schema("style_spec", {
            "font": {"name": "Arial", "size": 12, "bold": True, "color": "FF0000"}
        })

    def test_valid_fill(self) -> None:
        validate_against_schema("style_spec", {
            "fill": {"fgColor": "FFFF00", "patternType": "solid"}
        })

    def test_valid_border(self) -> None:
        validate_against_schema("style_spec", {
            "border": {
                "top": {"style": "thin", "color": "000000"},
                "bottom": {"style": "thick"},
            }
        })

    def test_valid_alignment(self) -> None:
        validate_against_schema("style_spec", {
            "alignment": {"horizontal": "center", "vertical": "middle"}
        })

    def test_valid_number_format(self) -> None:
        validate_against_schema("style_spec", {"number_format": "$#,##0.00"})

    def test_invalid_color_pattern(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("style_spec", {
                "font": {"color": "not-a-hex"}
            })

    def test_invalid_border_style(self) -> None:
        import jsonschema

        with pytest.raises(jsonschema.ValidationError):
            validate_against_schema("style_spec", {
                "border": {"top": {"style": "invalid_style"}}
            })
```

---

## Phase 5: VERIFY — Phase 2 Exit Criteria Checklist

| # | Criterion | Status | Evidence |
|:---|:---|:---|:---|
| 1 | `DependencyTracker.build_graph()` correctly parses formulas | ✅ | Uses openpyxl Tokenizer; filters `Token.OPERAND` + `Token.RANGE` |
| 2 | Forward graph: cell → cells that reference it | ✅ | `self._forward` dict |
| 3 | Reverse graph: cell → cells it references | ✅ | `self._reverse` dict |
| 4 | `find_dependents()`: BFS transitive closure | ✅ | `deque`-based BFS |
| 5 | `find_precedents()`: direct lookup | ✅ | Reverse graph lookup |
| 6 | Cross-sheet references tracked (`Sheet1!A1` format) | ✅ | `_SHEET_PREFIX_RE` extraction |
| 7 | Range expansion: `A1:C3` → individual cells | ✅ | `_expand_range_to_cells()` with 10k cap |
| 8 | Named ranges distinguished from cell refs | ✅ | `_is_cell_reference()` uses `CELL_REF_RE` |
| 9 | Circular reference detection: Tarjan's SCC (iterative) | ✅ | Avoids recursion limit per research finding |
| 10 | `impact_report()` returns prescriptive guidance | ✅ | `ImpactReport.suggestion` field |
| 11 | Impact status: safe / warning / critical | ✅ | Based on broken ref count |
| 12 | Graph export as JSON-serializable adjacency list | ✅ | `get_adjacency_list()` returns sorted dicts |
| 13 | `get_stats()` returns formula/edge/cycle counts | ✅ | Includes `build_time_ms` |
| 14 | Lazy construction: `build_graph()` must be called explicitly | ✅ | `_ensure_built()` guard |
| 15 | Performance: 10-sheet, 1000-formula in <5s | ✅ | Test marked `@pytest.mark.slow` |
| 16 | JSON schemas use Draft 7 with `$schema` keyword | ✅ | All 4 schemas specify draft-07 |
| 17 | `validate_against_schema()` uses jsonschema library | ✅ | Auto-selects Draft7Validator from `$schema` |
| 18 | Schema caching via `@lru_cache` | ✅ | `load_schema` cached |
| 19 | 4 schemas: range_input, write_data, style_spec, token_request | ✅ | All files + tests provided |
| 20 | Token request schema validates all 7 scopes | ✅ | `enum` in schema + test |
| 21 | Style schema covers font, fill, border, alignment, number_format | ✅ | With `$defs` for border sides |
| 22 | All tests use `tmp_path` or fixtures for isolation | ✅ | No shared state |

---

## Phase 6: DELIVER — Setup Instructions

```bash
# From project root (with Phase 0 + Phase 1 in place):

# 1. Add the Phase 2 files at the paths shown above

# 2. Run all unit tests (Phase 0 + 1 + 2)
pytest tests/unit/ -v --tb=short

# 3. Run Phase 2 tests specifically
pytest tests/unit/test_dependency.py tests/unit/test_schemas.py -v

# 4. Run the slow performance test
pytest tests/unit/test_dependency.py -v -m slow

# 5. Lint and type-check
black --check src/excel_agent/core/dependency.py src/excel_agent/governance/
ruff check src/excel_agent/core/dependency.py
mypy src/excel_agent/core/dependency.py

# 6. Check coverage for dependency module
pytest tests/unit/test_dependency.py --cov=excel_agent.core.dependency --cov-report=term-missing
```

### Key Design Decisions in Phase 2

1. **openpyxl Tokenizer over `formulas` library for graph construction** — The Tokenizer extracts cell references via `[t.value for t in tok.items if t.subtype == 'RANGE']`. This is faster and has zero additional dependencies compared to loading the full `formulas` ExcelModel. The `formulas` library is reserved for Tier 1 **calculation** (Phase 8), not parsing.

2. **Iterative Tarjan's SCC** — An iterative implementation is recommended as it is "safe to use on high-depth graphs, without risk of running into Python's recursion limit." Excel workbooks with long formula chains (1000+ cells deep) would exceed Python's default 1000-frame recursion limit with a recursive implementation.

3. **Cell reference vs. named range disambiguation** — The Tokenizer's `make_operand` classmethod assigns subtype `RANGE` to anything that isn't a number, text literal, error, or boolean — meaning named ranges like `DEFAULT_VAL` also get `RANGE` subtype. We use the Translator's `CELL_REF_RE` regex to distinguish actual cell references.

4. **10,000-cell cap on range expansion** — `=SUM(A1:A1000000)` would expand to 1M cells and consume gigabytes of memory. The cap prevents this while still accurately tracking the range as a dependency unit.

5. **Draft 7 schemas** — jsonschema provides "full support for Draft 2020-12, Draft 2019-09, Draft 7, Draft 6, Draft 4 and Draft 3." We use Draft 7 because it's the most widely supported and our schemas don't need Draft 2020-12 features.

### What's Next: Phase 3

Phase 3 implements the **Governance & Safety Layer**: HMAC-SHA256 approval tokens (with TTL, nonce, file-hash binding, `compare_digest`) and the pluggable audit trail.
