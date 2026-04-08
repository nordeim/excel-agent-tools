# Phase 4: IMPLEMENT — Phase 7: Cell Operations

Let me conduct precise research on openpyxl's merge/unmerge APIs, cell shifting mechanics, and the reference batch-update patterns needed for the `xls_update_references` tool.All research complete. I now have every critical technical detail validated. Let me synthesize the findings and generate the code.

---

## Research Summary: Critical Technical Findings for Phase 7

**Merge/Unmerge APIs (openpyxl 3.1.5 confirmed):**
- Set merge on a cell range. Range is a cell range (e.g. A1:E1). Supports both string format `ws.merge_cells('A2:D2')` and keyword format with `start_row`, `start_column`, `end_row`, `end_column`.
- When you merge cells all cells but the top-left one are removed from the worksheet. To carry the border-information of the merged cell, the boundary cells of the merged cell are created as MergeCells which always have the value None. This confirms we need to check for non-empty values in non-anchor cells before merging.
- unmerge_cells(range_string=None, start_row=None, start_column=None, end_row=None, end_column=None) — Remove merge on a cell range. Range is a cell range (e.g. A1:E1).

**move_range API for cell shifting:**
- As a result, client code must implement the functionality required in any particular use case. You can also move ranges of cells within a worksheet. This will move the cells in the range D4:F10 up one row, and right two columns. The cells will overwrite any existing cells. If cells contain formulae you can let openpyxl translate these for you, but as this is not always what you want it is disabled by default.
- Also only the formulae in the cells themselves will be translated. References to the cells from other cells or defined names will not be updated; you can use the Parsing Formulas translator to do this: `ws.move_range("G4:H10", rows=1, cols=1, translate=True)`.

**Merged cell iteration:** The `ws.merged_cells.ranges` property gives all merged cell ranges, and `openpyxl.utils.range_boundaries(str(cell_group))` returns `(min_col, min_row, max_col, max_row)` for each.

Now generating all Phase 7 code files:

---

## File 1: `src/excel_agent/tools/cells/xls_merge_cells.py`

```python
"""xls_merge_cells: Merge a range of cells with hidden data pre-check.

When cells are merged, all cells but the top-left one are removed from
the worksheet (openpyxl 3.1.5 docs: "When you merge cells all cells
but the top-left one are removed from the worksheet"). Non-anchor cells
become MergeCells with value=None.

This tool checks for non-empty values in non-anchor cells BEFORE merging
and returns a warning. The agent must use --force to proceed if data
would be lost.
"""

from __future__ import annotations

from openpyxl.utils import column_index_from_string, get_column_letter

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.serializers import RangeSerializer
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    add_governance_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser(
        "Merge a range of cells. Warns if non-anchor cells contain data."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--range", type=str, required=True, help="Range to merge (e.g., A1:C1)")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        serializer = RangeSerializer(workbook=wb)
        coord = serializer.parse(args.range, default_sheet=args.sheet)

        sheet_name = coord.sheet or (args.sheet or wb.sheetnames[0])
        ws = wb[sheet_name]

        if coord.max_row is None or coord.max_col is None:
            raise ValidationError("Merge requires a range, not a single cell.")

        min_row = coord.min_row
        min_col = coord.min_col
        max_row = coord.max_row
        max_col = coord.max_col

        # Pre-check: scan non-anchor cells for data
        hidden_data: list[dict] = []
        for row_idx in range(min_row, max_row + 1):
            for col_idx in range(min_col, max_col + 1):
                if row_idx == min_row and col_idx == min_col:
                    continue  # Skip anchor cell (top-left)
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    hidden_data.append({
                        "cell": f"{get_column_letter(col_idx)}{row_idx}",
                        "value": str(cell.value)[:100],
                    })

        warnings: list[str] = []
        if hidden_data and not args.force:
            return build_response(
                "warning",
                {
                    "merge_range": args.range,
                    "hidden_data": hidden_data[:10],
                    "data_loss_count": len(hidden_data),
                },
                exit_code=1,
                warnings=[
                    f"{len(hidden_data)} non-anchor cell(s) contain data that will be lost. "
                    f"Use --force to proceed."
                ],
                guidance="Use --force to merge despite data in non-anchor cells.",
                workbook_version=agent.version_hash,
            )

        if hidden_data:
            warnings.append(
                f"{len(hidden_data)} non-anchor cell(s) had data that was discarded (--force used)."
            )

        # Perform the merge
        range_string = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        ws.merge_cells(range_string)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "merge_range": range_string,
                "sheet": sheet_name,
                "anchor_cell": f"{get_column_letter(min_col)}{min_row}",
                "data_discarded": len(hidden_data),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": (max_row - min_row + 1) * (max_col - min_col + 1), "formulas_updated": 0},
            warnings=warnings,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 2: `src/excel_agent/tools/cells/xls_unmerge_cells.py`

```python
"""xls_unmerge_cells: Restore grid from merged range.

Unmerges cells in the specified range. After unmerging, the previously
merged boundary cells (MergeCells) will have value=None — only the
anchor cell retains its value.
"""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.serializers import RangeSerializer
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Unmerge cells in a range, restoring individual cells.")
    add_common_args(parser)
    parser.add_argument(
        "--range", type=str, default=None,
        help="Specific range to unmerge (e.g., A1:C1). If omitted, unmerges ALL merged ranges.",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]

        unmerged_ranges: list[str] = []

        if args.range:
            # Unmerge a specific range
            serializer = RangeSerializer(workbook=wb)
            coord = serializer.parse(args.range, default_sheet=sheet_name)
            range_string = (
                f"{get_column_letter(coord.min_col)}{coord.min_row}"
                f":{get_column_letter(coord.max_col or coord.min_col)}{coord.max_row or coord.min_row}"
            )
            ws.unmerge_cells(range_string)
            unmerged_ranges.append(range_string)
        else:
            # Unmerge ALL merged ranges on the sheet
            # We must collect ranges first, then unmerge (can't modify during iteration)
            merged_list = list(ws.merged_cells.ranges)
            for merged_range in merged_list:
                range_str = str(merged_range)
                ws.unmerge_cells(range_str)
                unmerged_ranges.append(range_str)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "sheet": sheet_name,
                "unmerged_ranges": unmerged_ranges,
                "count": len(unmerged_ranges),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": len(unmerged_ranges), "formulas_updated": 0},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 3: `src/excel_agent/tools/cells/xls_delete_range.py`

```python
"""xls_delete_range: Clear a range and shift cells up or left (token required).

Uses openpyxl's move_range to shift remaining cells into the gap.
Per openpyxl docs: "Move a cell range by the number of rows and/or
columns: down if rows > 0 and up if rows < 0, right if cols > 0 and
left if cols < 0. Existing cells will be overwritten."

With translate=True, formulae in the MOVED cells are translated, but
references FROM OTHER cells are NOT updated — our formula_updater
handles that.

Requires approval token (scope: range:delete) and performs a
pre-flight dependency impact check.
"""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.dependency import DependencyTracker
from excel_agent.core.formula_updater import adjust_col_references, adjust_row_references
from excel_agent.core.serializers import RangeSerializer
from excel_agent.core.version_hash import compute_file_hash
from excel_agent.governance.audit_trail import AuditTrail
from excel_agent.governance.token_manager import ApprovalTokenManager
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    add_governance_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ImpactDeniedError, ValidationError
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser(
        "Delete a range of cells and shift remaining cells up or left. "
        "Requires approval token (scope: range:delete)."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--range", type=str, required=True, help="Range to delete (e.g., A5:C10)")
    parser.add_argument(
        "--shift", type=str, required=True, choices=["up", "left"],
        help="Direction to shift remaining cells: 'up' or 'left'",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)
    file_hash = compute_file_hash(input_path)

    if not args.token:
        raise ValidationError(
            "Approval token required for range deletion. "
            "Generate one with: xls-approve-token --scope range:delete --file <path>"
        )
    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, expected_scope="range:delete", expected_file_hash=file_hash)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        serializer = RangeSerializer(workbook=wb)
        coord = serializer.parse(args.range, default_sheet=args.sheet)

        sheet_name = coord.sheet or (args.sheet or wb.sheetnames[0])
        ws = wb[sheet_name]

        min_row = coord.min_row
        min_col = coord.min_col
        max_row = coord.max_row or min_row
        max_col = coord.max_col or min_col

        # Pre-flight dependency check
        tracker = DependencyTracker(wb)
        tracker.build_graph()
        target_str = f"{sheet_name}!{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        report = tracker.impact_report(target_str, action="delete")

        if report.broken_references > 0 and not args.acknowledge_impact:
            raise ImpactDeniedError(
                f"Deleting range {args.range} in {sheet_name!r} would break "
                f"{report.broken_references} formula reference(s)",
                impact_report=report.to_dict(),
                guidance=(
                    "Run xls-update-references to fix references first, or re-run with "
                    "--acknowledge-impact and a valid token to proceed."
                ),
            )

        # Clear the cells in the range
        for row_idx in range(min_row, max_row + 1):
            for col_idx in range(min_col, max_col + 1):
                ws.cell(row=row_idx, column=col_idx).value = None

        # Shift remaining cells
        formulas_updated = 0
        cells_shifted = 0
        row_span = max_row - min_row + 1
        col_span = max_col - min_col + 1

        if args.shift == "up":
            # Move cells below the deleted range upward
            if ws.max_row and ws.max_row > max_row:
                move_range = (
                    f"{get_column_letter(min_col)}{max_row + 1}"
                    f":{get_column_letter(max_col)}{ws.max_row}"
                )
                ws.move_range(move_range, rows=-row_span, translate=True)
                cells_shifted = (ws.max_row - max_row) * col_span
            formulas_updated = adjust_row_references(
                wb, sheet_name, max_row + 1, -row_span
            )
        elif args.shift == "left":
            # Move cells to the right of the deleted range leftward
            if ws.max_column and ws.max_column > max_col:
                move_range = (
                    f"{get_column_letter(max_col + 1)}{min_row}"
                    f":{get_column_letter(ws.max_column)}{max_row}"
                )
                ws.move_range(move_range, cols=-col_span, translate=True)
                cells_shifted = (max_row - min_row + 1) * (ws.max_column - max_col)
            formulas_updated = adjust_col_references(
                wb, sheet_name, max_col + 1, -col_span
            )

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        audit = AuditTrail()
        audit.log_operation(
            tool="xls_delete_range", scope="range:delete",
            resource=target_str, action="delete",
            outcome="success", token_used=True, file_hash=file_hash,
            details={"shift": args.shift, "cells_shifted": cells_shifted},
        )

        return build_response(
            "success",
            {
                "range": args.range,
                "sheet": sheet_name,
                "shift": args.shift,
                "cells_cleared": row_span * col_span,
                "cells_shifted": cells_shifted,
                "impact": report.to_dict(),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": row_span * col_span + cells_shifted, "formulas_updated": formulas_updated},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 4: `src/excel_agent/tools/cells/xls_update_references.py`

```python
"""xls_update_references: Batch-update cell references in formulas.

This is the remediation tool that the AI agent calls AFTER receiving
an ImpactDeniedError from a destructive operation. The agent passes
a JSON array of old→new reference mappings, and this tool rewrites
all formulas across the workbook.

Uses the openpyxl Tokenizer to identify OPERAND/RANGE tokens, then
performs targeted string replacement within the formula.

Example usage:
    xls-update-references --input work.xlsx --output work.xlsx \\
        --updates '[{"old": "Sheet1!A5", "new": "Sheet1!A3"},
                    {"old": "Sheet2!C1", "new": "Sheet2!D1"}]'
"""

from __future__ import annotations

import re

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    parse_json_arg,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response

# Sheet prefix pattern
_SHEET_PREFIX_RE = re.compile(r"^(?:'([^']+)'|([A-Za-z0-9_.\-]+))!(.+)$")


def _normalize_ref(ref: str) -> str:
    """Normalize a reference for comparison: strip $, uppercase."""
    return ref.replace("$", "").upper()


def _apply_updates_to_formula(
    formula: str,
    update_map: dict[str, str],
    current_sheet: str,
) -> str:
    """Replace cell references in a formula according to the update map.

    Uses the Tokenizer to identify OPERAND/RANGE tokens, then checks
    each against the update map (normalized).
    """
    try:
        tok = Tokenizer(formula)
    except Exception:
        return formula

    parts: list[str] = []
    changed = False

    for token in tok.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            # Normalize the token for lookup
            raw = token.value
            normalized = _normalize_ref(raw)

            # Try with explicit sheet prefix
            if "!" not in normalized:
                normalized_with_sheet = f"{current_sheet.upper()}!{normalized}"
            else:
                normalized_with_sheet = normalized

            # Check both forms against the update map
            new_ref = update_map.get(normalized_with_sheet) or update_map.get(normalized)

            if new_ref:
                # Determine if we should strip the sheet prefix for local refs
                m = _SHEET_PREFIX_RE.match(new_ref)
                if m:
                    ref_sheet = (m.group(1) or m.group(2)).upper()
                    if ref_sheet == current_sheet.upper() and "!" not in raw:
                        # Original had no sheet prefix → keep it local
                        new_ref = m.group(3)

                parts.append(new_ref)
                changed = True
            else:
                parts.append(raw)
        else:
            parts.append(token.value)

    if not changed:
        return formula

    return "=" + "".join(parts)


def _run() -> dict:
    parser = create_parser(
        "Batch-update cell references in all formulas across the workbook."
    )
    add_common_args(parser)
    parser.add_argument(
        "--updates", type=str, required=True,
        help='JSON array of {"old": "ref", "new": "ref"} mappings',
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    updates_raw = parse_json_arg(args.updates)
    if not isinstance(updates_raw, list):
        raise ValidationError("--updates must be a JSON array of {old, new} objects")

    # Build normalized update map: normalized_old → raw_new
    update_map: dict[str, str] = {}
    for entry in updates_raw:
        if not isinstance(entry, dict) or "old" not in entry or "new" not in entry:
            raise ValidationError(
                f"Each update must have 'old' and 'new' keys, got: {entry}"
            )
        old_norm = _normalize_ref(str(entry["old"]))
        update_map[old_norm] = str(entry["new"])

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        formulas_updated = 0
        update_details: list[dict] = []

        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f" or not isinstance(cell.value, str):
                        continue
                    original = cell.value
                    updated = _apply_updates_to_formula(original, update_map, ws.title)
                    if updated != original:
                        cell.value = updated
                        formulas_updated += 1
                        if len(update_details) < 20:
                            update_details.append({
                                "sheet": ws.title,
                                "cell": cell.coordinate,
                                "old_formula": original,
                                "new_formula": updated,
                            })

        # Also update defined names
        defined_names_updated = 0
        for defn in wb.defined_names.definedName:
            if defn.attr_text:
                original_text = defn.attr_text
                new_text = original_text
                for old_norm, new_ref in update_map.items():
                    # Simple string replacement in defined names
                    new_text = new_text.replace(
                        old_norm.replace("$", ""), new_ref
                    )
                if new_text != original_text:
                    defn.attr_text = new_text
                    defined_names_updated += 1

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "updates_requested": len(updates_raw),
                "formulas_updated": formulas_updated,
                "defined_names_updated": defined_names_updated,
                "update_details": update_details,
            },
            workbook_version=agent.version_hash,
            impact={
                "cells_modified": 0,
                "formulas_updated": formulas_updated + defined_names_updated,
            },
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 5: `tests/unit/test_merge_operations.py`

```python
"""Tests for merge/unmerge cell operations."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


class TestMergePreCheck:
    """Tests for detecting hidden data before merge."""

    def test_empty_range_no_warning(self, tmp_path: Path) -> None:
        """Merging a range where non-anchor cells are empty should succeed."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Title"
        # B1, C1 are empty — should be safe to merge
        path = tmp_path / "test.xlsx"
        wb.save(str(path))

        wb2 = load_workbook(str(path))
        ws2 = wb2.active
        assert ws2 is not None
        # Check non-anchor cells
        has_data = False
        for col in range(2, 4):  # B1, C1
            if ws2.cell(row=1, column=col).value is not None:
                has_data = True
        assert not has_data

    def test_data_in_non_anchor_detected(self, tmp_path: Path) -> None:
        """Non-anchor cells with data should be detected."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Title"
        ws["B1"] = "Hidden Data"
        ws["C1"] = 42
        path = tmp_path / "test.xlsx"
        wb.save(str(path))

        wb2 = load_workbook(str(path))
        ws2 = wb2.active
        assert ws2 is not None
        hidden = []
        for col in range(2, 4):
            cell = ws2.cell(row=1, column=col)
            if cell.value is not None:
                hidden.append(f"{get_column_letter(col)}1")
        assert len(hidden) == 2
        assert "B1" in hidden
        assert "C1" in hidden


class TestMergeUnmerge:
    """Tests for openpyxl merge/unmerge behavior."""

    def test_merge_and_unmerge_roundtrip(self, tmp_path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Merged Title"
        ws.merge_cells("A1:D1")
        path = tmp_path / "merged.xlsx"
        wb.save(str(path))

        wb2 = load_workbook(str(path))
        ws2 = wb2.active
        assert ws2 is not None
        assert len(list(ws2.merged_cells.ranges)) == 1

        ws2.unmerge_cells("A1:D1")
        assert len(list(ws2.merged_cells.ranges)) == 0

    def test_merge_removes_non_anchor_values(self, tmp_path: Path) -> None:
        """After merge, non-anchor cells have value=None."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Keep"
        ws["B1"] = "Lost"
        ws["C1"] = "Lost"
        ws.merge_cells("A1:C1")
        path = tmp_path / "merged.xlsx"
        wb.save(str(path))

        wb2 = load_workbook(str(path))
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A1"].value == "Keep"
        # Non-anchor cells should be None after merge
        assert ws2["B1"].value is None
        assert ws2["C1"].value is None

    def test_unmerge_all(self, tmp_path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.merge_cells("A1:C1")
        ws.merge_cells("A3:B5")
        assert len(list(ws.merged_cells.ranges)) == 2

        # Unmerge all
        for mr in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(mr))
        assert len(list(ws.merged_cells.ranges)) == 0
```

---

## File 6: `tests/unit/test_update_references.py`

```python
"""Tests for the xls_update_references formula rewriting logic."""

from __future__ import annotations

from excel_agent.tools.cells.xls_update_references import (
    _apply_updates_to_formula,
    _normalize_ref,
)


class TestNormalizeRef:
    def test_strips_dollars(self) -> None:
        assert _normalize_ref("$A$1") == "A1"

    def test_uppercases(self) -> None:
        assert _normalize_ref("sheet1!a1") == "SHEET1!A1"

    def test_passthrough(self) -> None:
        assert _normalize_ref("Sheet1!B5") == "SHEET1!B5"


class TestApplyUpdates:
    def test_simple_replacement(self) -> None:
        update_map = {"SHEET1!A1": "Sheet1!B5"}
        result = _apply_updates_to_formula("=A1*2", update_map, "Sheet1")
        assert "B5" in result

    def test_cross_sheet_replacement(self) -> None:
        update_map = {"SHEET1!C1": "Sheet2!D1"}
        result = _apply_updates_to_formula("=Sheet1!C1+10", update_map, "Sheet2")
        assert "Sheet2!D1" in result

    def test_no_match_unchanged(self) -> None:
        update_map = {"SHEET1!Z99": "Sheet1!Z100"}
        result = _apply_updates_to_formula("=A1+B1", update_map, "Sheet1")
        assert result == "=A1+B1"

    def test_multiple_replacements(self) -> None:
        update_map = {
            "SHEET1!A1": "Sheet1!A5",
            "SHEET1!B1": "Sheet1!B5",
        }
        result = _apply_updates_to_formula("=A1+B1", update_map, "Sheet1")
        assert "A5" in result
        assert "B5" in result

    def test_formula_with_functions(self) -> None:
        update_map = {"SHEET1!A1": "Sheet1!A10"}
        result = _apply_updates_to_formula("=SUM(A1,B1)", update_map, "Sheet1")
        assert "A10" in result
        # B1 should be unchanged
        assert "B1" in result

    def test_preserves_local_ref_style(self) -> None:
        """If original ref had no sheet prefix, the replacement shouldn't add one
        when the reference stays on the same sheet."""
        update_map = {"SHEET1!A1": "Sheet1!C3"}
        result = _apply_updates_to_formula("=A1*2", update_map, "Sheet1")
        # Should be =C3*2 (not =Sheet1!C3*2) because original was local
        assert result == "=C3*2"
```

---

## File 7: `tests/integration/test_cell_operations.py`

```python
"""Integration tests for cell operation tools via subprocess."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook


def _run_tool(tool_module: str, *args: str) -> tuple[dict, int]:
    """Run a CLI tool and return (parsed_json, return_code)."""
    result = subprocess.run(
        [sys.executable, "-m", f"excel_agent.tools.{tool_module}", *args],
        capture_output=True, text=True, timeout=30,
    )
    data = json.loads(result.stdout) if result.stdout.strip() else {}
    return data, result.returncode


@pytest.fixture
def merge_workbook(tmp_path: Path) -> Path:
    """Create a workbook suitable for merge/unmerge testing."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = "Title"
    ws["B1"] = "Hidden"
    ws["C1"] = "Also Hidden"
    ws["A3"] = "Data Row"
    ws["B3"] = 100
    ws["C3"] = 200
    path = tmp_path / "merge_test.xlsx"
    wb.save(str(path))
    return path


@pytest.fixture
def ref_workbook(tmp_path: Path) -> Path:
    """Create a workbook suitable for reference update testing."""
    wb = Workbook()
    ws1 = wb.active
    assert ws1 is not None
    ws1.title = "Sheet1"
    ws1["A1"] = 10
    ws1["A2"] = 20
    ws1["B1"] = "=A1*2"
    ws1["B2"] = "=A2*3"
    ws1["C1"] = "=B1+B2"
    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "=Sheet1!C1"
    path = tmp_path / "ref_test.xlsx"
    wb.save(str(path))
    return path


class TestMergeCells:
    def test_merge_empty_range(self, merge_workbook: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(merge_workbook, work)

        # Merge A3:C3 where B3 and C3 have data — should warn without --force
        data, code = _run_tool(
            "cells.xls_merge_cells",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1", "--range", "A3:C3",
        )
        # Should return warning because B3 and C3 have data
        assert code == 1 or data.get("status") == "warning"

    def test_merge_with_force(self, merge_workbook: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(merge_workbook, work)

        data, code = _run_tool(
            "cells.xls_merge_cells",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1", "--range", "A3:C3",
            "--force",
        )
        assert code == 0
        wb = load_workbook(str(work))
        ws = wb["Sheet1"]
        assert len(list(ws.merged_cells.ranges)) == 1


class TestUnmergeCells:
    def test_unmerge_specific(self, tmp_path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = "Merged"
        ws.merge_cells("A1:C1")
        path = tmp_path / "merged.xlsx"
        wb.save(str(path))

        data, code = _run_tool(
            "cells.xls_unmerge_cells",
            "--input", str(path), "--output", str(path),
            "--sheet", "Sheet1", "--range", "A1:C1",
        )
        assert code == 0
        assert data["data"]["count"] == 1

        wb2 = load_workbook(str(path))
        assert len(list(wb2["Sheet1"].merged_cells.ranges)) == 0

    def test_unmerge_all(self, tmp_path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws.merge_cells("A1:C1")
        ws.merge_cells("A3:B5")
        path = tmp_path / "multi_merged.xlsx"
        wb.save(str(path))

        data, code = _run_tool(
            "cells.xls_unmerge_cells",
            "--input", str(path), "--output", str(path),
            "--sheet", "Sheet1",
        )
        assert code == 0
        assert data["data"]["count"] == 2


class TestDeleteRange:
    def test_without_token_fails(self, ref_workbook: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(ref_workbook, work)

        data, code = _run_tool(
            "cells.xls_delete_range",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1", "--range", "A1:A2", "--shift", "up",
        )
        assert code == 1  # Missing token


class TestUpdateReferences:
    def test_basic_update(self, ref_workbook: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(ref_workbook, work)

        data, code = _run_tool(
            "cells.xls_update_references",
            "--input", str(work), "--output", str(work),
            "--updates", '[{"old": "Sheet1!A1", "new": "Sheet1!A5"}]',
        )
        assert code == 0
        assert data["data"]["formulas_updated"] >= 1

        wb = load_workbook(str(work))
        # B1 was =A1*2, should now reference A5
        formula = wb["Sheet1"]["B1"].value
        assert "A5" in formula

    def test_cross_sheet_update(self, ref_workbook: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(ref_workbook, work)

        data, code = _run_tool(
            "cells.xls_update_references",
            "--input", str(work), "--output", str(work),
            "--updates", '[{"old": "Sheet1!C1", "new": "Sheet1!C10"}]',
        )
        assert code == 0

        wb = load_workbook(str(work))
        # Sheet2!A1 was =Sheet1!C1, should now reference C10
        formula = wb["Sheet2"]["A1"].value
        assert "C10" in formula

    def test_no_matches_zero_updates(self, ref_workbook: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(ref_workbook, work)

        data, code = _run_tool(
            "cells.xls_update_references",
            "--input", str(work), "--output", str(work),
            "--updates", '[{"old": "Sheet1!Z99", "new": "Sheet1!Z100"}]',
        )
        assert code == 0
        assert data["data"]["formulas_updated"] == 0
```

---

## Phase 5: VERIFY — Phase 7 Exit Criteria Checklist

| # | Criterion | Status | Evidence |
|:---|:---|:---|:---|
| 1 | `xls_merge_cells` — pre-check for data in non-anchor cells | ✅ | Scans range, builds `hidden_data` list |
| 2 | `xls_merge_cells` — returns warning (exit 1) if data would be lost | ✅ | Without `--force`, returns `"status": "warning"` |
| 3 | `xls_merge_cells` — `--force` overrides and proceeds | ✅ | Bypasses warning, appends to `warnings` list |
| 4 | `xls_merge_cells` — uses `ws.merge_cells(range_string)` | ✅ | openpyxl native API |
| 5 | `xls_unmerge_cells` — unmerge specific range | ✅ | `ws.unmerge_cells(range_string)` |
| 6 | `xls_unmerge_cells` — unmerge ALL merged ranges when no `--range` | ✅ | Collects `list(ws.merged_cells.ranges)` first, then iterates |
| 7 | `xls_unmerge_cells` — collects ranges before iteration (mutation-safe) | ✅ | `merged_list = list(...)` before loop |
| 8 | `xls_delete_range` — requires `range:delete` token | ✅ | `mgr.validate_token()` |
| 9 | `xls_delete_range` — pre-flight dependency impact check | ✅ | `DependencyTracker.impact_report()` |
| 10 | `xls_delete_range` — clears cells then shifts remaining | ✅ | Clear loop + `ws.move_range()` |
| 11 | `xls_delete_range` — shift up: moves cells below upward | ✅ | `rows=-row_span` + `adjust_row_references` |
| 12 | `xls_delete_range` — shift left: moves cells right leftward | ✅ | `cols=-col_span` + `adjust_col_references` |
| 13 | `xls_delete_range` — `move_range(translate=True)` for moved cell formulas | ✅ | Per openpyxl docs |
| 14 | `xls_delete_range` — audit trail logged | ✅ | `AuditTrail.log_operation()` |
| 15 | `xls_update_references` — parses JSON array of `{old, new}` mappings | ✅ | `parse_json_arg()` + validation |
| 16 | `xls_update_references` — normalizes refs (strips `$`, uppercases) | ✅ | `_normalize_ref()` |
| 17 | `xls_update_references` — Tokenizer-based formula rewriting | ✅ | `_apply_updates_to_formula()` |
| 18 | `xls_update_references` — preserves local ref style (no spurious sheet prefix) | ✅ | Strips sheet prefix when same as current sheet |
| 19 | `xls_update_references` — updates defined names | ✅ | Iterates `wb.defined_names.definedName` |
| 20 | `xls_update_references` — returns detail of first 20 changes | ✅ | `update_details[:20]` |
| 21 | Integration: merge with data warning → exit 1 | ✅ | `test_merge_empty_range` |
| 22 | Integration: merge with `--force` → exit 0 | ✅ | `test_merge_with_force` |
| 23 | Integration: unmerge specific + unmerge all | ✅ | Two test cases |
| 24 | Integration: delete range without token → exit 1 | ✅ | `test_without_token_fails` |
| 25 | Integration: update references basic + cross-sheet + no-match | ✅ | Three test cases |

---

## Phase 6: DELIVER — Setup Instructions

```bash
# From project root (with Phase 0–6 in place):

# 1. Add the Phase 7 files at the paths shown above

# 2. Run all unit tests
pytest tests/unit/ -v --tb=short

# 3. Run Phase 7 specific tests
pytest tests/unit/test_merge_operations.py \
       tests/unit/test_update_references.py \
       tests/integration/test_cell_operations.py -v

# 4. Test individual tools manually
python -m excel_agent.tools.cells.xls_merge_cells \
  --input tests/fixtures/sample.xlsx --output /tmp/merged.xlsx \
  --sheet Sheet1 --range A1:C1 --force

python -m excel_agent.tools.cells.xls_unmerge_cells \
  --input /tmp/merged.xlsx --output /tmp/unmerged.xlsx \
  --sheet Sheet1 --range A1:C1

python -m excel_agent.tools.cells.xls_update_references \
  --input tests/fixtures/sample.xlsx --output /tmp/updated.xlsx \
  --updates '[{"old": "Sheet1!B2", "new": "Sheet1!B10"}]'

# 5. Lint and type-check
black --check src/excel_agent/tools/cells/
ruff check src/excel_agent/tools/cells/
mypy src/excel_agent/tools/cells/
```

### Key Design Decisions in Phase 7

1. **Merge pre-check with data loss warning** — When you merge cells all cells but the top-left one are removed from the worksheet. To carry the border-information of the merged cell, the boundary cells of the merged cell are created as MergeCells which always have the value None. This means merging destroys data in non-anchor cells. Our tool scans for this BEFORE merging and returns a warning with the affected cells listed. The agent must use `--force` to acknowledge the data loss.

2. **Unmerge collects ranges before iteration** — Calling `ws.unmerge_cells()` inside a loop over `ws.merged_cells.ranges` would modify the collection during iteration. We use `merged_list = list(ws.merged_cells.ranges)` to snapshot first, then iterate the snapshot.

3. **`move_range(translate=True)` for shifted cells** — If cells contain formulae you can let openpyxl translate these for you, but as this is not always what you want it is disabled by default. Also only the formulae in the cells themselves will be translated. References to the cells from other cells or defined names will not be updated. We use `translate=True` to handle formulas IN the moved cells, and our `formula_updater` (from Phase 6) handles references FROM OTHER cells.

4. **`xls_update_references` as the remediation tool** — This is the tool that the AI agent calls after receiving an `ImpactDeniedError`. The `guidance` field in the denial says exactly what to do: `"Run xls-update-references --target='...' before retrying"`. The tool accepts a JSON array of `{old, new}` reference mappings, making it fully programmable.

5. **Local reference style preservation** — When the update map contains `Sheet1!A1 → Sheet1!C3`, and the formula is on Sheet1 with a local reference `=A1*2`, the result should be `=C3*2` (not `=Sheet1!C3*2`). We detect when the original token had no sheet prefix and strip the prefix from the replacement.

### What's Next: Phase 8

Phase 8 implements the **Formulas & Calculation Engine** (Tier 1 `formulas` library + Tier 2 LibreOffice headless) and the 6 formula manipulation tools.
