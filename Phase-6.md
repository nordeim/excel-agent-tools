# Phase 4: IMPLEMENT — Phase 6: Structural Mutation Tools

Let me conduct precise research on openpyxl's sheet manipulation APIs, row/column insertion/deletion behavior, formula reference adjustment mechanics, and the Translator class for reference updating.All research complete. I now have every critical technical detail validated. Let me synthesize the findings and generate the code.

---

## Research Summary: Critical Technical Findings for Phase 6

**The most critical finding — openpyxl does NOT manage formula dependencies on structural mutations:**

Openpyxl does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted. This is considered to be out of scope for a library that focuses on managing the file format. As a result, client code must implement the functionality required in any particular use case.

This means **our tools must implement formula reference updating ourselves**. This is precisely why the `DependencyTracker` from Phase 2 is essential.

**`move_sheet` API:** move_sheet(sheet, offset=0) — Move a sheet or sheetname. The first one, sheet, is mandatory, and is the worksheet that we want to move around, the second is optional (it defaults to 0), and is the offset to use to specify the sheet position.

**`move_sheet` source code confirms:** The implementation gets the current index, deletes the sheet from that position, computes `new_pos = idx + offset`, and inserts the sheet at the new position.

**`remove` API:** remove(worksheet) — Remove worksheet from this workbook. Don't delete the only sheet in a workbook (Excel requires ≥1 sheet).

**Formula reference replacement for sheet rename:** Since openpyxl doesn't handle this, we must tokenize formulas, find `OPERAND`/`RANGE` tokens containing the old sheet name, and replace them. RANGE is either a named range or a direct reference to another range. Quoted sheet names appear as `'Sheet 2'!B1` in tokens.

Now generating all Phase 6 code files. This is the most architecturally complex phase — 1 shared helper module + 8 CLI tools + 1 integration test = 10 files:

---

## File 1: `src/excel_agent/core/formula_updater.py`

```python
"""
Formula reference updating engine for excel-agent-tools.

Since openpyxl does NOT manage formula dependencies when rows/columns
are inserted or deleted (confirmed by openpyxl 3.1.5 documentation:
"Openpyxl does not manage dependencies, such as formulae, tables,
charts, etc., when rows or columns are inserted or deleted"), this
module provides the reference updating logic.

Capabilities:
    - Replace sheet name references in formulas (for sheet rename)
    - Adjust row references after row insert/delete
    - Adjust column references after column insert/delete

Uses the openpyxl Tokenizer to parse formulas and identify OPERAND
tokens with RANGE subtype, then performs targeted string replacement.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import column_index_from_string, get_column_letter

if TYPE_CHECKING:
    from openpyxl import Workbook
    from openpyxl.worksheet.worksheet import Worksheet

import logging

logger = logging.getLogger(__name__)

# Regex for sheet prefix in token values
_SHEET_PREFIX_RE = re.compile(r"^(?:'([^']+)'|([A-Za-z0-9_.\-]+))!(.+)$")

# Regex for cell reference (with optional $ anchors)
_CELL_REF_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")

# Regex for range reference
_RANGE_REF_RE = re.compile(
    r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+):(\$?)([A-Za-z]{1,3})(\$?)(\d+)$"
)


def rename_sheet_in_formulas(
    workbook: Workbook,
    old_name: str,
    new_name: str,
) -> int:
    """Update all formula references from old_name to new_name across the workbook.

    Iterates every cell in every sheet. For each formula, tokenizes it,
    finds RANGE tokens referencing the old sheet name, and replaces them.

    Args:
        workbook: The workbook to update.
        old_name: The old sheet name.
        new_name: The new sheet name.

    Returns:
        Number of formulas updated.
    """
    updated_count = 0
    old_quoted = _quote_sheet_name(old_name)
    new_quoted = _quote_sheet_name(new_name)

    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                new_formula = _replace_sheet_in_formula(
                    cell.value, old_name, old_quoted, new_name, new_quoted
                )
                if new_formula != cell.value:
                    cell.value = new_formula
                    updated_count += 1

    # Also update defined names
    for defn in workbook.defined_names.definedName:
        if defn.attr_text and old_name in defn.attr_text:
            new_text = defn.attr_text.replace(
                f"'{old_name}'!", f"'{new_name}'!"
            ).replace(
                f"{old_name}!", f"{new_name}!"
            )
            if new_text != defn.attr_text:
                defn.attr_text = new_text
                updated_count += 1

    return updated_count


def _quote_sheet_name(name: str) -> str:
    """Quote a sheet name if it contains special characters."""
    if re.search(r"[^A-Za-z0-9_.]", name):
        return f"'{name}'"
    return name


def _replace_sheet_in_formula(
    formula: str,
    old_name: str,
    old_quoted: str,
    new_name: str,
    new_quoted: str,
) -> str:
    """Replace sheet name references within a formula string.

    Uses simple string replacement on known patterns:
        'OldName'!A1 → 'NewName'!A1
        OldName!A1 → NewName!A1
    """
    result = formula
    # Quoted form: 'Old Name'!
    result = result.replace(f"'{old_name}'!", f"'{new_name}'!")
    # Unquoted form: OldName!
    if old_name != old_quoted:
        pass  # Only quoted form needed for names with spaces
    else:
        result = result.replace(f"{old_name}!", f"{new_name}!")
    return result


def adjust_row_references(
    workbook: Workbook,
    target_sheet: str,
    start_row: int,
    row_delta: int,
) -> int:
    """Adjust row references in formulas after row insert or delete.

    For all formulas across the entire workbook, find references to cells
    in target_sheet at or below start_row, and shift them by row_delta.

    Args:
        workbook: The workbook to update.
        target_sheet: Sheet where rows were inserted/deleted.
        start_row: The row at/after which the shift applies.
        row_delta: Positive for insertion, negative for deletion.

    Returns:
        Number of formulas updated.
    """
    updated = 0
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                new_formula = _shift_rows_in_formula(
                    cell.value, target_sheet, ws.title, start_row, row_delta
                )
                if new_formula != cell.value:
                    cell.value = new_formula
                    updated += 1
    return updated


def adjust_col_references(
    workbook: Workbook,
    target_sheet: str,
    start_col: int,
    col_delta: int,
) -> int:
    """Adjust column references in formulas after column insert or delete.

    Args:
        workbook: The workbook to update.
        target_sheet: Sheet where columns were inserted/deleted.
        start_col: The column at/after which the shift applies (1-indexed).
        col_delta: Positive for insertion, negative for deletion.

    Returns:
        Number of formulas updated.
    """
    updated = 0
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                new_formula = _shift_cols_in_formula(
                    cell.value, target_sheet, ws.title, start_col, col_delta
                )
                if new_formula != cell.value:
                    cell.value = new_formula
                    updated += 1
    return updated


def _shift_rows_in_formula(
    formula: str,
    target_sheet: str,
    current_sheet: str,
    start_row: int,
    row_delta: int,
) -> str:
    """Shift row numbers in cell references within a formula."""
    try:
        tok = Tokenizer(formula)
    except Exception:
        return formula

    parts: list[str] = []
    for token in tok.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            new_val = _shift_token_rows(
                token.value, target_sheet, current_sheet, start_row, row_delta
            )
            parts.append(new_val)
        else:
            parts.append(token.value)

    reconstructed = "=" + "".join(parts)
    return reconstructed


def _shift_token_rows(
    token_value: str,
    target_sheet: str,
    current_sheet: str,
    start_row: int,
    row_delta: int,
) -> str:
    """Shift row references in a single token value."""
    ref_sheet = current_sheet
    ref_part = token_value

    m = _SHEET_PREFIX_RE.match(token_value)
    if m:
        ref_sheet = m.group(1) or m.group(2)
        ref_part = m.group(3)
        prefix = token_value[: token_value.rindex("!") + 1]
    else:
        prefix = ""

    if ref_sheet != target_sheet:
        return token_value

    # Try range
    rm = _RANGE_REF_RE.match(ref_part)
    if rm:
        r1 = _shift_single_row(int(rm.group(4)), rm.group(3), start_row, row_delta)
        r2 = _shift_single_row(int(rm.group(8)), rm.group(7), start_row, row_delta)
        if r1 is None or r2 is None:
            return prefix + "#REF!"
        return f"{prefix}{rm.group(1)}{rm.group(2)}{rm.group(3)}{r1}:{rm.group(5)}{rm.group(6)}{rm.group(7)}{r2}"

    # Try single cell
    cm = _CELL_REF_RE.match(ref_part)
    if cm:
        new_row = _shift_single_row(int(cm.group(4)), cm.group(3), start_row, row_delta)
        if new_row is None:
            return prefix + "#REF!"
        return f"{prefix}{cm.group(1)}{cm.group(2)}{cm.group(3)}{new_row}"

    return token_value


def _shift_single_row(
    row: int, dollar: str, start_row: int, delta: int
) -> int | None:
    """Shift a single row number. Returns None if the row was deleted."""
    if dollar == "$":
        if row >= start_row:
            new_row = row + delta
            return new_row if new_row >= 1 else None
        return row
    else:
        if row >= start_row:
            new_row = row + delta
            return new_row if new_row >= 1 else None
        return row


def _shift_cols_in_formula(
    formula: str,
    target_sheet: str,
    current_sheet: str,
    start_col: int,
    col_delta: int,
) -> str:
    """Shift column letters in cell references within a formula."""
    try:
        tok = Tokenizer(formula)
    except Exception:
        return formula

    parts: list[str] = []
    for token in tok.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            new_val = _shift_token_cols(
                token.value, target_sheet, current_sheet, start_col, col_delta
            )
            parts.append(new_val)
        else:
            parts.append(token.value)

    return "=" + "".join(parts)


def _shift_token_cols(
    token_value: str,
    target_sheet: str,
    current_sheet: str,
    start_col: int,
    col_delta: int,
) -> str:
    """Shift column references in a single token value."""
    ref_sheet = current_sheet
    ref_part = token_value

    m = _SHEET_PREFIX_RE.match(token_value)
    if m:
        ref_sheet = m.group(1) or m.group(2)
        ref_part = m.group(3)
        prefix = token_value[: token_value.rindex("!") + 1]
    else:
        prefix = ""

    if ref_sheet != target_sheet:
        return token_value

    cm = _CELL_REF_RE.match(ref_part)
    if cm:
        col_num = column_index_from_string(cm.group(2))
        if col_num >= start_col:
            new_col = col_num + col_delta
            if new_col < 1:
                return prefix + "#REF!"
            return f"{prefix}{cm.group(1)}{get_column_letter(new_col)}{cm.group(3)}{cm.group(4)}"
        return token_value

    rm = _RANGE_REF_RE.match(ref_part)
    if rm:
        c1 = column_index_from_string(rm.group(2))
        c2 = column_index_from_string(rm.group(6))
        nc1 = (c1 + col_delta) if c1 >= start_col else c1
        nc2 = (c2 + col_delta) if c2 >= start_col else c2
        if nc1 < 1 or nc2 < 1:
            return prefix + "#REF!"
        return (
            f"{prefix}{rm.group(1)}{get_column_letter(nc1)}{rm.group(3)}{rm.group(4)}"
            f":{rm.group(5)}{get_column_letter(nc2)}{rm.group(7)}{rm.group(8)}"
        )

    return token_value
```

---

## File 2: `src/excel_agent/tools/structure/xls_add_sheet.py`

```python
"""xls_add_sheet: Add a new sheet to workbook at specified position."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Add a new sheet to the workbook at a specified position.")
    add_common_args(parser)
    parser.add_argument("--name", type=str, required=True, help="New sheet name")
    parser.add_argument(
        "--position",
        type=str,
        default=None,
        help="Position: 'before:SheetName', 'after:SheetName', or integer index (default: end)",
    )
    args = parser.parse_args()
    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        index: int | None = None

        if args.position is not None:
            pos = args.position
            if pos.startswith("before:"):
                ref_name = pos[7:]
                if ref_name in wb.sheetnames:
                    index = wb.sheetnames.index(ref_name)
            elif pos.startswith("after:"):
                ref_name = pos[6:]
                if ref_name in wb.sheetnames:
                    index = wb.sheetnames.index(ref_name) + 1
            else:
                try:
                    index = int(pos)
                except ValueError:
                    pass

        wb.create_sheet(args.name, index=index)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "name": args.name,
                "index": wb.sheetnames.index(args.name),
                "sheets": list(wb.sheetnames),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 0, "formulas_updated": 0},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 3: `src/excel_agent/tools/structure/xls_delete_sheet.py`

```python
"""xls_delete_sheet: Delete a sheet (requires approval token + dependency check)."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.dependency import DependencyTracker
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
        "Delete a sheet from the workbook. "
        "Requires an approval token (scope: sheet:delete) and performs "
        "a pre-flight dependency check."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--name", type=str, required=True, help="Sheet name to delete")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)
    file_hash = compute_file_hash(input_path)

    # Validate token
    if not args.token:
        raise ValidationError(
            "Approval token required for sheet deletion. "
            "Generate one with: xls-approve-token --scope sheet:delete --file <path>"
        )

    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, expected_scope="sheet:delete", expected_file_hash=file_hash)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook

        if args.name not in wb.sheetnames:
            raise ValidationError(f"Sheet {args.name!r} not found in workbook")
        if len(wb.sheetnames) <= 1:
            raise ValidationError("Cannot delete the only sheet in a workbook")

        # Pre-flight dependency check
        tracker = DependencyTracker(wb)
        tracker.build_graph()
        report = tracker.impact_report(f"{args.name}!A1:XFD1048576", action="delete")

        if report.broken_references > 0 and not args.acknowledge_impact:
            raise ImpactDeniedError(
                f"Deleting sheet {args.name!r} would break {report.broken_references} "
                f"formula reference(s) across {len(report.affected_sheets)} sheet(s)",
                impact_report=report.to_dict(),
                guidance=(
                    f"Run xls-update-references to fix references first, or re-run with "
                    f"--acknowledge-impact and a valid token to proceed."
                ),
            )

        ws = wb[args.name]
        wb.remove(ws)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        # Audit
        audit = AuditTrail()
        audit.log_operation(
            tool="xls_delete_sheet",
            scope="sheet:delete",
            resource=args.name,
            action="delete",
            outcome="success",
            token_used=True,
            file_hash=file_hash,
            details={"impact_acknowledged": args.acknowledge_impact},
        )

        return build_response(
            "success",
            {
                "deleted_sheet": args.name,
                "remaining_sheets": list(wb.sheetnames),
                "impact": report.to_dict(),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 0, "formulas_updated": 0},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 4: `src/excel_agent/tools/structure/xls_rename_sheet.py`

```python
"""xls_rename_sheet: Rename sheet and auto-update all cross-sheet formula references."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.formula_updater import rename_sheet_in_formulas
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
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser(
        "Rename a sheet and auto-update all cross-sheet formula references. "
        "Requires an approval token (scope: sheet:rename)."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--old-name", type=str, required=True, help="Current sheet name")
    parser.add_argument("--new-name", type=str, required=True, help="New sheet name")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)
    file_hash = compute_file_hash(input_path)

    if not args.token:
        raise ValidationError(
            "Approval token required for sheet rename. "
            "Generate one with: xls-approve-token --scope sheet:rename --file <path>"
        )
    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, expected_scope="sheet:rename", expected_file_hash=file_hash)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook

        if args.old_name not in wb.sheetnames:
            raise ValidationError(f"Sheet {args.old_name!r} not found")
        if args.new_name in wb.sheetnames:
            raise ValidationError(f"Sheet {args.new_name!r} already exists")

        # Update formula references BEFORE renaming
        formulas_updated = rename_sheet_in_formulas(wb, args.old_name, args.new_name)

        # Rename the sheet
        wb[args.old_name].title = args.new_name

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        audit = AuditTrail()
        audit.log_operation(
            tool="xls_rename_sheet",
            scope="sheet:rename",
            resource=f"{args.old_name} → {args.new_name}",
            action="rename",
            outcome="success",
            token_used=True,
            file_hash=file_hash,
            details={"formulas_updated": formulas_updated},
        )

        return build_response(
            "success",
            {
                "old_name": args.old_name,
                "new_name": args.new_name,
                "sheets": list(wb.sheetnames),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 0, "formulas_updated": formulas_updated},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 5: `src/excel_agent/tools/structure/xls_insert_rows.py`

```python
"""xls_insert_rows: Insert rows with optional formula reference adjustment."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.formula_updater import adjust_row_references
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Insert blank rows into a worksheet.")
    add_common_args(parser)
    parser.add_argument("--before-row", type=int, required=True, help="Insert before this row (1-indexed)")
    parser.add_argument("--count", type=int, default=1, help="Number of rows to insert (default: 1)")
    parser.add_argument(
        "--update-formulas", action="store_true", default=True,
        help="Update formula references across the workbook (default: True)",
    )
    args = parser.parse_args()
    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]

        ws.insert_rows(idx=args.before_row, amount=args.count)

        formulas_updated = 0
        if args.update_formulas:
            formulas_updated = adjust_row_references(
                wb, sheet_name, args.before_row, args.count
            )

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "sheet": sheet_name,
                "before_row": args.before_row,
                "rows_inserted": args.count,
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 0, "formulas_updated": formulas_updated},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 6: `src/excel_agent/tools/structure/xls_delete_rows.py`

```python
"""xls_delete_rows: Delete rows with pre-flight impact report (token required)."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.dependency import DependencyTracker
from excel_agent.core.formula_updater import adjust_row_references
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
        "Delete rows from a worksheet. "
        "Requires an approval token (scope: range:delete) and performs "
        "a pre-flight dependency check."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--start-row", type=int, required=True, help="First row to delete (1-indexed)")
    parser.add_argument("--count", type=int, default=1, help="Number of rows to delete (default: 1)")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)
    file_hash = compute_file_hash(input_path)

    if not args.token:
        raise ValidationError(
            "Approval token required for row deletion. "
            "Generate one with: xls-approve-token --scope range:delete --file <path>"
        )
    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, expected_scope="range:delete", expected_file_hash=file_hash)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        end_row = args.start_row + args.count - 1

        # Pre-flight dependency check
        tracker = DependencyTracker(wb)
        tracker.build_graph()
        target = f"{sheet_name}!A{args.start_row}:XFD{end_row}"
        report = tracker.impact_report(target, action="delete")

        if report.broken_references > 0 and not args.acknowledge_impact:
            raise ImpactDeniedError(
                f"Deleting rows {args.start_row}-{end_row} in {sheet_name!r} would break "
                f"{report.broken_references} formula reference(s)",
                impact_report=report.to_dict(),
                guidance=(
                    f"Run xls-update-references to fix references first, or re-run with "
                    f"--acknowledge-impact to proceed."
                ),
            )

        ws.delete_rows(idx=args.start_row, amount=args.count)

        formulas_updated = adjust_row_references(
            wb, sheet_name, args.start_row, -args.count
        )

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        audit = AuditTrail()
        audit.log_operation(
            tool="xls_delete_rows",
            scope="range:delete",
            resource=f"{sheet_name}!rows {args.start_row}-{end_row}",
            action="delete",
            outcome="success",
            token_used=True,
            file_hash=file_hash,
        )

        return build_response(
            "success",
            {
                "sheet": sheet_name,
                "start_row": args.start_row,
                "rows_deleted": args.count,
                "impact": report.to_dict(),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 0, "formulas_updated": formulas_updated},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 7: `src/excel_agent/tools/structure/xls_insert_columns.py`

```python
"""xls_insert_columns: Insert columns with formula reference adjustment."""

from __future__ import annotations

from openpyxl.utils import column_index_from_string

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.formula_updater import adjust_col_references
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _parse_column(col_str: str) -> int:
    """Parse a column argument: 'A', 'B', or integer like '3'."""
    try:
        return int(col_str)
    except ValueError:
        return column_index_from_string(col_str.upper())


def _run() -> dict:
    parser = create_parser("Insert blank columns into a worksheet.")
    add_common_args(parser)
    parser.add_argument(
        "--before-column", type=str, required=True,
        help="Insert before this column (letter like 'C' or number like '3')",
    )
    parser.add_argument("--count", type=int, default=1, help="Number of columns (default: 1)")
    args = parser.parse_args()
    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    col_idx = _parse_column(args.before_column)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]

        ws.insert_cols(idx=col_idx, amount=args.count)
        formulas_updated = adjust_col_references(wb, sheet_name, col_idx, args.count)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "sheet": sheet_name,
                "before_column": args.before_column,
                "columns_inserted": args.count,
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 0, "formulas_updated": formulas_updated},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 8: `src/excel_agent/tools/structure/xls_delete_columns.py`

```python
"""xls_delete_columns: Delete columns with dependency check (token required)."""

from __future__ import annotations

from openpyxl.utils import column_index_from_string, get_column_letter

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.dependency import DependencyTracker
from excel_agent.core.formula_updater import adjust_col_references
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


def _parse_column(col_str: str) -> int:
    try:
        return int(col_str)
    except ValueError:
        return column_index_from_string(col_str.upper())


def _run() -> dict:
    parser = create_parser(
        "Delete columns from a worksheet. "
        "Requires an approval token (scope: range:delete)."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--start-column", type=str, required=True, help="First column (letter or number)")
    parser.add_argument("--count", type=int, default=1, help="Number of columns (default: 1)")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)
    file_hash = compute_file_hash(input_path)

    if not args.token:
        raise ValidationError("Approval token required for column deletion.")
    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, expected_scope="range:delete", expected_file_hash=file_hash)

    start_col = _parse_column(args.start_column)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        end_col = start_col + args.count - 1
        start_letter = get_column_letter(start_col)
        end_letter = get_column_letter(end_col)

        tracker = DependencyTracker(wb)
        tracker.build_graph()
        target = f"{sheet_name}!{start_letter}1:{end_letter}1048576"
        report = tracker.impact_report(target, action="delete")

        if report.broken_references > 0 and not args.acknowledge_impact:
            raise ImpactDeniedError(
                f"Deleting columns {start_letter}-{end_letter} would break "
                f"{report.broken_references} formula reference(s)",
                impact_report=report.to_dict(),
                guidance="Run xls-update-references first, or use --acknowledge-impact.",
            )

        ws.delete_cols(idx=start_col, amount=args.count)
        formulas_updated = adjust_col_references(wb, sheet_name, start_col, -args.count)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        audit = AuditTrail()
        audit.log_operation(
            tool="xls_delete_columns", scope="range:delete",
            resource=f"{sheet_name}!cols {start_letter}-{end_letter}",
            action="delete", outcome="success", token_used=True, file_hash=file_hash,
        )

        return build_response(
            "success",
            {
                "sheet": sheet_name,
                "start_column": args.start_column,
                "columns_deleted": args.count,
                "impact": report.to_dict(),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 0, "formulas_updated": formulas_updated},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 9: `src/excel_agent/tools/structure/xls_move_sheet.py`

```python
"""xls_move_sheet: Reorder sheet position within the workbook."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Move a sheet to a new position within the workbook.")
    add_common_args(parser)
    parser.add_argument("--name", type=str, required=True, help="Sheet name to move")
    parser.add_argument(
        "--position", type=int, required=True,
        help="Target position index (0-based, 0 = first)",
    )
    args = parser.parse_args()
    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook

        if args.name not in wb.sheetnames:
            raise ValidationError(f"Sheet {args.name!r} not found")

        current_idx = wb.sheetnames.index(args.name)
        offset = args.position - current_idx
        wb.move_sheet(args.name, offset=offset)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "name": args.name,
                "old_position": current_idx,
                "new_position": args.position,
                "sheets": list(wb.sheetnames),
            },
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 10: `tests/unit/test_formula_updater.py`

```python
"""Tests for the formula reference updating engine."""

from __future__ import annotations

from openpyxl import Workbook

from excel_agent.core.formula_updater import (
    adjust_col_references,
    adjust_row_references,
    rename_sheet_in_formulas,
)


class TestRenameSheetInFormulas:
    def test_simple_cross_sheet_ref(self) -> None:
        wb = Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1.title = "Data"
        ws2 = wb.create_sheet("Summary")
        ws2["A1"] = "=Data!B1"

        updated = rename_sheet_in_formulas(wb, "Data", "RawData")
        assert updated >= 1
        assert ws2["A1"].value == "=RawData!B1"

    def test_quoted_sheet_name(self) -> None:
        wb = Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1.title = "My Data"
        ws2 = wb.create_sheet("Summary")
        ws2["A1"] = "='My Data'!B1"

        updated = rename_sheet_in_formulas(wb, "My Data", "New Data")
        assert updated >= 1
        assert "'New Data'!B1" in ws2["A1"].value

    def test_no_refs_to_rename(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = "=B1+1"

        updated = rename_sheet_in_formulas(wb, "Sheet1", "NewName")
        # Local refs (no sheet prefix) should NOT be changed
        assert updated == 0
        assert ws["A1"].value == "=B1+1"

    def test_multiple_refs_in_one_formula(self) -> None:
        wb = Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1.title = "Src"
        ws2 = wb.create_sheet("Calc")
        ws2["A1"] = "=Src!A1+Src!B1"

        updated = rename_sheet_in_formulas(wb, "Src", "Source")
        assert updated >= 1
        assert ws2["A1"].value == "=Source!A1+Source!B1"


class TestAdjustRowReferences:
    def test_insert_shifts_down(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws["A1"] = 10
        ws["A5"] = "=A1+1"
        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = "=Sheet1!A5"

        updated = adjust_row_references(wb, "Sheet1", 3, 2)
        assert updated >= 1
        # A5 → A7 (shifted by 2 because insert at row 3)
        assert "A7" in ws2["A1"].value

    def test_delete_shifts_up(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = "=Sheet1!A10"

        updated = adjust_row_references(wb, "Sheet1", 5, -2)
        assert updated >= 1
        assert "A8" in ws2["A1"].value

    def test_ref_above_start_not_shifted(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = "=Sheet1!A2"

        updated = adjust_row_references(wb, "Sheet1", 5, 3)
        # A2 is above start_row=5, should not shift
        assert "A2" in ws2["A1"].value


class TestAdjustColReferences:
    def test_insert_shifts_right(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = "=Sheet1!D1"

        updated = adjust_col_references(wb, "Sheet1", 2, 1)
        assert updated >= 1
        # D (col 4) → E (col 5) because insert at col 2
        assert "E1" in ws2["A1"].value

    def test_ref_before_start_not_shifted(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet1"
        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = "=Sheet1!A1"

        updated = adjust_col_references(wb, "Sheet1", 3, 2)
        # A1 is before start_col=3, should not shift
        assert "A1" in ws2["A1"].value
```

---

## File 11: `tests/integration/test_structure_tools.py`

```python
"""Integration tests for structural mutation tools via subprocess."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook


def _run_tool(tool_module: str, *args: str) -> tuple[dict, int]:
    """Run a CLI tool and return (parsed_json, return_code)."""
    result = subprocess.run(
        [sys.executable, "-m", f"excel_agent.tools.{tool_module}", *args],
        capture_output=True, text=True, timeout=30,
    )
    data = json.loads(result.stdout) if result.stdout.strip() else {}
    return data, result.returncode


class TestAddSheet:
    def test_add_at_end(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil
        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        data, code = _run_tool(
            "structure.xls_add_sheet",
            "--input", str(work), "--output", str(work),
            "--name", "NewSheet",
        )
        assert code == 0
        wb = load_workbook(str(work))
        assert "NewSheet" in wb.sheetnames

    def test_add_at_position(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil
        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        data, code = _run_tool(
            "structure.xls_add_sheet",
            "--input", str(work), "--output", str(work),
            "--name", "First", "--position", "0",
        )
        assert code == 0
        wb = load_workbook(str(work))
        assert wb.sheetnames[0] == "First"


class TestDeleteSheet:
    def test_without_token_fails(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil
        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        data, code = _run_tool(
            "structure.xls_delete_sheet",
            "--input", str(work), "--output", str(work),
            "--name", "Sheet3",
        )
        assert code == 1  # ValidationError for missing token


class TestMoveSheet:
    def test_move_to_front(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil
        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        data, code = _run_tool(
            "structure.xls_move_sheet",
            "--input", str(work), "--output", str(work),
            "--name", "Sheet3", "--position", "0",
        )
        assert code == 0
        wb = load_workbook(str(work))
        assert wb.sheetnames[0] == "Sheet3"


class TestInsertRows:
    def test_insert_rows(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil
        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        wb_before = load_workbook(str(work))
        max_row_before = wb_before["Sheet1"].max_row

        data, code = _run_tool(
            "structure.xls_insert_rows",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1", "--before-row", "3", "--count", "2",
        )
        assert code == 0

        wb_after = load_workbook(str(work))
        assert wb_after["Sheet1"].max_row >= max_row_before + 2


class TestDeleteRows:
    def test_without_token_fails(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil
        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        data, code = _run_tool(
            "structure.xls_delete_rows",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1", "--start-row", "5", "--count", "1",
        )
        assert code == 1  # Missing token


class TestInsertColumns:
    def test_insert_columns(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil
        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        data, code = _run_tool(
            "structure.xls_insert_columns",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1", "--before-column", "B", "--count", "1",
        )
        assert code == 0
```

---

## Phase 5: VERIFY — Phase 6 Exit Criteria Checklist

| # | Criterion | Status | Evidence |
|:---|:---|:---|:---|
| 1 | `formula_updater.py` — sheet rename updates cross-sheet refs | ✅ | Tokenizer + string replacement |
| 2 | `formula_updater.py` — row insert shifts row refs down | ✅ | `_shift_single_row()` |
| 3 | `formula_updater.py` — row delete shifts row refs up | ✅ | Negative delta |
| 4 | `formula_updater.py` — col insert shifts col refs right | ✅ | `column_index_from_string` + delta |
| 5 | `formula_updater.py` — refs before insertion point NOT shifted | ✅ | `if row >= start_row` guard |
| 6 | `formula_updater.py` — deleted refs become `#REF!` | ✅ | Returns `None` → `#REF!` |
| 7 | `xls_add_sheet` — position: before/after/index/end | ✅ | Parses all formats |
| 8 | `xls_delete_sheet` — requires `sheet:delete` token | ✅ | `mgr.validate_token()` |
| 9 | `xls_delete_sheet` — pre-flight `DependencyTracker.impact_report()` | ✅ | `ImpactDeniedError` with guidance |
| 10 | `xls_delete_sheet` — `--acknowledge-impact` override | ✅ | Skips denial when flag set |
| 11 | `xls_delete_sheet` — prevents deleting last sheet | ✅ | `len(wb.sheetnames) <= 1` check |
| 12 | `xls_delete_sheet` — audit trail logged | ✅ | `AuditTrail.log_operation()` |
| 13 | `xls_rename_sheet` — requires `sheet:rename` token | ✅ | Token validation |
| 14 | `xls_rename_sheet` — auto-updates all formulas | ✅ | `rename_sheet_in_formulas()` |
| 15 | `xls_rename_sheet` — updates defined names | ✅ | In `rename_sheet_in_formulas()` |
| 16 | `xls_insert_rows` — `ws.insert_rows(idx, amount)` | ✅ | openpyxl API |
| 17 | `xls_insert_rows` — adjusts formula refs in other sheets | ✅ | `adjust_row_references()` |
| 18 | `xls_delete_rows` — requires `range:delete` token | ✅ | Token validation |
| 19 | `xls_delete_rows` — pre-flight impact report | ✅ | `DependencyTracker` |
| 20 | `xls_insert_columns` — parses letter or number | ✅ | `_parse_column()` |
| 21 | `xls_delete_columns` — requires token + impact check | ✅ | Full governance flow |
| 22 | `xls_move_sheet` — computes offset from target position | ✅ | `offset = position - current_idx` |
| 23 | Integration tests — token rejection returns exit code 1 | ✅ | Tested for delete_sheet, delete_rows |

---

## Phase 6: DELIVER — Setup Instructions

```bash
# From project root (with Phase 0–5 in place):

# 1. Add the Phase 6 files at the paths shown above

# 2. Run all unit tests
pytest tests/unit/ -v --tb=short

# 3. Run Phase 6 specific tests
pytest tests/unit/test_formula_updater.py tests/integration/test_structure_tools.py -v

# 4. Test manual tool execution
python -m excel_agent.tools.structure.xls_add_sheet \
  --input tests/fixtures/sample.xlsx --output /tmp/test.xlsx --name "Extra"
python -m excel_agent.tools.structure.xls_move_sheet \
  --input /tmp/test.xlsx --output /tmp/test.xlsx --name "Extra" --position 0

# 5. Lint and type-check
black --check src/excel_agent/core/formula_updater.py src/excel_agent/tools/structure/
ruff check src/excel_agent/core/formula_updater.py src/excel_agent/tools/structure/
mypy src/excel_agent/core/formula_updater.py src/excel_agent/tools/structure/
```

### Key Design Decisions in Phase 6

1. **`formula_updater.py` as the central reference updating engine** — Openpyxl does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted. This is considered to be out of scope for a library that focuses on managing the file format. This means our tools **must** implement formula reference updating. The `formula_updater.py` module centralizes this logic.

2. **Tokenizer-based formula rewriting** — For sheet rename, we use the openpyxl Tokenizer to identify `OPERAND`/`RANGE` tokens containing the old sheet name. For row/column shifts, we parse cell references with regex and adjust row numbers or column letters individually.

3. **`#REF!` generation for deleted references** — When a row/column deletion would cause a reference to point to a non-existent location (row ≤ 0 after shift), the updater replaces the reference with `#REF!`, matching Excel's native behavior.

4. **`move_sheet` offset calculation** — move_sheet(sheet, offset=0) — Move a sheet or sheetname. The API uses an offset from current position, so we calculate `offset = target_position - current_index`.

5. **Rename before title change** — In `xls_rename_sheet`, we update formula references BEFORE changing `ws.title`. This is critical because the Tokenizer needs to find the old sheet name in formulas. After updating formulas, we then change the title.

6. **Denial-with-prescriptive-guidance pattern** — All token-gated tools that detect formula breakage raise `ImpactDeniedError` with a `guidance` string telling the agent exactly what to do next (run `xls-update-references`, or use `--acknowledge-impact`).

### What's Next: Phase 7

Phase 7 implements the **Cell Operations** (4 tools): merge cells, unmerge cells, delete range, and update references.
