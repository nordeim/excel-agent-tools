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

from excel_agent.core.edit_session import EditSession
from excel_agent.core.version_hash import compute_file_hash
from excel_agent.governance.audit_trail import AuditTrail
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


def _run() -> dict[str, object]:
    parser = create_parser("Batch-update cell references in all formulas across the workbook.")
    add_common_args(parser)
    parser.add_argument(
        "--updates",
        type=str,
        required=True,
        help='JSON array of {"old": "ref", "new": "ref"} mappings',
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)
    file_hash = compute_file_hash(input_path)

    updates_raw = parse_json_arg(args.updates)
    if not isinstance(updates_raw, list):
        raise ValidationError("--updates must be a JSON array of {old, new} objects")

    # Build normalized update map: normalized_old → raw_new
    update_map: dict[str, str] = {}
    for entry in updates_raw:
        if not isinstance(entry, dict) or "old" not in entry or "new" not in entry:
            raise ValidationError(f"Each update must have 'old' and 'new' keys, got: {entry}")
        old_norm = _normalize_ref(str(entry["old"]))
        update_map[old_norm] = str(entry["new"])

    session = EditSession.prepare(input_path, output_path)
    with session:
        wb = session.workbook
        formulas_updated = 0
        update_details: list[dict[str, object]] = []

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
                            update_details.append(
                                {
                                    "sheet": ws.title,
                                    "cell": cell.coordinate,
                                    "old_formula": original,
                                    "new_formula": updated,
                                }
                            )

        # Also update defined names
        defined_names_updated = 0
        for _name, defn in wb.defined_names.items():
            if hasattr(defn, "attr_text") and defn.attr_text:
                original_text = defn.attr_text
                new_text = original_text
                for old_norm, new_ref in update_map.items():
                    # Simple string replacement in defined names
                    new_text = new_text.replace(old_norm.replace("$", ""), new_ref)
                if new_text != original_text:
                    defn.attr_text = new_text
                    defined_names_updated += 1

        # Capture version hash before exiting context
        version_hash = session.version_hash

        # EditSession handles save automatically on exit

        # Audit - log the reference update operation
        audit = AuditTrail()
        audit.log(
            tool="xls_update_references",
            scope="structure:modify",
            target_file=output_path,
            file_version_hash=file_hash,
            actor_nonce="system",  # No token required for this remediation tool
            operation_details={
                "updates_requested": len(updates_raw),
                "update_details": update_details,
            },
            impact={
                "cells_modified": 0,
                "formulas_updated": formulas_updated + defined_names_updated,
            },
            success=True,
            exit_code=0,
        )

        return build_response(
            "success",
            {
                "updates_requested": len(updates_raw),
                "formulas_updated": formulas_updated,
                "defined_names_updated": defined_names_updated,
                "update_details": update_details,
            },
            workbook_version=version_hash,
            impact={
                "cells_modified": 0,
                "formulas_updated": formulas_updated + defined_names_updated,
            },
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
