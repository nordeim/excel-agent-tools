"""xls_set_data_validation: Configure data validation rules.

Supports list, whole number, decimal, date, time, text length, and custom
validation types with input messages and error alerts.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation

from excel_agent.core.version_hash import compute_file_hash
from excel_agent.governance.audit_trail import AuditTrail
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response

# Data validation types
VALIDATION_TYPES = {
    "list": "List validation from comma-separated values",
    "whole": "Whole number (integer)",
    "decimal": "Decimal number",
    "date": "Date value",
    "time": "Time value",
    "textLength": "Text length constraint",
    "custom": "Custom formula",
}

# Excel limit for list validation
MAX_LIST_LENGTH = 255


def _run() -> dict[str, object]:
    parser = create_parser("Configure data validation rules for cell ranges.")
    add_common_args(parser)
    parser.add_argument(
        "--range",
        type=str,
        required=True,
        help='Target range (e.g., "B2:B100")',
    )
    parser.add_argument(
        "--type",
        type=str,
        required=True,
        choices=list(VALIDATION_TYPES.keys()),
        help="Validation type",
    )
    parser.add_argument(
        "--formula1",
        type=str,
        required=True,
        help="Primary constraint (e.g., 'Option1,Option2' for list, '10' for >=10)",
    )
    parser.add_argument(
        "--formula2",
        type=str,
        default=None,
        help="Secondary constraint (for between/notBetween types)",
    )
    parser.add_argument(
        "--operator",
        type=str,
        default="greaterThanOrEqual",
        choices=[
            "between",
            "notBetween",
            "equal",
            "notEqual",
            "greaterThan",
            "greaterThanOrEqual",
            "lessThan",
            "lessThanOrEqual",
        ],
        help="Comparison operator (for numeric types)",
    )
    parser.add_argument(
        "--allow-blank",
        action="store_true",
        default=True,
        help="Allow blank cells (default: True)",
    )
    parser.add_argument(
        "--show-input",
        action="store_true",
        default=False,
        help="Show input message when cell is selected",
    )
    parser.add_argument(
        "--input-title",
        type=str,
        default="",
        help="Input message title",
    )
    parser.add_argument(
        "--input-message",
        type=str,
        default="",
        help="Input message text",
    )
    parser.add_argument(
        "--show-error",
        action="store_true",
        default=True,
        help="Show error alert on invalid input (default: True)",
    )
    parser.add_argument(
        "--error-title",
        type=str,
        default="Invalid Input",
        help="Error alert title",
    )
    parser.add_argument(
        "--error-message",
        type=str,
        default="The value you entered is not valid.",
        help="Error alert message",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or str(input_path), create_parents=True)

    file_hash = compute_file_hash(input_path)

    # Load workbook
    wb = load_workbook(str(input_path))
    ws = wb[args.sheet] if args.sheet else wb.active

    # Parse range
    try:
        c1, r1, c2, r2 = range_boundaries(args.range)
        if c1 is None or r1 is None:
            raise ValueError("Invalid range format")
        if c2 is None:
            c2 = c1
        if r2 is None:
            r2 = r1
    except Exception as e:
        return build_response(
            "error",
            None,
            exit_code=1,
            warnings=[f"Failed to parse range '{args.range}': {e}"],
        )

    warnings = []

    # Validate list length
    if args.type == "list":
        if len(args.formula1) > MAX_LIST_LENGTH:
            return build_response(
                "error",
                None,
                exit_code=1,
                warnings=[
                    f"List validation exceeds {MAX_LIST_LENGTH} character limit",
                    f"Current length: {len(args.formula1)} characters",
                    "Consider using a reference range instead",
                ],
            )

    # Build formula strings
    if args.type == "list":
        # List validation - formula1 should be quoted comma-separated values
        if not args.formula1.startswith('"') and not args.formula1.startswith("="):
            formula1 = f'"{args.formula1}"'
        else:
            formula1 = args.formula1
    elif args.type == "custom":
        # Custom formula - should start with =
        if not args.formula1.startswith("="):
            formula1 = f"={args.formula1}"
        else:
            formula1 = args.formula1
    else:
        # Numeric/date/time/textLength
        formula1 = args.formula1

    formula2 = args.formula2 if args.formula2 else None

    # Create data validation
    try:
        dv = DataValidation(
            type=args.type,
            formula1=formula1,
            formula2=formula2,
            allow_blank=args.allow_blank,
            showErrorMessage=args.show_error,
            showInputMessage=args.show_input,
            errorTitle=args.error_title if args.show_error else None,
            error=args.error_message if args.show_error else None,
            promptTitle=args.input_title if args.show_input else None,
            prompt=args.input_message if args.show_input else None,
        )
        dv.add(args.range)
        ws.add_data_validation(dv)
    except Exception as e:
        return build_response(
            "error",
            None,
            exit_code=5,
            warnings=[f"Failed to add data validation: {e}"],
        )

    # Save workbook
    wb.save(str(output_path))

    # Log to audit trail
    audit = AuditTrail()
    audit.log(
        tool="xls_set_data_validation",
        scope="structure:modify",
        target_file=input_path,
        file_version_hash=file_hash,
        actor_nonce="auto",
        operation_details={
            "range": args.range,
            "validation_type": args.type,
            "formula1": args.formula1,
            "sheet": ws.title,
        },
        impact={
            "validation_added": True,
            "affected_cells": (c2 - c1 + 1) * (r2 - r1 + 1),
        },
        success=True,
        exit_code=0,
    )

    return build_response(
        "success",
        {
            "range": args.range,
            "validation_type": args.type,
            "formula1": args.formula1,
            "formula2": args.formula2,
            "sheet": ws.title,
            "allow_blank": args.allow_blank,
            "show_input": args.show_input,
            "show_error": args.show_error,
            "affected_cells": (c2 - c1 + 1) * (r2 - r1 + 1),
        },
        warnings=warnings if warnings else None,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
