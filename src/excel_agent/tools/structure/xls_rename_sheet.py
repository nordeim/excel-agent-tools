"""xls_rename_sheet: Rename sheet and auto-update all cross-sheet formula references."""

from __future__ import annotations

from excel_agent.core.edit_session import EditSession
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
    mgr.validate_token(args.token, "sheet:rename", input_path)

    session = EditSession.prepare(input_path, output_path)
    with session:
        wb = session.workbook

        if args.old_name not in wb.sheetnames:
            raise ValidationError(f"Sheet {args.old_name!r} not found")
        if args.new_name in wb.sheetnames:
            raise ValidationError(f"Sheet {args.new_name!r} already exists")

        # Update formula references BEFORE renaming
        formulas_updated = rename_sheet_in_formulas(wb, args.old_name, args.new_name)

        # Rename the sheet
        wb[args.old_name].title = args.new_name

        # Capture version hash before exiting context
        version_hash = session.version_hash

        # EditSession handles save automatically on exit

        # Audit - extract nonce from token for actor identification
        audit = AuditTrail()
        token_parts = args.token.split("|") if args.token else ["", "", ""]
        actor_nonce = token_parts[2] if len(token_parts) > 2 else ""
        audit.log(
            tool="xls_rename_sheet",
            scope="sheet:rename",
            target_file=output_path,
            file_version_hash=file_hash,
            actor_nonce=actor_nonce,
            operation_details={
                "old_name": args.old_name,
                "new_name": args.new_name,
                "formulas_updated": formulas_updated,
            },
            impact={"cells_modified": 0, "formulas_updated": formulas_updated},
            success=True,
            exit_code=0,
        )

        return build_response(
            "success",
            {
                "old_name": args.old_name,
                "new_name": args.new_name,
                "sheets": list(wb.sheetnames),
            },
            workbook_version=version_hash,
            impact={"cells_modified": 0, "formulas_updated": formulas_updated},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
