"""xls_delete_rows: Delete rows with pre-flight impact report (token required)."""

from __future__ import annotations

from excel_agent.core.dependency import DependencyTracker
from excel_agent.core.edit_session import EditSession
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
    parser.add_argument(
        "--start-row", type=int, required=True, help="First row to delete (1-indexed)"
    )
    parser.add_argument(
        "--count", type=int, default=1, help="Number of rows to delete (default: 1)"
    )
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
    mgr.validate_token(args.token, "range:delete", input_path)

        ws.delete_rows(idx=args.start_row, amount=args.count)

        formulas_updated = adjust_row_references(wb, sheet_name, args.start_row, -args.count)

        # Capture hashes before exiting context
        version_hash = session.version_hash
        file_hash = session.file_hash

        # EditSession handles save automatically on exit

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
        workbook_version=version_hash,
        impact={"cells_modified": 0, "formulas_updated": formulas_updated},
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
