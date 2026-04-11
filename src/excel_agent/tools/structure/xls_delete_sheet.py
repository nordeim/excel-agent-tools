"""xls_delete_sheet: Delete a sheet (requires approval token + dependency check)."""

from __future__ import annotations

from excel_agent.core.edit_session import EditSession
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
    mgr.validate_token(args.token, "sheet:delete", input_path)

    session = EditSession.prepare(input_path, output_path)
    with session:
        wb = session.workbook

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

        # Capture version hash before exiting context
        version_hash = session.version_hash

        # EditSession handles save automatically on exit

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
            workbook_version=version_hash,
            impact={"cells_modified": 0, "formulas_updated": 0},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
