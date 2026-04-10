"""xls_define_name: Create or update named ranges in the workbook."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.governance.audit_trail import AuditTrail
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response
from openpyxl.workbook.defined_name import DefinedName


def _run() -> dict:
    parser = create_parser("Create or update a named range.")
    add_common_args(parser)
    parser.add_argument("--name", type=str, required=True, help="Name for the range")
    parser.add_argument(
        "--refers-to", type=str, required=True, help="Range reference (e.g., Sheet1!A1:B10)"
    )
    parser.add_argument(
        "--scope", type=str, default=None, help="Sheet name for sheet-scoped name (optional)"
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook

        # Create the defined name
        defn = DefinedName(args.name, attr_text=args.refers_to)

        # If sheet-scoped, add to specific sheet
        if args.scope:
            if args.scope in wb.sheetnames:
                ws = wb[args.scope]
                # Sheet-scoped names use sheet.defined_names
                ws.defined_names.add(defn)
            else:
                return build_response(
                    "error",
                    None,
                    exit_code=1,
                    warnings=[f"Sheet '{args.scope}' not found"],
                )
        else:
            # Workbook-scoped
            wb.defined_names.add(defn)

    # Log to audit
    audit = AuditTrail()
    audit.log(
        tool="xls_define_name",
        scope=None,
        resource=args.name,
        action="define_name",
        outcome="success",
        token_used=False,
        file_hash=str(agent.version_hash),
    )

    return build_response(
        "success",
        {
            "name": args.name,
            "refers_to": args.refers_to,
            "scope": args.scope or "workbook",
        },
        impact={"cells_modified": 0, "named_ranges_added": 1},
        workbook_version=agent.version_hash,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
