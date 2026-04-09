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
        "--position",
        type=int,
        required=True,
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
