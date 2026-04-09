"""xls_add_sheet: Add a new sheet to workbook at specified position."""

from __future__ import annotations

import contextlib

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
                with contextlib.suppress(ValueError):
                    index = int(pos)

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
