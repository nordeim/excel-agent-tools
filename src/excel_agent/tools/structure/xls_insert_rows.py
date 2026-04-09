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
    parser.add_argument(
        "--before-row", type=int, required=True, help="Insert before this row (1-indexed)"
    )
    parser.add_argument(
        "--count", type=int, default=1, help="Number of rows to insert (default: 1)"
    )
    parser.add_argument(
        "--update-formulas",
        action="store_true",
        default=True,
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
            formulas_updated = adjust_row_references(wb, sheet_name, args.before_row, args.count)

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
