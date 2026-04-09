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
        "--before-column",
        type=str,
        required=True,
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
