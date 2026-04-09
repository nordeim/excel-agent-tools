"""xls_write_range: Write 2D data array to a range with type inference."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.serializers import RangeSerializer
from excel_agent.core.type_coercion import infer_cell_value
from excel_agent.governance.schemas import validate_against_schema
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    parse_json_arg,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Write a 2D data array to a cell range with type inference.")
    add_common_args(parser)
    parser.add_argument("--range", type=str, required=True, help="Start cell (e.g., A1)")
    parser.add_argument(
        "--data",
        type=str,
        required=True,
        help='JSON 2D array (e.g., \'[["Name","Age"],["Alice",30]]\')',
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    data_parsed = parse_json_arg(args.data)
    validate_against_schema("write_data.schema.json", {"data": data_parsed})
    data: list[list] = data_parsed

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        serializer = RangeSerializer(workbook=wb)
        coord = serializer.parse(args.range, default_sheet=args.sheet)

        sheet_name = coord.sheet or (args.sheet or wb.sheetnames[0])
        ws = wb[sheet_name]

        start_row = coord.min_row
        start_col = coord.min_col

        cells_written = 0
        formulas_written = 0

        for row_idx, row_data in enumerate(data):
            for col_idx, raw_value in enumerate(row_data):
                cell_row = start_row + row_idx
                cell_col = start_col + col_idx
                cell = ws.cell(row=cell_row, column=cell_col)

                coerced = infer_cell_value(raw_value)
                cell.value = coerced
                cells_written += 1

                if isinstance(coerced, str) and coerced.startswith("="):
                    formulas_written += 1

        # If output differs from input, save to new path
        if str(output_path) != str(input_path):
            wb.save(str(output_path))

    return build_response(
        "success",
        {
            "range": args.range,
            "sheet": sheet_name,
            "rows_written": len(data),
            "cols_written": max((len(row) for row in data), default=0),
        },
        workbook_version=agent.version_hash,
        impact={"cells_modified": cells_written, "formulas_updated": formulas_written},
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
