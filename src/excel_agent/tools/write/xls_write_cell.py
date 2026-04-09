"""xls_write_cell: Write a single cell with explicit type coercion."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.type_coercion import coerce_cell_value, infer_cell_value
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response

_VALID_TYPES = frozenset(
    {
        "string",
        "number",
        "integer",
        "float",
        "boolean",
        "date",
        "datetime",
        "formula",
    }
)


def _run() -> dict:
    parser = create_parser("Write a single cell value with optional explicit type coercion.")
    add_common_args(parser)
    parser.add_argument("--cell", type=str, required=True, help="Cell reference (e.g., A1)")
    parser.add_argument("--value", type=str, required=True, help="Value to write")
    parser.add_argument(
        "--type",
        type=str,
        default=None,
        choices=sorted(_VALID_TYPES),
        help="Explicit type coercion (overrides auto-inference)",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        cell = ws[args.cell]

        if args.type is not None:
            coerced = coerce_cell_value(args.value, args.type)
        else:
            coerced = infer_cell_value(args.value)

        cell.value = coerced

        is_formula = isinstance(coerced, str) and coerced.startswith("=")

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

    return build_response(
        "success",
        {
            "cell": args.cell,
            "sheet": sheet_name,
            "value_written": str(coerced),
            "type_used": args.type or "auto",
            "is_formula": is_formula,
        },
        workbook_version=agent.version_hash,
        impact={
            "cells_modified": 1,
            "formulas_updated": 1 if is_formula else 0,
        },
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
