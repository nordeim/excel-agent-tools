"""xls_detect_errors: Scan workbook for formula errors (#REF!, #VALUE!, etc.)."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import add_common_args, create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Scan workbook for formula errors.")
    add_common_args(parser)
    args = parser.parse_args()

    path = validate_input_path(args.input)

    errors = []

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type == "f" and isinstance(cell.value, str):
                        # Check if cell value is an error
                        if cell.value.startswith("#") and cell.value.endswith("!"):
                            errors.append(
                                {
                                    "sheet": sheet_name,
                                    "cell": f"{cell.column_letter}{cell.row}",
                                    "error": cell.value,
                                }
                            )

    return build_response(
        "success",
        {
            "error_count": len(errors),
            "errors": errors,
        },
        workbook_version=agent.version_hash,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
