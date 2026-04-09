"""xls_get_workbook_metadata: High-level workbook statistics."""

from __future__ import annotations

import os
import zipfile
from pathlib import Path

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Get high-level workbook statistics.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)
    file_size = os.path.getsize(path)

    has_macros = False
    try:
        with zipfile.ZipFile(path, "r") as zf:
            has_macros = "xl/vbaProject.bin" in zf.namelist()
    except zipfile.BadZipFile:
        pass

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        total_formulas = 0
        total_tables = 0

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        total_formulas += 1
            total_tables += len(ws.tables)

        named_range_count = len(wb.defined_names)

        return build_response(
            "success",
            {
                "sheet_count": len(wb.sheetnames),
                "total_formulas": total_formulas,
                "named_ranges": named_range_count,
                "tables": total_tables,
                "has_macros": has_macros,
                "file_size_bytes": file_size,
            },
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
