"""xls_add_chart: Create charts (Bar, Line, Pie, Scatter).

Creates various chart types from data ranges with customizable styling.
Supports Bar, Line, Pie, and Scatter charts with proper data validation.
"""

from __future__ import annotations

from typing import Any

from openpyxl import load_workbook
from openpyxl.chart import (
    BarChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
)
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import range_boundaries

from excel_agent.core.version_hash import compute_file_hash
from excel_agent.governance.audit_trail import AuditTrail
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response

# Supported chart types
CHART_TYPES = {
    "bar": BarChart,
    "line": LineChart,
    "pie": PieChart,
    "scatter": ScatterChart,
}


def _has_numeric_data(ws, min_col: int, min_row: int, max_col: int, max_row: int) -> bool:
    """Check if range contains numeric data."""
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value is not None:
                try:
                    float(cell.value)
                    return True
                except (ValueError, TypeError):
                    pass
    return False


def _create_chart(
    chart_type: str,
    data_ref: Reference,
    cats_ref: Reference | None,
    title: str,
    style: int | None,
) -> Any:
    """Create chart of specified type."""
    chart_class = CHART_TYPES[chart_type]
    chart = chart_class()

    # Set chart type specifics
    if chart_type == "bar":
        chart.type = "col"  # vertical bars
        chart.grouping = "clustered"
    elif chart_type == "line":
        chart.grouping = "standard"
    elif chart_type == "pie":
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True

    # Set title
    if title:
        chart.title = title

    # Set style (1-48 for most chart types)
    if style:
        chart.style = style

    # Add data
    chart.add_data(data_ref, titles_from_data=True)

    # Set categories (not applicable for pie charts)
    if cats_ref and chart_type != "pie":
        chart.set_categories(cats_ref)

    return chart


def _run() -> dict[str, object]:
    parser = create_parser("Create charts from data ranges.")
    add_common_args(parser)
    parser.add_argument(
        "--type",
        type=str,
        required=True,
        choices=list(CHART_TYPES.keys()),
        help="Chart type: bar, line, pie, scatter",
    )
    parser.add_argument(
        "--data-range",
        type=str,
        required=True,
        help='Data range (e.g., "B1:E7") - must include headers',
    )
    parser.add_argument(
        "--categories-range",
        type=str,
        default=None,
        help='Category labels range (e.g., "A2:A7") - optional',
    )
    parser.add_argument(
        "--title",
        type=str,
        default="",
        help="Chart title",
    )
    parser.add_argument(
        "--position",
        type=str,
        required=True,
        help='Chart position as cell reference (e.g., "G2")',
    )
    parser.add_argument(
        "--style",
        type=int,
        default=None,
        help="Chart style number (1-48, optional)",
    )
    parser.add_argument(
        "--width",
        type=int,
        default=15,
        help="Chart width in cm (default: 15)",
    )
    parser.add_argument(
        "--height",
        type=int,
        default=10,
        help="Chart height in cm (default: 10)",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or str(input_path), create_parents=True)

    file_hash = compute_file_hash(input_path)

    # Load workbook
    wb = load_workbook(str(input_path))
    ws = wb[args.sheet] if args.sheet else wb.active

    # Parse data range
    try:
        dc1, dr1, dc2, dr2 = range_boundaries(args.data_range)
        if dc1 is None or dr1 is None:
            raise ValueError("Invalid data range format")
        if dc2 is None:
            dc2 = dc1
        if dr2 is None:
            dr2 = dr1
    except Exception as e:
        return build_response(
            "error",
            None,
            exit_code=1,
            warnings=[f"Failed to parse data range '{args.data_range}': {e}"],
        )

    # Validate data range contains numeric data
    if not _has_numeric_data(ws, dc1, dr1 + 1, dc2, dr2):  # Skip header row
        return build_response(
            "error",
            None,
            exit_code=1,
            warnings=[
                f"Data range {args.data_range} does not contain numeric data",
                "Charts require numeric values for plotting",
            ],
        )

    # Parse categories range if provided
    cats_ref = None
    if args.categories_range:
        try:
            cc1, cr1, cc2, cr2 = range_boundaries(args.categories_range)
            if cc1 is None or cr1 is None:
                raise ValueError("Invalid categories range format")
            if cc2 is None:
                cc2 = cc1
            if cr2 is None:
                cr2 = cr1

            # Validate dimensions match data
            data_series_count = dr2 - dr1
            cat_count = cr2 - cr1 + 1
            if cat_count != data_series_count:
                return build_response(
                    "error",
                    None,
                    exit_code=1,
                    warnings=[
                        f"Categories range has {cat_count} items",
                        f"Data range has {data_series_count} series",
                        "Counts must match for proper chart display",
                    ],
                )

            cats_ref = Reference(ws, min_col=cc1, min_row=cr1, max_col=cc2, max_row=cr2)
        except Exception as e:
            return build_response(
                "error",
                None,
                exit_code=1,
                warnings=[f"Failed to parse categories range '{args.categories_range}': {e}"],
            )

    # Create data reference
    data_ref = Reference(ws, min_col=dc1, min_row=dr1, max_col=dc2, max_row=dr2)

    # Validate position
    try:
        from openpyxl.utils import coordinate_to_tuple

        coordinate_to_tuple(args.position)
    except Exception:
        return build_response(
            "error",
            None,
            exit_code=1,
            warnings=[f"Invalid position: {args.position}", "Must be a valid cell reference"],
        )

    # Create chart
    try:
        chart = _create_chart(args.type, data_ref, cats_ref, args.title, args.style)
        chart.width = args.width
        chart.height = args.height
    except Exception as e:
        return build_response(
            "error",
            None,
            exit_code=5,
            warnings=[f"Failed to create chart: {e}"],
        )

    # Add chart to worksheet
    ws.add_chart(chart, args.position)

    # Save workbook
    wb.save(str(output_path))

    # Log to audit trail
    audit = AuditTrail()
    audit.log(
        tool="xls_add_chart",
        scope="structure:modify",
        target_file=input_path,
        file_version_hash=file_hash,
        actor_nonce="auto",
        operation_details={
            "chart_type": args.type,
            "data_range": args.data_range,
            "position": args.position,
            "title": args.title,
            "sheet": ws.title,
        },
        impact={
            "chart_created": True,
            "chart_type": args.type,
        },
        success=True,
        exit_code=0,
    )

    return build_response(
        "success",
        {
            "chart_type": args.type,
            "data_range": args.data_range,
            "categories_range": args.categories_range,
            "position": args.position,
            "title": args.title,
            "style": args.style,
            "sheet": ws.title,
            "width_cm": args.width,
            "height_cm": args.height,
        },
        warnings=[],
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
