"""xls_recalculate: Force recalculation using two-tier strategy."""

from __future__ import annotations

from pathlib import Path

from excel_agent.calculation.tier1_engine import Tier1Calculator
from excel_agent.calculation.tier2_libreoffice import Tier2Calculator
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _run() -> dict[str, object]:
    parser = create_parser(
        "Force recalculation of all formulas. "
        "Default: Try Tier 1 (formulas library), fall back to Tier 2 (LibreOffice) if needed."
    )
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    parser.add_argument("--output", type=str, required=True, help="Output workbook path")
    parser.add_argument(
        "--tier",
        type=int,
        choices=[1, 2],
        default=None,
        help="Force specific tier: 1=formulas library, 2=LibreOffice (default: auto)",
    )
    parser.add_argument(
        "--circular", action="store_true", help="Enable circular reference support"
    )
    parser.add_argument("--timeout", type=int, default=120, help="Tier 2 timeout in seconds")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output, create_parents=True)

    if args.tier == 2:
        result = _run_tier2(input_path, output_path, args.timeout)
    elif args.tier == 1:
        result_obj = Tier1Calculator(input_path).calculate(output_path, circular=args.circular)
        result = result_obj.to_dict()
    else:
        # Auto: try Tier 1, fall back to Tier 2
        t1 = Tier1Calculator(input_path)
        t1_result = t1.calculate(output_path, circular=args.circular)

        if t1_result.unsupported_functions or t1_result.error_count > 0:
            # Fall back to Tier 2
            t2 = Tier2Calculator()
            if t2.is_available():
                t2_result = t2.recalculate(input_path, output_path, timeout=args.timeout)
                result = t2_result.to_dict()
                result["tier1_fallback_reason"] = (
                    t1_result.unsupported_functions or t1_result.errors
                )[:5]
            else:
                result = t1_result.to_dict()
                result["warnings"] = [
                    "Tier 1 had errors but Tier 2 (LibreOffice) is not available."
                ]
        else:
            result = t1_result.to_dict()

    return build_response("success", result)


def _run_tier2(input_path: Path, output_path: Path, timeout: int) -> dict[str, object]:
    t2 = Tier2Calculator()
    if not t2.is_available():
        return {
            "error": "LibreOffice not installed. Install with: apt-get install libreoffice-calc",
            "engine": "tier2_libreoffice",
            "version": t2.get_version(),
        }
    return t2.recalculate(input_path, output_path, timeout=timeout).to_dict()


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
