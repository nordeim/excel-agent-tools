"""
Base runner for all excel-agent-tools CLI tools.

Provides a standard execution wrapper that:
1. Catches ExcelAgentError → maps to JSON error + exit code
2. Catches unexpected exceptions → exit code 5
3. Ensures only JSON goes to stdout
"""

from __future__ import annotations

import sys
import traceback
from typing import Callable

from excel_agent.utils.exceptions import ExcelAgentError
from excel_agent.utils.exit_codes import ExitCode
from excel_agent.utils.json_io import build_response, print_json


def run_tool(func: Callable[[], dict]) -> None:
    """Execute a tool function with standard error handling.

    Args:
        func: A callable that returns a dict suitable for print_json().
            Should raise ExcelAgentError subclasses for known errors.
    """
    try:
        result = func()
        print_json(result)
        sys.exit(result.get("exit_code", 0))
    except ExcelAgentError as exc:
        # Determine status based on exit code
        status = "denied" if exc.exit_code == 4 else "error"
        error_response = build_response(
            status,
            None,
            exit_code=exc.exit_code,
            warnings=[str(exc)],
        )
        error_response["error"] = str(exc)
        error_response["details"] = exc.details

        # Special handling for ImpactDeniedError
        if hasattr(exc, "impact_report"):
            error_response["impact"] = exc.impact_report  # type: ignore[attr-defined]
        if hasattr(exc, "guidance"):
            error_response["guidance"] = exc.guidance  # type: ignore[attr-defined]

        print_json(error_response)
        sys.exit(exc.exit_code)
    except Exception as exc:
        error_response = build_response(
            "error",
            None,
            exit_code=ExitCode.INTERNAL_ERROR,
        )
        error_response["error"] = f"Internal error: {exc}"
        error_response["traceback"] = traceback.format_exc()
        print_json(error_response)
        sys.exit(ExitCode.INTERNAL_ERROR)
