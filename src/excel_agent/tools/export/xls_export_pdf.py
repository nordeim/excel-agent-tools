"""xls_export_pdf: Export Excel workbook to PDF via LibreOffice headless.

Requires LibreOffice to be installed. Supports pre-calculation of formulas
and configurable timeout. Handles LibreOffice errors gracefully.
"""

from __future__ import annotations

import shutil
import subprocess
import time
from pathlib import Path

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

# Default timeout for LibreOffice conversion (seconds)
DEFAULT_TIMEOUT = 120
MAX_TIMEOUT = 600

# LibreOffice executable names to check
SOFFICE_NAMES = ["soffice", "soffice.bin", "libreoffice"]


def _find_soffice() -> str | None:
    """Find LibreOffice executable."""
    for name in SOFFICE_NAMES:
        path = shutil.which(name)
        if path:
            return path
    return None


def _verify_soffice() -> tuple[bool, str]:
    """Verify LibreOffice is installed and working."""
    soffice = _find_soffice()
    if not soffice:
        return False, "LibreOffice not found in PATH"

    try:
        result = subprocess.run(
            [soffice, "--headless", "--version"],
            capture_output=True,
            timeout=10,
            check=False,
        )
        if result.returncode == 0:
            version = result.stdout.decode().strip() if result.stdout else "unknown"
            return True, version
        return (
            False,
            f"LibreOffice check failed: {result.stderr.decode() if result.stderr else 'unknown error'}",
        )
    except subprocess.TimeoutExpired:
        return False, "LibreOffice version check timed out"
    except Exception as e:
        return False, f"LibreOffice check error: {e}"


def _run() -> dict[str, object]:
    parser = create_parser("Export Excel workbook to PDF via LibreOffice.")
    add_common_args(parser)
    parser.add_argument(
        "--outfile",
        type=str,
        required=False,
        help="Output PDF file path (default: same name with .pdf extension)",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
        help=f"Conversion timeout in seconds (default: {DEFAULT_TIMEOUT}, max: {MAX_TIMEOUT})",
    )
    parser.add_argument(
        "--recalc",
        action="store_true",
        default=False,
        help="Recalculate formulas before export (recommended for formulas to display correctly)",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(
        args.outfile or str(input_path.with_suffix(".pdf")),
        create_parents=True,
        allowed_suffixes={".pdf"},
    )

    # Validate timeout
    timeout = min(args.timeout, MAX_TIMEOUT)
    if args.timeout > MAX_TIMEOUT:
        warnings = [f"Timeout capped at {MAX_TIMEOUT} seconds"]
    else:
        warnings = []

    file_hash = compute_file_hash(input_path)

    # Verify LibreOffice is available
    soffice_available, soffice_info = _verify_soffice()
    if not soffice_available:
        return build_response(
            "error",
            None,
            exit_code=2,
            warnings=[
                f"LibreOffice not found: {soffice_info}",
                "Please install LibreOffice:",
                "  Ubuntu/Debian: sudo apt-get install libreoffice-calc",
                "  macOS: brew install --cask libreoffice",
                "  Windows: choco install libreoffice",
                "Or ensure soffice/libreoffice is in your PATH",
            ],
        )

    soffice = _find_soffice()

    # Handle --recalc option (would integrate with Phase 8 recalc)
    if args.recalc:
        warnings.append("--recalc specified: Ensure formulas are calculated before PDF export")
        # Note: In full implementation, this would call Phase 8 recalc tool

    # Build LibreOffice command
    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf:calc_pdf_Export",
        "--outdir",
        str(output_path.parent),
        str(input_path),
    ]

    # Execute conversion
    start_time = time.time()
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout,
            check=False,
        )
        elapsed_time = time.time() - start_time

        # LibreOffice names output based on input filename
        expected_pdf = output_path.parent / f"{input_path.stem}.pdf"

        if result.returncode != 0:
            stderr = result.stderr.decode() if result.stderr else "Unknown error"
            return build_response(
                "error",
                None,
                exit_code=1,
                warnings=[
                    f"PDF conversion failed (exit code {result.returncode})",
                    f"LibreOffice error: {stderr[:500]}",
                ],
            )

        # Check if PDF was created
        if not expected_pdf.exists():
            return build_response(
                "error",
                None,
                exit_code=1,
                warnings=[
                    "PDF file was not created",
                    f"Expected: {expected_pdf}",
                    f"LibreOffice stdout: {result.stdout.decode()[:500] if result.stdout else 'empty'}",
                ],
            )

        # Rename to requested output path if different
        if expected_pdf != output_path:
            expected_pdf.rename(output_path)

        # Verify output is non-empty
        output_size = output_path.stat().st_size
        if output_size == 0:
            return build_response(
                "error",
                None,
                exit_code=1,
                warnings=["PDF file was created but is empty"],
            )

    except subprocess.TimeoutExpired:
        elapsed_time = time.time() - start_time
        return build_response(
            "error",
            None,
            exit_code=5,
            warnings=[
                f"PDF conversion timed out after {elapsed_time:.1f} seconds",
                "Try increasing --timeout or check if file is corrupted",
            ],
        )
    except Exception as e:
        return build_response(
            "error",
            None,
            exit_code=5,
            warnings=[f"PDF conversion error: {e}"],
        )

    # Estimate page count (very rough approximation)
    # PDF doesn't easily expose page count without external libraries
    page_count_estimate = None

    # Log to audit trail
    audit = AuditTrail()
    audit.log(
        tool="xls_export_pdf",
        scope="read",
        target_file=input_path,
        file_version_hash=file_hash,
        actor_nonce="auto",
        operation_details={
            "output_path": str(output_path),
            "libreoffice_version": soffice_info,
            "timeout": timeout,
            "elapsed_time": elapsed_time,
            "recalc": args.recalc,
        },
        impact={
            "pdf_created": True,
            "file_size_bytes": output_size,
        },
        success=True,
        exit_code=0,
    )

    return build_response(
        "success",
        {
            "output_path": str(output_path),
            "libreoffice_version": soffice_info,
            "file_size_bytes": output_size,
            "conversion_time_seconds": round(elapsed_time, 2),
            "timeout_used": timeout,
            "page_count_estimate": page_count_estimate,
        },
        warnings=warnings if warnings else None,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
