"""xls_inject_vba_project: Inject pre-extracted vbaProject.bin (token-gated, pre-scan required).

Injects a VBA project from a pre-extracted .bin file into a target workbook.
This is an advanced operation that requires:
1. A pre-scanned vbaProject.bin file (extracted from a trusted source)
2. A valid approval token (scope: macro:inject)
3. Automatic safety scan of the .bin before injection

Security:
- ALWAYS runs xls_validate_macro_safety on the .bin before injection
- Denies injection if risk level is high/critical (unless --force flag)
- Logs full audit trail (but never logs VBA source code)
- Token required for this irreversible operation
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any

from excel_agent.core.macro_handler import get_analyzer
from excel_agent.core.version_hash import compute_file_hash
from excel_agent.governance.audit_trail import AuditTrail
from excel_agent.governance.token_manager import ApprovalTokenManager
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    add_governance_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _inject_vba(
    target_path: Path,
    vba_bin_path: Path,
    output_path: Path,
) -> dict[str, Any]:
    """Inject vbaProject.bin into target workbook."""
    # Read the vbaProject.bin
    with open(vba_bin_path, "rb") as f:
        vba_data = f.read()

    # Read the target workbook as a zip
    with zipfile.ZipFile(target_path, "r") as zin:
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zout:
            # Copy all existing files except vbaProject.bin
            for item in zin.infolist():
                if item.filename == "xl/vbaProject.bin":
                    continue
                data = zin.read(item.filename)
                zout.writestr(item, data)

            # Add the new vbaProject.bin
            vba_info = zipfile.ZipInfo(filename="xl/vbaProject.bin")
            zout.writestr(vba_info, vba_data)

            # Ensure [Content_Types].xml includes VBA content type
            # This is typically handled automatically by openpyxl when saving .xlsm
            # but we need to ensure the target becomes a macro-enabled file

    # Write to output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        f.write(buffer.getvalue())

    return {
        "vba_injected": True,
        "output_path": str(output_path),
        "vba_bin_size": len(vba_data),
    }


def _run() -> dict[str, object]:
    parser = create_parser(
        "Inject VBA project from .bin file into workbook (token-gated, pre-scan required)."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument(
        "--vba-bin",
        type=str,
        required=True,
        help="Path to vbaProject.bin file to inject",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Inject even if risk level is high/critical",
    )
    args = parser.parse_args()

    target_path = validate_input_path(args.input)
    vba_bin_path = validate_input_path(args.vba_bin)
    output_path = validate_output_path(
        args.output or str(target_path).replace(".xlsx", "_with_macros.xlsm"),
        create_parents=True,
    )

    file_hash = compute_file_hash(target_path)

    # Validate token
    if not args.token:
        raise ValidationError("Approval token required (--token)")

    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, "macro:inject", target_path)

    # ALWAYS pre-scan the vbaProject.bin before injection
    analyzer = get_analyzer()
    scan_result = analyzer.analyze(vba_bin_path)

    # Deny if risk level is high/critical (unless --force)
    if scan_result.risk_level in ("high", "critical") and not args.force:
        return build_response(
            "denied",
            {
                "risk_level": scan_result.risk_level,
                "risk_score": scan_result.risk_score,
                "auto_exec_functions": scan_result.auto_exec_functions,
            },
            exit_code=1,
            guidance="Run xls_validate_macro_safety --bin <path> to review risks. "
            "Use --force to inject despite risks (not recommended).",
        )

    # Inject VBA
    result = _inject_vba(target_path, vba_bin_path, output_path)

    # Log audit (never include VBA source code)
    audit = AuditTrail()
    audit.log(
        tool="xls_inject_vba_project",
        scope="macro:inject",
        target_file=target_path,
        file_version_hash=file_hash,
        actor_nonce=args.token[:32],
        operation_details={
            "output_path": str(output_path),
            "vba_bin_path": str(vba_bin_path),
            "risk_level": scan_result.risk_level,
            "risk_score": scan_result.risk_score,
        },
        impact={"vba_injected": True, "vba_bin_size": result["vba_bin_size"]},
        success=True,
        exit_code=0,
    )

    return build_response(
        "success",
        {
            **result,
            "risk_level": scan_result.risk_level,
            "risk_score": scan_result.risk_score,
        },
        warnings=[
            f"VBA project injected. Risk level: {scan_result.risk_level}",
            "Review macro safety before opening in Excel.",
        ],
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
