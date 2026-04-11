"""xls_add_image: Insert images with aspect ratio preservation.

Supports PNG, JPEG, BMP, GIF formats. Optionally resizes while maintaining
aspect ratio. Warns on large files.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

from excel_agent.core.edit_session import EditSession
from excel_agent.governance.audit_trail import AuditTrail
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    check_macro_contract,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response

# Supported image formats
SUPPORTED_FORMATS = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}

# File size warnings (bytes)
WARNING_SIZE_1MB = 1 * 1024 * 1024
WARNING_SIZE_5MB = 5 * 1024 * 1024


def _get_image_info(image_path: Path) -> dict:
    """Get image information using PIL."""
    with PILImage.open(image_path) as img:
        return {
            "width": img.width,
            "height": img.height,
            "format": img.format,
            "mode": img.mode,
        }


def _calculate_dimensions(
    original_width: int,
    original_height: int,
    requested_width: int | None,
    requested_height: int | None,
) -> tuple[int, int]:
    """Calculate new dimensions while preserving aspect ratio.

    If only one dimension is provided, calculate the other to preserve ratio.
    If neither is provided, use original dimensions.
    If both are provided, use requested dimensions (may distort aspect).
    """
    if requested_width is None and requested_height is None:
        return original_width, original_height

    aspect_ratio = original_width / original_height

    if requested_width is not None and requested_height is not None:
        # Both specified - use as-is (may distort)
        return requested_width, requested_height
    elif requested_width is not None:
        # Only width specified - calculate height
        return requested_width, int(requested_width / aspect_ratio)
    else:
        # Only height specified - calculate width
        return int(requested_height * aspect_ratio), requested_height


def _run() -> dict[str, object]:
    parser = create_parser("Insert images into worksheets.")
    add_common_args(parser)
    parser.add_argument(
        "--image-path",
        type=str,
        required=True,
        help="Path to image file (PNG, JPEG, BMP, GIF)",
    )
    parser.add_argument(
        "--position",
        type=str,
        required=True,
        help='Anchor cell (e.g., "A1")',
    )
    parser.add_argument(
        "--width",
        type=int,
        default=None,
        help="Width in pixels (optional, preserves aspect if only one dimension set)",
    )
    parser.add_argument(
        "--height",
        type=int,
        default=None,
        help="Height in pixels (optional, preserves aspect if only one dimension set)",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    image_path = Path(args.image_path).resolve()
    output_path = validate_output_path(
        args.output or str(input_path),
        create_parents=True,
    )

    # Check for macro loss warning
    macro_warning = check_macro_contract(input_path, output_path)
    warnings = [macro_warning] if macro_warning else []

    # Validate image path
    if not image_path.exists():
        return build_response(
            "error",
            None,
            exit_code=2,
            warnings=[f"Image file not found: {image_path}"],
        )

    # Check file extension
    ext = image_path.suffix.lower()
    if ext not in SUPPORTED_FORMATS:
        return build_response(
            "error",
            None,
            exit_code=1,
            warnings=[
                f"Unsupported image format: {ext}",
                f"Supported formats: {', '.join(sorted(SUPPORTED_FORMATS))}",
            ],
        )

    # Check file size
    file_size = image_path.stat().st_size
    if file_size > WARNING_SIZE_5MB:
        warnings.append(
            f"Image is very large ({file_size / 1024 / 1024:.1f}MB). "
            "This may impact workbook performance."
        )
    elif file_size > WARNING_SIZE_1MB:
        warnings.append(
            f"Image is large ({file_size / 1024 / 1024:.1f}MB). "
            "Consider compressing before insertion."
        )

    # Get image info and calculate dimensions
    try:
        img_info = _get_image_info(image_path)
        new_width, new_height = _calculate_dimensions(
            img_info["width"],
            img_info["height"],
            args.width,
            args.height,
        )
    except Exception as e:
        return build_response(
            "error",
            None,
            exit_code=1,
            warnings=[f"Failed to process image: {e}"],
        )

    # Use EditSession for proper locking and save semantics
    session = EditSession.prepare(input_path, output_path)

    with session:
        wb = session.workbook
        ws = wb[args.sheet] if args.sheet else wb.active

        # Validate position
        try:
            from openpyxl.utils import coordinate_to_tuple

            coordinate_to_tuple(args.position)
        except Exception:
            return build_response(
                "error",
                None,
                exit_code=1,
                warnings=[f"Invalid position: {args.position}"],
            )

        # Create and add image
        try:
            img = XLImage(str(image_path))
            img.width = new_width
            img.height = new_height
            ws.add_image(img, args.position)
        except Exception as e:
            return build_response(
                "error",
                None,
                exit_code=5,
                warnings=[f"Failed to insert image: {e}"],
            )

        # Capture hashes before exiting context
        version_hash = session.version_hash
        file_hash = session.file_hash

        # EditSession handles save automatically on exit

    # Log to audit trail (after successful save)
    audit = AuditTrail()
    audit.log(
        tool="xls_add_image",
        scope="structure:modify",
        target_file=input_path,
        file_version_hash=file_hash,
        actor_nonce="auto",
        operation_details={
            "image_path": str(image_path),
            "position": args.position,
            "original_size": f"{img_info['width']}x{img_info['height']}",
            "final_size": f"{new_width}x{new_height}",
            "sheet": ws.title,
        },
        impact={
            "image_inserted": True,
            "file_size_kb": file_size / 1024,
        },
        success=True,
        exit_code=0,
    )

    return build_response(
        "success",
        {
            "image_path": str(image_path),
            "position": args.position,
            "sheet": ws.title,
            "original_width": img_info["width"],
            "original_height": img_info["height"],
            "final_width": new_width,
            "final_height": new_height,
            "format": img_info["format"],
            "file_size_kb": round(file_size / 1024, 2),
        },
        workbook_version=version_hash,
        warnings=warnings if warnings else None,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
