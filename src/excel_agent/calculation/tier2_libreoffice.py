"""
Tier 2 calculation engine: LibreOffice headless.

Provides full-fidelity recalculation by opening the workbook in
LibreOffice, which recalculates all formulas on load, then
re-saving as .xlsx.

Command pattern:
soffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" \
    --outdir <dir> <file>

This forces a complete recalculation. All 500+ Excel functions
are supported. Requires LibreOffice to be installed.
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import time
from pathlib import Path

from excel_agent.calculation.tier1_engine import CalculationResult

logger = logging.getLogger(__name__)

_COMMON_SOFFICE_PATHS = [
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/usr/local/bin/soffice",
    "/snap/bin/libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def _find_soffice() -> str | None:
    """Find the soffice binary on the system."""
    # Check PATH first
    soffice = shutil.which("soffice")
    if soffice:
        return soffice
    soffice = shutil.which("libreoffice")
    if soffice:
        return soffice
    # Check common installation paths
    for path in _COMMON_SOFFICE_PATHS:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path
    return None


class Tier2Calculator:
    """Full-fidelity recalculation via LibreOffice headless.

    Usage::

        calc = Tier2Calculator()
        if calc.is_available():
            result = calc.recalculate(Path("in.xlsx"), Path("out.xlsx"))
    """

    def __init__(self, *, soffice_path: str | None = None) -> None:
        if soffice_path:
            self._soffice = soffice_path
        else:
            self._soffice = _find_soffice()

    def is_available(self) -> bool:
        """Check if LibreOffice is installed and accessible."""
        if not self._soffice:
            return False
        try:
            result = subprocess.run(
                [self._soffice, "--headless", "--version"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            return result.returncode == 0
        except (OSError, subprocess.TimeoutExpired):
            return False

    def get_version(self) -> str:
        """Get LibreOffice version string."""
        if not self._soffice:
            return "not installed"
        try:
            result = subprocess.run(
                [self._soffice, "--headless", "--version"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            return result.stdout.strip() or "unknown"
        except (OSError, subprocess.TimeoutExpired):
            return "unavailable"

    def recalculate(
        self,
        workbook_path: Path,
        output_path: Path,
        *,
        timeout: int = 120,
    ) -> CalculationResult:
        """Recalculate a workbook via LibreOffice headless.

        Opens the workbook in LibreOffice (which forces a full recalc),
        then saves it as .xlsx.

        Args:
            workbook_path: Input workbook path.
            output_path: Where to write the recalculated workbook.
            timeout: Max seconds to wait for LibreOffice (default: 120).

        Returns:
            CalculationResult with timing info.
        """
        result = CalculationResult(engine="tier2_libreoffice")
        start = time.monotonic()

        if not self._soffice:
            result.errors.append(
                "LibreOffice not found. Install with: apt-get install libreoffice-calc"
            )
            result.error_count = 1
            return result

        output_dir = output_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            # Use a user profile to avoid locking issues with concurrent runs
            user_profile = output_dir / f".lo_profile_{os.getpid()}"
            env = os.environ.copy()
            env["HOME"] = str(user_profile)

            cmd = [
                self._soffice,
                "--headless",
                "--norestore",
                f"-env:UserInstallation=file://{user_profile}",
                "--convert-to",
                'xlsx:"Calc MS Excel 2007 XML"',
                "--outdir",
                str(output_dir),
                str(workbook_path.resolve()),
            ]

            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                env=env,
            )

            if proc.returncode != 0:
                result.errors.append(f"LibreOffice exited with code {proc.returncode}")
                if proc.stderr:
                    result.errors.append(proc.stderr[:500])
                result.error_count = 1
            else:
                # LibreOffice outputs to outdir with the same stem + .xlsx
                lo_output = output_dir / f"{workbook_path.stem}.xlsx"
                if lo_output.exists() and lo_output != output_path:
                    shutil.move(str(lo_output), str(output_path))
                result.output_path = str(output_path)

            # Clean up temp profile
            if user_profile.exists():
                shutil.rmtree(user_profile, ignore_errors=True)

        except subprocess.TimeoutExpired:
            result.errors.append(f"LibreOffice timed out after {timeout}s")
            result.error_count = 1
        except OSError as exc:
            result.errors.append(f"Failed to execute LibreOffice: {exc}")
            result.error_count = 1

        result.recalc_time_ms = (time.monotonic() - start) * 1000
        logger.info("Tier2 recalc: %.1fms, errors=%d", result.recalc_time_ms, result.error_count)
        return result
