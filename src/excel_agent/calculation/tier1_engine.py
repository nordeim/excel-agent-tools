"""
Tier 1 calculation engine: in-process via the formulas library.

The formulas library compiles Excel workbooks to Python and executes
without the Excel COM server. It supports 483 out of 536 Excel
functions (90.1% coverage as of v1.3.4).

Key API sequence:
xl_model = formulas.ExcelModel().loads(path).finish()
xl_model.calculate()
xl_model.write(dirpath=output_dir)

Limitation: The formulas library calculates from the file on disk —
it cannot recalculate after in-memory modifications via openpyxl.
The workflow must be: save changes → run Tier 1 → reload.

For circular references, add circular=True to finish().
Guard against XlError for #DIV/0!, #REF!, etc.
"""

from __future__ import annotations

import logging
import shutil
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)


@dataclass
class CalculationResult:
    """Result of a calculation engine run."""

    formula_count: int = 0
    calculated_count: int = 0
    error_count: int = 0
    unsupported_functions: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    recalc_time_ms: float = 0.0
    engine: str = "tier1_formulas"
    output_path: str = ""

    def to_dict(self) -> dict[str, object]:
        return {
            "formula_count": self.formula_count,
            "calculated_count": self.calculated_count,
            "error_count": self.error_count,
            "unsupported_functions": self.unsupported_functions,
            "errors": self.errors[:20],
            "recalc_time_ms": round(self.recalc_time_ms, 1),
            "engine": self.engine,
            "output_path": self.output_path,
        }


class Tier1Calculator:
    """In-process Excel calculation via the formulas library.

    Usage::

        calc = Tier1Calculator(Path("workbook.xlsx"))
        result = calc.calculate()
        if result.unsupported_functions:
            # Fall back to Tier 2
            ...
    """

    def __init__(self, workbook_path: Path) -> None:
        self._path = workbook_path.resolve()

    def calculate(
        self,
        output_path: Path | None = None,
        *,
        circular: bool = False,
    ) -> CalculationResult:
        """Calculate all formulas in the workbook.

        Args:
            output_path: Where to write the recalculated workbook.
                If None, writes to a temp directory.
            circular: Set True for workbooks with circular references.

        Returns:
            CalculationResult with stats and any errors.
        """
        import formulas

        result = CalculationResult()
        start = time.monotonic()

        try:
            xl_model = formulas.ExcelModel().loads(str(self._path)).finish(circular=circular)

            # Count formulas in the model
            try:
                nodes = xl_model.dsp.data_nodes
                formula_count = sum(
                    1 for key, val in nodes.items() if isinstance(key, str) and "!" in key
                )
                result.formula_count = formula_count
            except Exception:
                pass

            sol = xl_model.calculate()

            # Check solution for errors
            if sol is not None:
                for key, val in sol.items():
                    if hasattr(val, "value"):
                        v = val.value
                        if hasattr(v, "flat"):
                            for item in v.flat:
                                if _is_xl_error(item):
                                    result.errors.append(f"{key}: {item}")
                                    result.error_count += 1
                                else:
                                    result.calculated_count += 1
                        else:
                            if _is_xl_error(v):
                                result.errors.append(f"{key}: {v}")
                                result.error_count += 1
                            else:
                                result.calculated_count += 1
                    else:
                        result.calculated_count += 1

            # Write results
            if output_path is not None:
                out_dir = output_path.parent
                out_dir.mkdir(parents=True, exist_ok=True)
                written = xl_model.write(dirpath=str(out_dir))
                # formulas writes using the original filename uppercased
                # Move it to the requested output path
                for book_name, book_dict in written.items():
                    src_file = out_dir / book_name
                    if src_file.exists() and src_file != output_path:
                        shutil.move(str(src_file), str(output_path))
                    break
                result.output_path = str(output_path)
            else:
                with tempfile.TemporaryDirectory() as tmp:
                    xl_model.write(dirpath=tmp)
                    result.output_path = tmp

        except Exception as exc:
            error_msg = str(exc)
            if "not implemented" in error_msg.lower() or "not supported" in error_msg.lower():
                # Extract function name from error if possible
                result.unsupported_functions.append(error_msg[:100])
            else:
                result.errors.append(f"Tier1 error: {error_msg[:200]}")
                result.error_count += 1

        result.recalc_time_ms = (time.monotonic() - start) * 1000
        logger.info(
            "Tier1 calculation: %d formulas, %d calculated, %d errors in %.1fms",
            result.formula_count,
            result.calculated_count,
            result.error_count,
            result.recalc_time_ms,
        )
        return result


def _is_xl_error(value: object) -> bool:
    """Check if a value is an Excel error (XlError or error string)."""
    if value is None:
        return False
    val_str = str(value)
    return val_str.startswith("#") and val_str.endswith("!")
