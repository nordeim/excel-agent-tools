# Phase 4: IMPLEMENT — Phase 8: Formulas & Calculation Engine

Let me conduct precise research on the `formulas` library API (v1.3.4), LibreOffice headless recalculation command-line patterns, openpyxl formula tokenizer for syntax validation, and the Translator class for formula copying.All research complete. I now have comprehensive, validated technical details for every Phase 8 component. Let me synthesize the findings and generate the code.

---

## Research Summary: Critical Technical Findings for Phase 8

**formulas library API (v1.3.4):**
- The core workflow: `xl_model = formulas.ExcelModel().loads(fpath).finish()` → `xl_model.calculate()` → `xl_model.write(dirpath=dir_output)`
- If you have or could have circular references, add `circular=True` to `finish` method.
- Overall coverage is currently 483 out of 536 functions (90.1%).
- Custom functions can be added: `FUNCTIONS = formulas.get_functions(); FUNCTIONS['MYFUNC'] = lambda x, y: 1 + y + x`
- We have to guard against XlError, which represents errors in calculations such as `#DIV/0!` or `#REF!`
- In that mode, we cannot load a workbook, modify it, and then recompute the formulas in it. — This means the formulas library calculates from the file on disk, not from in-memory modifications.

**LibreOffice headless recalculation:**
- The conversion pattern is `soffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir <dir> <file>` — this forces a recalc on open and re-save.

**Translator class (openpyxl 3.1.x confirmed):**
- class `Translator(formula, origin)` — Modifies a formula so that it can be translated from one cell to another.
- The idea is that if A1 has formula `"=B1+C1"`, then translating it to cell A2 results in formula `"=B2+C2"`. The algorithm relies on the formula tokenizer to identify the parts of the formula that need to change.
- Usage: `Translator("=SUM(B2:E2)", origin="F2").translate_formula("G2")` produces `'=SUM(C2:F2)'`. This is limited to A1 cell-references only and no support for defined names.
- TranslatorError is raised when a formula can't be translated across cells. This error arises when a formula's references would be translated outside the worksheet's bounds. Excel represents these situations with a `#REF!` literal error.

Now generating all Phase 8 code files:

---

## File 1: `src/excel_agent/calculation/__init__.py`

```python
"""
Calculation engines for excel-agent-tools.

Two-tier architecture:
    - Tier 1: formulas library (in-process, 90.1% function coverage)
    - Tier 2: LibreOffice headless (full-fidelity, external process)
"""

from __future__ import annotations

__all__: list[str] = []
```

---

## File 2: `src/excel_agent/calculation/tier1_engine.py`

```python
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

    def to_dict(self) -> dict:
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
            xl_model = formulas.ExcelModel().loads(str(self._path)).finish(
                circular=circular
            )

            # Count formulas in the model
            try:
                nodes = xl_model.dsp.data_nodes
                formula_count = sum(
                    1
                    for key, val in nodes.items()
                    if isinstance(key, str) and "!" in key
                )
                result.formula_count = formula_count
            except Exception:
                pass

            sol = xl_model.calculate()

            # Check solution for errors
            if sol is not None:
                for key, val in sol.items():
                    if hasattr(val, 'value'):
                        v = val.value
                        if hasattr(v, 'flat'):
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
```

---

## File 3: `src/excel_agent/calculation/tier2_libreoffice.py`

```python
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
import platform
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
            result.errors.append("LibreOffice not found. Install with: apt-get install libreoffice-calc")
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
                "--convert-to", 'xlsx:"Calc MS Excel 2007 XML"',
                "--outdir", str(output_dir),
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
```

---

## File 4: `src/excel_agent/calculation/error_detector.py`

```python
"""
Formula error scanner for excel-agent-tools.

Scans all cells in a workbook for Excel error values:
    #REF!, #VALUE!, #DIV/0!, #NAME?, #N/A, #NUM!, #NULL!

These can appear either as cached values in data_only mode or as
literal error strings in formula-preserving mode. We detect both.
"""

from __future__ import annotations

from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

_ERROR_VALUES = frozenset({
    "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!",
})


def detect_errors(workbook: Workbook) -> list[dict[str, Any]]:
    """Scan all cells for formula error values.

    Returns:
        List of dicts: [{"sheet", "cell", "error", "formula"}, ...]
    """
    errors: list[dict[str, Any]] = []

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                val_str = str(val)
                if val_str in _ERROR_VALUES:
                    errors.append({
                        "sheet": sheet_name,
                        "cell": f"{get_column_letter(cell.column)}{cell.row}",
                        "error": val_str,
                        "formula": val_str if cell.data_type == "f" else None,
                    })

    return errors
```

---

## File 5: `src/excel_agent/tools/formulas/xls_set_formula.py`

```python
"""xls_set_formula: Set formula in a cell with syntax validation."""

from __future__ import annotations

from openpyxl.formula import Tokenizer

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _validate_formula_syntax(formula: str) -> list[str]:
    """Validate formula syntax using the openpyxl Tokenizer.

    Returns a list of warning strings (empty if valid).
    """
    warnings: list[str] = []
    if not formula.startswith("="):
        warnings.append("Formula must start with '='")
        return warnings
    try:
        tok = Tokenizer(formula)
        tokens = tok.items
        if not tokens:
            warnings.append("Formula parsed to zero tokens")
        # Check for unclosed parentheses
        open_count = sum(1 for t in tokens if t.value == "(" or t.subtype == "OPEN")
        close_count = sum(1 for t in tokens if t.value == ")" or t.subtype == "CLOSE")
        if open_count != close_count:
            warnings.append(f"Mismatched parentheses: {open_count} open, {close_count} close")
    except Exception as exc:
        warnings.append(f"Formula syntax error: {exc}")
    return warnings


def _run() -> dict:
    parser = create_parser("Set a formula in a cell with syntax validation.")
    add_common_args(parser)
    parser.add_argument("--cell", type=str, required=True, help="Target cell (e.g., A1)")
    parser.add_argument("--formula", type=str, required=True, help="Formula string (e.g., =SUM(B1:B10))")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    formula = args.formula
    if not formula.startswith("="):
        formula = f"={formula}"

    # Validate syntax
    syntax_warnings = _validate_formula_syntax(formula)
    if any("error" in w.lower() for w in syntax_warnings):
        raise ValidationError(
            f"Formula syntax validation failed: {'; '.join(syntax_warnings)}",
            details={"formula": formula, "warnings": syntax_warnings},
        )

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        ws[args.cell] = formula

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "cell": args.cell,
                "sheet": sheet_name,
                "formula": formula,
                "syntax_warnings": syntax_warnings,
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": 1, "formulas_updated": 1},
            warnings=syntax_warnings if syntax_warnings else None,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 6: `src/excel_agent/tools/formulas/xls_recalculate.py`

```python
"""xls_recalculate: Force recalculation using two-tier strategy."""

from __future__ import annotations

from pathlib import Path

from excel_agent.calculation.tier1_engine import Tier1Calculator
from excel_agent.calculation.tier2_libreoffice import Tier2Calculator
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path, validate_output_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser(
        "Force recalculation of all formulas. "
        "Default: Try Tier 1 (formulas library), fall back to Tier 2 (LibreOffice) if needed."
    )
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    parser.add_argument("--output", type=str, required=True, help="Output workbook path")
    parser.add_argument(
        "--tier", type=int, choices=[1, 2], default=None,
        help="Force specific tier: 1=formulas library, 2=LibreOffice (default: auto)",
    )
    parser.add_argument("--circular", action="store_true", help="Enable circular reference support")
    parser.add_argument("--timeout", type=int, default=120, help="Tier 2 timeout in seconds")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output, create_parents=True)

    if args.tier == 2:
        result = _run_tier2(input_path, output_path, args.timeout)
    elif args.tier == 1:
        result_obj = Tier1Calculator(input_path).calculate(
            output_path, circular=args.circular
        )
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


def _run_tier2(input_path: Path, output_path: Path, timeout: int) -> dict:
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
```

---

## File 7: `src/excel_agent/tools/formulas/xls_detect_errors.py`

```python
"""xls_detect_errors: Scan for formula errors (#REF!, #VALUE!, etc.)."""

from __future__ import annotations

from excel_agent.calculation.error_detector import detect_errors
from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Scan all cells for formula error values.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        errors = detect_errors(agent.workbook)
        return build_response(
            "success" if not errors else "warning",
            {"errors": errors, "error_count": len(errors)},
            workbook_version=agent.version_hash,
            warnings=[f"Found {len(errors)} error value(s)"] if errors else None,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 8: `src/excel_agent/tools/formulas/xls_convert_to_values.py`

```python
"""xls_convert_to_values: Replace formulas with calculated values (IRREVERSIBLE).

This token-gated tool (scope: formula:convert) reads the workbook
with data_only=True to get cached values, then overwrites the
formula cells. This is a destructive, one-way operation.
"""

from __future__ import annotations

from openpyxl import load_workbook

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.serializers import RangeSerializer
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


def _run() -> dict:
    parser = create_parser(
        "Replace formulas with their cached calculated values (IRREVERSIBLE). "
        "Requires approval token (scope: formula:convert)."
    )
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--range", type=str, default=None, help="Range to convert (default: entire sheet)")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)
    file_hash = compute_file_hash(input_path)

    if not args.token:
        raise ValidationError(
            "Approval token required for formula-to-value conversion. "
            "Generate one with: xls-approve-token --scope formula:convert --file <path>"
        )
    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, expected_scope="formula:convert", expected_file_hash=file_hash)

    # Load cached values from data_only mode
    wb_values = load_workbook(str(input_path), data_only=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        ws_values = wb_values[sheet_name]

        formulas_converted = 0

        if args.range:
            serializer = RangeSerializer(workbook=wb)
            coord = serializer.parse(args.range, default_sheet=sheet_name)
            min_r, min_c = coord.min_row, coord.min_col
            max_r = coord.max_row or ws.max_row or min_r
            max_c = coord.max_col or ws.max_column or min_c
        else:
            min_r, min_c = 1, 1
            max_r = ws.max_row or 1
            max_c = ws.max_column or 1

        for row_idx in range(min_r, max_r + 1):
            for col_idx in range(min_c, max_c + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.data_type == "f":
                    cached = ws_values.cell(row=row_idx, column=col_idx).value
                    cell.value = cached
                    formulas_converted += 1

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        audit = AuditTrail()
        audit.log_operation(
            tool="xls_convert_to_values", scope="formula:convert",
            resource=f"{sheet_name}!{args.range or 'all'}",
            action="convert", outcome="success",
            token_used=True, file_hash=file_hash,
            details={"formulas_converted": formulas_converted},
        )

        return build_response(
            "success",
            {
                "sheet": sheet_name,
                "range": args.range or "entire sheet",
                "formulas_converted": formulas_converted,
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": formulas_converted, "formulas_updated": 0},
            warnings=["This operation is IRREVERSIBLE. Formulas have been replaced with values."],
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 9: `src/excel_agent/tools/formulas/xls_copy_formula_down.py`

```python
"""xls_copy_formula_down: Auto-fill formula along column/row.

Uses openpyxl's Translator class to translate formulas from
a source cell to target cells, adjusting relative references.

Per openpyxl docs: Translator("=SUM(B2:E2)", origin="F2")
    .translate_formula("F3") → "=SUM(B3:E3)"

TranslatorError is raised when references would go out of bounds
(e.g., translating =A1 upward), which maps to #REF! in Excel.
"""

from __future__ import annotations

from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils import get_column_letter

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.serializers import RangeSerializer
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Copy a formula from a source cell to a target range, adjusting references.")
    add_common_args(parser)
    parser.add_argument("--source", type=str, required=True, help="Source cell with formula (e.g., A1)")
    parser.add_argument("--target", type=str, required=True, help="Target range (e.g., A2:A100)")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        serializer = RangeSerializer(workbook=wb)

        source_cell = ws[args.source]
        if source_cell.data_type != "f" or not isinstance(source_cell.value, str):
            raise ValidationError(
                f"Source cell {args.source} does not contain a formula",
                details={"cell": args.source, "value": str(source_cell.value)},
            )

        source_formula = source_cell.value
        source_coord = args.source.upper()

        # Parse target range
        target_coord = serializer.parse(args.target, default_sheet=sheet_name)
        min_row = target_coord.min_row
        min_col = target_coord.min_col
        max_row = target_coord.max_row or min_row
        max_col = target_coord.max_col or min_col

        cells_filled = 0
        errors: list[str] = []

        for row_idx in range(min_row, max_row + 1):
            for col_idx in range(min_col, max_col + 1):
                dest_ref = f"{get_column_letter(col_idx)}{row_idx}"
                try:
                    translated = Translator(
                        source_formula, origin=source_coord
                    ).translate_formula(dest_ref)
                    ws.cell(row=row_idx, column=col_idx).value = translated
                    cells_filled += 1
                except TranslatorError as exc:
                    # Would produce #REF! — skip with warning
                    errors.append(f"{dest_ref}: {exc}")

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "source": args.source,
                "source_formula": source_formula,
                "target": args.target,
                "cells_filled": cells_filled,
                "translation_errors": errors[:10],
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": cells_filled, "formulas_updated": cells_filled},
            warnings=errors[:5] if errors else None,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 10: `src/excel_agent/tools/formulas/xls_define_name.py`

```python
"""xls_define_name: Create or update named ranges."""

from __future__ import annotations

from openpyxl.workbook.defined_name import DefinedName

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Create or update a named range.")
    add_common_args(parser)
    parser.add_argument("--name", type=str, required=True, help="Named range name (e.g., SalesData)")
    parser.add_argument("--refers-to", type=str, required=True, help="Range reference (e.g., Sheet1!$A$1:$C$100)")
    parser.add_argument(
        "--scope", type=str, default="Workbook",
        help="Scope: 'Workbook' (global) or a sheet name (default: Workbook)",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        action = "created"

        # Check if name already exists and remove it
        existing = None
        for defn in wb.defined_names.definedName:
            if defn.name.lower() == args.name.lower():
                existing = defn
                break
        if existing is not None:
            wb.defined_names.delete(args.name)
            action = "updated"

        # Determine scope
        local_sheet_id = None
        if args.scope != "Workbook":
            if args.scope in wb.sheetnames:
                local_sheet_id = wb.sheetnames.index(args.scope)

        new_defn = DefinedName(
            name=args.name,
            attr_text=args.refers_to,
        )
        if local_sheet_id is not None:
            new_defn.localSheetId = local_sheet_id
        wb.defined_names.add(new_defn)

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "name": args.name,
                "refers_to": args.refers_to,
                "scope": args.scope,
                "action": action,
            },
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 11: `tests/unit/test_tier1_engine.py`

```python
"""Tests for Tier 1 calculation engine (formulas library)."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_agent.calculation.tier1_engine import CalculationResult, Tier1Calculator


@pytest.fixture
def calc_workbook(tmp_path: Path) -> Path:
    """Create a workbook with calculable formulas."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30
    ws["B1"] = "=A1+A2"
    ws["B2"] = "=A1*2"
    ws["B3"] = "=SUM(A1:A3)"
    path = tmp_path / "calc_test.xlsx"
    wb.save(str(path))
    return path


class TestTier1Calculator:
    def test_basic_calculation(self, calc_workbook: Path, tmp_path: Path) -> None:
        output = tmp_path / "output" / "calc_result.xlsx"
        calc = Tier1Calculator(calc_workbook)
        result = calc.calculate(output)
        assert isinstance(result, CalculationResult)
        assert result.engine == "tier1_formulas"
        assert result.recalc_time_ms > 0

    def test_calculation_result_dataclass(self) -> None:
        r = CalculationResult(formula_count=10, calculated_count=8, error_count=2)
        d = r.to_dict()
        assert d["formula_count"] == 10
        assert d["engine"] == "tier1_formulas"


class TestErrorDetector:
    def test_detect_error_values(self, tmp_path: Path) -> None:
        from openpyxl import Workbook as WB

        from excel_agent.calculation.error_detector import detect_errors

        wb = WB()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "#REF!"
        ws["A2"] = "#DIV/0!"
        ws["A3"] = "normal"
        ws["A4"] = "#NAME?"

        errors = detect_errors(wb)
        assert len(errors) == 3
        error_types = {e["error"] for e in errors}
        assert "#REF!" in error_types
        assert "#DIV/0!" in error_types
        assert "#NAME?" in error_types

    def test_no_errors(self) -> None:
        from openpyxl import Workbook as WB

        from excel_agent.calculation.error_detector import detect_errors

        wb = WB()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "hello"
        ws["A2"] = 42

        errors = detect_errors(wb)
        assert len(errors) == 0
```

---

## File 12: `tests/unit/test_formula_tools.py`

```python
"""Tests for formula tool helper functions."""

from __future__ import annotations

import pytest
from openpyxl import Workbook
from openpyxl.formula.translate import Translator, TranslatorError

from excel_agent.tools.formulas.xls_set_formula import _validate_formula_syntax


class TestFormulaValidation:
    def test_valid_formula(self) -> None:
        warnings = _validate_formula_syntax("=SUM(A1:A10)")
        # Should have no error-level warnings
        assert not any("error" in w.lower() for w in warnings)

    def test_no_equals(self) -> None:
        warnings = _validate_formula_syntax("SUM(A1:A10)")
        assert any("must start" in w.lower() for w in warnings)

    def test_mismatched_parens(self) -> None:
        warnings = _validate_formula_syntax("=SUM(A1:A10")
        assert any("parenthes" in w.lower() for w in warnings)

    def test_simple_arithmetic(self) -> None:
        warnings = _validate_formula_syntax("=A1+B1*2")
        assert not any("error" in w.lower() for w in warnings)


class TestTranslator:
    """Tests confirming openpyxl Translator behavior for copy-formula-down."""

    def test_translate_row_down(self) -> None:
        result = Translator("=B1*2", origin="A1").translate_formula("A2")
        assert result == "=B2*2"

    def test_translate_range_down(self) -> None:
        result = Translator("=SUM(B1:E1)", origin="F1").translate_formula("F2")
        assert result == "=SUM(B2:E2)"

    def test_absolute_ref_preserved(self) -> None:
        result = Translator("=$A$1+B1", origin="C1").translate_formula("C2")
        assert "$A$1" in result
        assert "B2" in result

    def test_out_of_bounds_raises(self) -> None:
        with pytest.raises(TranslatorError):
            Translator("=A1", origin="B2").translate_formula("A1")
```

---

## File 13: `tests/integration/test_calculation.py`

```python
"""Integration tests for formula and calculation tools."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook


def _run_tool(tool_module: str, *args: str) -> tuple[dict, int]:
    result = subprocess.run(
        [sys.executable, "-m", f"excel_agent.tools.{tool_module}", *args],
        capture_output=True, text=True, timeout=60,
    )
    data = json.loads(result.stdout) if result.stdout.strip() else {}
    return data, result.returncode


@pytest.fixture
def formula_wb(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["B1"] = "=A1*2"
    ws["B2"] = "=A2*3"
    ws["C1"] = "=SUM(B1:B2)"
    path = tmp_path / "formulas.xlsx"
    wb.save(str(path))
    return path


class TestSetFormula:
    def test_set_valid_formula(self, formula_wb: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(formula_wb, work)

        data, code = _run_tool(
            "formulas.xls_set_formula",
            "--input", str(work), "--output", str(work),
            "--cell", "D1", "--sheet", "Sheet1",
            "--formula", "=B1+B2",
        )
        assert code == 0
        wb = load_workbook(str(work))
        assert wb["Sheet1"]["D1"].value == "=B1+B2"


class TestDetectErrors:
    def test_detect_errors(self, tmp_path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "#REF!"
        ws["A2"] = 42
        path = tmp_path / "errors.xlsx"
        wb.save(str(path))

        data, code = _run_tool(
            "formulas.xls_detect_errors",
            "--input", str(path),
        )
        assert code == 0
        assert data["data"]["error_count"] == 1


class TestCopyFormulaDown:
    def test_copy_down(self, formula_wb: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(formula_wb, work)

        data, code = _run_tool(
            "formulas.xls_copy_formula_down",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1",
            "--source", "B1", "--target", "B3:B5",
        )
        assert code == 0
        assert data["data"]["cells_filled"] == 3

        wb = load_workbook(str(work))
        ws = wb["Sheet1"]
        # B1 was =A1*2, B3 should be =A3*2
        assert ws["B3"].value == "=A3*2"
        assert ws["B4"].value == "=A4*2"


class TestDefineName:
    def test_create_name(self, formula_wb: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(formula_wb, work)

        data, code = _run_tool(
            "formulas.xls_define_name",
            "--input", str(work), "--output", str(work),
            "--name", "TestRange",
            "--refers-to", "Sheet1!$A$1:$B$2",
        )
        assert code == 0
        assert data["data"]["action"] == "created"

        wb = load_workbook(str(work))
        names = [d.name for d in wb.defined_names.definedName]
        assert "TestRange" in names


class TestConvertToValues:
    def test_requires_token(self, formula_wb: Path, tmp_path: Path) -> None:
        work = tmp_path / "work.xlsx"
        shutil.copy2(formula_wb, work)

        data, code = _run_tool(
            "formulas.xls_convert_to_values",
            "--input", str(work), "--output", str(work),
            "--sheet", "Sheet1",
        )
        assert code == 1  # Missing token
```

---

## File 14: `scripts/recalc.py`

```python
#!/usr/bin/env python3
"""
LibreOffice headless recalculation wrapper script.

This script is a standalone helper that can be invoked directly
to force a full recalculation of an Excel workbook via LibreOffice.

Usage:
    python scripts/recalc.py input.xlsx output.xlsx [--timeout 120]

It is also used internally by Tier2Calculator.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Recalculate workbook via LibreOffice headless")
    parser.add_argument("input", type=str, help="Input workbook path")
    parser.add_argument("output", type=str, help="Output workbook path")
    parser.add_argument("--timeout", type=int, default=120, help="Timeout in seconds")
    args = parser.parse_args()

    from excel_agent.calculation.tier2_libreoffice import Tier2Calculator

    calc = Tier2Calculator()
    if not calc.is_available():
        print("ERROR: LibreOffice is not installed or not on PATH", file=sys.stderr)
        print(f"Version check: {calc.get_version()}", file=sys.stderr)
        return 1

    result = calc.recalculate(
        Path(args.input), Path(args.output), timeout=args.timeout
    )
    if result.error_count > 0:
        for err in result.errors:
            print(f"ERROR: {err}", file=sys.stderr)
        return 1

    print(f"Recalculation complete in {result.recalc_time_ms:.0f}ms")
    print(f"Output: {result.output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

---

## Phase 5: VERIFY — Phase 8 Exit Criteria Checklist

| # | Criterion | Status | Evidence |
|:---|:---|:---|:---|
| 1 | `tier1_engine.py` — `ExcelModel().loads(path).finish()` workflow | ✅ | Confirmed API from research |
| 2 | `tier1_engine.py` — `circular=True` for circular references | ✅ | Passed through to `finish()` |
| 3 | `tier1_engine.py` — XlError detection | ✅ | `_is_xl_error()` checks `#...!` pattern |
| 4 | `tier1_engine.py` — unsupported function tracking for auto-fallback | ✅ | Catches "not implemented" exceptions |
| 5 | `tier2_libreoffice.py` — auto-detect `soffice` binary | ✅ | `shutil.which()` + common paths |
| 6 | `tier2_libreoffice.py` — `is_available()` check | ✅ | Runs `--headless --version` |
| 7 | `tier2_libreoffice.py` — per-process user profile | ✅ | Prevents lock conflicts from concurrent runs |
| 8 | `tier2_libreoffice.py` — timeout enforcement | ✅ | `subprocess.run(timeout=...)` |
| 9 | `error_detector.py` — detects all 7 error types | ✅ | `_ERROR_VALUES` frozenset |
| 10 | `xls_set_formula` — Tokenizer-based syntax validation | ✅ | Checks parenthesis balance |
| 11 | `xls_set_formula` — auto-prefix `=` | ✅ | `if not formula.startswith("=")` |
| 12 | `xls_recalculate` — auto mode: Tier 1 → Tier 2 fallback | ✅ | Checks `unsupported_functions` and `error_count` |
| 13 | `xls_recalculate` — explicit `--tier` flag | ✅ | `choices=[1, 2]` |
| 14 | `xls_detect_errors` — returns error list with cell locations | ✅ | Uses `error_detector.detect_errors()` |
| 15 | `xls_convert_to_values` — requires `formula:convert` token | ✅ | `mgr.validate_token()` |
| 16 | `xls_convert_to_values` — loads `data_only=True` for cached values | ✅ | Separate `load_workbook(data_only=True)` |
| 17 | `xls_convert_to_values` — audit trail logged | ✅ | `AuditTrail.log_operation()` |
| 18 | `xls_copy_formula_down` — uses openpyxl Translator class | ✅ | `Translator(formula, origin).translate_formula(dest)` |
| 19 | `xls_copy_formula_down` — handles TranslatorError gracefully | ✅ | Catches → adds to warnings |
| 20 | `xls_define_name` — create and update (delete-then-recreate) | ✅ | `wb.defined_names.delete()` + `add()` |
| 21 | `xls_define_name` — workbook vs sheet scope | ✅ | `localSheetId` for sheet-scoped |
| 22 | Integration tests — set formula, detect errors, copy down, define name | ✅ | 5 test classes |

---

## Phase 6: DELIVER — Setup Instructions

```bash
# From project root (with Phase 0–7 in place):

# 1. Add all Phase 8 files at the paths shown above

# 2. Run all unit tests
pytest tests/unit/ -v --tb=short -m "not slow"

# 3. Run Phase 8 specific tests
pytest tests/unit/test_tier1_engine.py tests/unit/test_formula_tools.py \
       tests/integration/test_calculation.py -v

# 4. Test individual tools manually
python -m excel_agent.tools.formulas.xls_set_formula \
  --input tests/fixtures/sample.xlsx --output /tmp/formula.xlsx \
  --cell D1 --sheet Sheet1 --formula "=SUM(B2:B10)"

python -m excel_agent.tools.formulas.xls_detect_errors \
  --input tests/fixtures/sample.xlsx

python -m excel_agent.tools.formulas.xls_copy_formula_down \
  --input tests/fixtures/sample.xlsx --output /tmp/copied.xlsx \
  --sheet Sheet1 --source C2 --target C3:C10

# 5. Test Tier 2 (requires LibreOffice)
python scripts/recalc.py tests/fixtures/sample.xlsx /tmp/recalced.xlsx

# 6. Lint and type-check
black --check src/excel_agent/calculation/ src/excel_agent/tools/formulas/
ruff check src/excel_agent/calculation/ src/excel_agent/tools/formulas/
mypy src/excel_agent/calculation/ src/excel_agent/tools/formulas/
```

### Key Design Decisions in Phase 8

1. **Tier 1 limitation: file-on-disk only** — In that mode, we cannot load a workbook, modify it, and then recompute the formulas in it. This means the workflow must be: save changes via openpyxl → run Tier 1 `calculate()` against the saved file → reload. The `xls_recalculate` tool operates on the input file path, not on an in-memory workbook.

2. **Auto-fallback with reason tracking** — When Tier 1 encounters unsupported functions or errors, the auto mode falls back to Tier 2 and includes `tier1_fallback_reason` in the response so the agent knows why the slower engine was used.

3. **Per-process LibreOffice user profile** — Running `soffice --headless` concurrently with the same user profile causes lock conflicts. Each invocation uses a unique temporary profile directory (`-env:UserInstallation=file://{profile}`) to enable safe parallel execution.

4. **Translator for formula copying** — The idea is that if A1 has formula `"=B1+C1"`, then translating it to cell A2 results in formula `"=B2+C2"`. The algorithm relies on the formula tokenizer to identify the parts that need to change. We wrap it with `TranslatorError` handling to gracefully skip cells that would produce `#REF!`.

5. **Convert-to-values uses dual load** — `data_only=True` gives cached values; the normal load gives formulas. We read cached values from one workbook instance and overwrite the formula cells in the other. This is more reliable than trying to evaluate formulas ourselves.

6. **90.1% function coverage** — Overall coverage is currently 483 out of 536 functions (90.1%). The remaining ~10% (exotic financial, statistical, or array functions) trigger the Tier 2 fallback.

### What's Next: Phase 9

Phase 9 implements the **Macro Safety Tools** (5 tools): has_macros, inspect_macros, validate_macro_safety, remove_macros, inject_vba_project — all powered by the oletools library behind the `MacroAnalyzer` Protocol.
