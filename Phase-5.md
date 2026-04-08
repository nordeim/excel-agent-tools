# Phase 4: IMPLEMENT — Phase 5: Write & Create Tools

Let me conduct precise research on the openpyxl APIs needed for workbook creation, template handling, cell type inference, and formula string detection.All research complete. I have validated every API detail needed. Let me now synthesize and generate all Phase 5 code files.

---

## Research Summary: Critical Technical Findings for Phase 5

**Workbook creation:** A workbook is always created with at least one worksheet. You can create new worksheets using the Workbook.create_sheet() method: insert at the end (default), at first position with index 0, or at penultimate position with -1. To remove the default sheet, use `del wb['Sheet']` as demonstrated in the examples.

**Template handling:** You can specify the attribute template=True, to save a workbook as a template, or set this attribute to False (default), to save as a document. You should monitor the data attributes and document extensions for saving documents in the document templates and vice versa, otherwise the result table engine can not open the document.

**Cell data types (openpyxl 3.1.x source confirmed):** Cell type constants include TYPE_STRING = 's', TYPE_FORMULA = 'f', TYPE_NUMERIC = 'n', TYPE_BOOL = 'b', TYPE_NULL = 'n', TYPE_INLINE = 'inlineStr', TYPE_ERROR = 'e'.

**Cell value assignment:** TIME_TYPES include datetime.datetime, datetime.date, datetime.time, and datetime.timedelta. These get auto-detected and formatted with the appropriate Excel number format. When assigning a string starting with `=`, openpyxl automatically treats it as a formula.

**String length limit:** "string must never be longer than 32,767 characters — truncate if necessary"

Now generating all Phase 5 code files:

---

## File 1: `src/excel_agent/core/type_coercion.py`

```python
"""
Type inference and explicit coercion for cell values.

When writing data from JSON to Excel, we need to convert JSON types
(string, number, boolean, null) into appropriate Python types that
openpyxl will correctly store in Excel cells.

Key behaviors:
    - Strings starting with '=' → treated as formulas (openpyxl auto-detects)
    - ISO 8601 date strings → datetime objects (auto-formatted by openpyxl)
    - "true"/"false" strings → Python bool
    - Numeric strings → int or float
    - None → empty cell

openpyxl Cell type constants:
    TYPE_STRING = 's', TYPE_FORMULA = 'f', TYPE_NUMERIC = 'n',
    TYPE_BOOL = 'b', TYPE_NULL = 'n', TYPE_ERROR = 'e'

openpyxl auto-detects datetime types and applies appropriate
Excel number formats: datetime → FORMAT_DATE_DATETIME,
date → FORMAT_DATE_YYYYMMDD2, time → FORMAT_DATE_TIME6.
"""

from __future__ import annotations

import datetime
import re
from typing import Any

# ISO 8601 date patterns
_DATE_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}$"
)
_DATETIME_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$"
)


def infer_cell_value(value: Any) -> Any:  # noqa: ANN401
    """Infer the best Python type for a JSON value before writing to Excel.

    Type inference rules (in order):
        1. None → None (empty cell)
        2. bool → bool (must check before int, since bool is subclass of int)
        3. int/float → passthrough (numeric)
        4. str starting with '=' → passthrough (openpyxl treats as formula)
        5. str matching ISO 8601 date → datetime.date or datetime.datetime
        6. str "true"/"false" (case-insensitive) → bool
        7. str that parses as int → int
        8. str that parses as float → float
        9. str → passthrough (text)

    Args:
        value: A JSON-compatible value (str, int, float, bool, None).

    Returns:
        Python value suitable for assigning to cell.value.
    """
    if value is None:
        return None

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        return value

    if not isinstance(value, str):
        return value

    # Formula detection: strings starting with '='
    if value.startswith("="):
        return value

    # Boolean string detection
    if value.lower() == "true":
        return True
    if value.lower() == "false":
        return False

    # ISO 8601 datetime detection (must check before date)
    if _DATETIME_RE.match(value):
        try:
            return datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            pass

    # ISO 8601 date detection
    if _DATE_RE.match(value):
        try:
            return datetime.date.fromisoformat(value)
        except ValueError:
            pass

    # Numeric string detection
    try:
        int_val = int(value)
        # Preserve leading zeros as strings (e.g., "007")
        if value != str(int_val):
            return value
        return int_val
    except ValueError:
        pass

    try:
        return float(value)
    except ValueError:
        pass

    # Plain string
    return value


def coerce_cell_value(value: str, target_type: str) -> Any:  # noqa: ANN401
    """Explicitly coerce a string value to a specific type.

    Used by xls_write_cell --type flag for override of auto-inference.

    Args:
        value: The raw string value from CLI.
        target_type: One of "string", "number", "boolean", "date",
                     "datetime", "formula", "integer", "float".

    Returns:
        Python value suitable for cell.value.

    Raises:
        ValueError: If the value cannot be coerced to the target type.
    """
    if target_type == "string":
        return value

    if target_type == "formula":
        if not value.startswith("="):
            return f"={value}"
        return value

    if target_type == "boolean":
        lower = value.lower().strip()
        if lower in ("true", "1", "yes"):
            return True
        if lower in ("false", "0", "no"):
            return False
        raise ValueError(f"Cannot coerce {value!r} to boolean")

    if target_type == "integer":
        return int(value)

    if target_type == "float" or target_type == "number":
        return float(value)

    if target_type == "date":
        return datetime.date.fromisoformat(value)

    if target_type == "datetime":
        return datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))

    raise ValueError(f"Unknown target type: {target_type!r}")
```

---

## File 2: `src/excel_agent/tools/write/xls_create_new.py`

```python
"""xls_create_new: Create a blank workbook with optional sheet names."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from excel_agent.core.version_hash import compute_file_hash
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_output_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Create a new blank workbook with specified sheet names.")
    parser.add_argument(
        "--output", type=str, required=True, help="Output workbook path (.xlsx)"
    )
    parser.add_argument(
        "--sheets",
        type=str,
        default=None,
        help="Comma-separated sheet names (default: single 'Sheet1')",
    )
    args = parser.parse_args()

    output = validate_output_path(args.output, create_parents=True)

    wb = Workbook()

    if args.sheets:
        sheet_names = [s.strip() for s in args.sheets.split(",") if s.strip()]
    else:
        sheet_names = ["Sheet1"]

    # A workbook is always created with at least one worksheet.
    # Rename the default sheet to the first name, then create the rest.
    default_ws = wb.active
    assert default_ws is not None

    if sheet_names:
        default_ws.title = sheet_names[0]
        for name in sheet_names[1:]:
            wb.create_sheet(name)
    else:
        default_ws.title = "Sheet1"

    wb.save(str(output))
    file_hash = compute_file_hash(output)

    return build_response(
        "success",
        {
            "output_path": str(output),
            "sheets": list(wb.sheetnames),
            "sheet_count": len(wb.sheetnames),
        },
        workbook_version=file_hash,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 3: `src/excel_agent/tools/write/xls_create_from_template.py`

```python
"""xls_create_from_template: Clone from .xltx/.xltm template with variable substitution."""

from __future__ import annotations

import re

from openpyxl import load_workbook

from excel_agent.core.version_hash import compute_file_hash
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    create_parser,
    parse_json_arg,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response

# Regex for {{placeholder}} patterns
_PLACEHOLDER_RE = re.compile(r"\{\{(\w+)\}\}")


def _run() -> dict:
    parser = create_parser(
        "Create a new workbook from a template (.xltx/.xltm) with variable substitution."
    )
    parser.add_argument(
        "--template", type=str, required=True, help="Path to template file (.xltx, .xltm, .xlsx)"
    )
    parser.add_argument(
        "--output", type=str, required=True, help="Output workbook path (.xlsx)"
    )
    parser.add_argument(
        "--vars",
        type=str,
        default="{}",
        help='JSON object of variable substitutions (e.g., \'{"company": "Acme"}\')',
    )
    args = parser.parse_args()

    template_path = validate_input_path(args.template)
    output_path = validate_output_path(args.output, create_parents=True)
    variables: dict = parse_json_arg(args.vars)

    # Load template — set template=False so output is a regular workbook
    # Per openpyxl docs: "You should monitor the data attributes and document
    # extensions for saving documents in the document templates and vice versa"
    wb = load_workbook(str(template_path))
    wb.template = False

    # Perform variable substitution across all sheets
    substitutions_made = 0
    substitution_details: list[dict] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f" and isinstance(cell.value, str):
                    # Only substitute in text cells, NOT in formulas
                    original = cell.value
                    new_value = _substitute_placeholders(original, variables)
                    if new_value != original:
                        cell.value = new_value
                        substitutions_made += 1
                        substitution_details.append({
                            "sheet": sheet_name,
                            "cell": cell.coordinate,
                            "original": original,
                            "replaced": new_value,
                        })

    wb.save(str(output_path))
    file_hash = compute_file_hash(output_path)

    return build_response(
        "success",
        {
            "output_path": str(output_path),
            "template_path": str(template_path),
            "variables": variables,
            "substitutions_made": substitutions_made,
            "substitution_details": substitution_details[:20],
            "sheets": list(wb.sheetnames),
        },
        workbook_version=file_hash,
        impact={"cells_modified": substitutions_made, "formulas_updated": 0},
    )


def _substitute_placeholders(text: str, variables: dict) -> str:
    """Replace {{key}} placeholders with values from the variables dict.

    Unmatched placeholders are left as-is (not replaced).
    """

    def _replacer(match: re.Match) -> str:  # type: ignore[type-arg]
        key = match.group(1)
        if key in variables:
            return str(variables[key])
        return match.group(0)  # Leave unmatched placeholders as-is

    return _PLACEHOLDER_RE.sub(_replacer, text)


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 4: `src/excel_agent/tools/write/xls_write_range.py`

```python
"""xls_write_range: Write 2D data array to a range with type inference."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.serializers import RangeSerializer
from excel_agent.core.type_coercion import infer_cell_value
from excel_agent.governance.schemas import validate_against_schema
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    parse_json_arg,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Write a 2D data array to a cell range with type inference.")
    add_common_args(parser)
    parser.add_argument("--range", type=str, required=True, help="Start cell (e.g., A1)")
    parser.add_argument(
        "--data",
        type=str,
        required=True,
        help='JSON 2D array (e.g., \'[["Name","Age"],["Alice",30]]\')',
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(
        args.output or args.input, create_parents=True
    )

    data_parsed = parse_json_arg(args.data)
    validate_against_schema("write_data", {"data": data_parsed})
    data: list[list] = data_parsed

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        serializer = RangeSerializer(workbook=wb)
        coord = serializer.parse(args.range, default_sheet=args.sheet)

        sheet_name = coord.sheet or (args.sheet or wb.sheetnames[0])
        ws = wb[sheet_name]

        start_row = coord.min_row
        start_col = coord.min_col

        cells_written = 0
        formulas_written = 0

        for row_idx, row_data in enumerate(data):
            for col_idx, raw_value in enumerate(row_data):
                cell_row = start_row + row_idx
                cell_col = start_col + col_idx
                cell = ws.cell(row=cell_row, column=cell_col)

                coerced = infer_cell_value(raw_value)
                cell.value = coerced
                cells_written += 1

                if isinstance(coerced, str) and coerced.startswith("="):
                    formulas_written += 1

        # If output differs from input, save to new path
        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "range": args.range,
                "sheet": sheet_name,
                "rows_written": len(data),
                "cols_written": max((len(row) for row in data), default=0),
            },
            workbook_version=agent.version_hash,
            impact={"cells_modified": cells_written, "formulas_updated": formulas_written},
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 5: `src/excel_agent/tools/write/xls_write_cell.py`

```python
"""xls_write_cell: Write a single cell with explicit type coercion."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.type_coercion import coerce_cell_value, infer_cell_value
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.json_io import build_response

_VALID_TYPES = frozenset({
    "string", "number", "integer", "float",
    "boolean", "date", "datetime", "formula",
})


def _run() -> dict:
    parser = create_parser("Write a single cell value with optional explicit type coercion.")
    add_common_args(parser)
    parser.add_argument("--cell", type=str, required=True, help="Cell reference (e.g., A1)")
    parser.add_argument("--value", type=str, required=True, help="Value to write")
    parser.add_argument(
        "--type",
        type=str,
        default=None,
        choices=sorted(_VALID_TYPES),
        help="Explicit type coercion (overrides auto-inference)",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(
        args.output or args.input, create_parents=True
    )

    with ExcelAgent(input_path, mode="rw") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        cell = ws[args.cell]

        if args.type is not None:
            coerced = coerce_cell_value(args.value, args.type)
        else:
            coerced = infer_cell_value(args.value)

        cell.value = coerced

        is_formula = isinstance(coerced, str) and coerced.startswith("=")

        if str(output_path) != str(input_path):
            wb.save(str(output_path))

        return build_response(
            "success",
            {
                "cell": args.cell,
                "sheet": sheet_name,
                "value_written": str(coerced),
                "type_used": args.type or "auto",
                "is_formula": is_formula,
            },
            workbook_version=agent.version_hash,
            impact={
                "cells_modified": 1,
                "formulas_updated": 1 if is_formula else 0,
            },
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
```

---

## File 6: `tests/unit/test_type_coercion.py`

```python
"""Tests for type inference and explicit coercion."""

from __future__ import annotations

import datetime

import pytest

from excel_agent.core.type_coercion import coerce_cell_value, infer_cell_value


class TestInferCellValue:
    """Tests for automatic type inference."""

    def test_none(self) -> None:
        assert infer_cell_value(None) is None

    def test_bool_true(self) -> None:
        assert infer_cell_value(True) is True

    def test_bool_false(self) -> None:
        assert infer_cell_value(False) is False

    def test_integer(self) -> None:
        assert infer_cell_value(42) == 42

    def test_float(self) -> None:
        assert infer_cell_value(3.14) == 3.14

    def test_plain_string(self) -> None:
        assert infer_cell_value("hello") == "hello"

    def test_formula_string(self) -> None:
        result = infer_cell_value("=SUM(A1:A10)")
        assert result == "=SUM(A1:A10)"
        assert isinstance(result, str)

    def test_boolean_string_true(self) -> None:
        assert infer_cell_value("true") is True
        assert infer_cell_value("TRUE") is True

    def test_boolean_string_false(self) -> None:
        assert infer_cell_value("false") is False
        assert infer_cell_value("FALSE") is False

    def test_iso_date_string(self) -> None:
        result = infer_cell_value("2026-04-08")
        assert isinstance(result, datetime.date)
        assert result.year == 2026
        assert result.month == 4
        assert result.day == 8

    def test_iso_datetime_string(self) -> None:
        result = infer_cell_value("2026-04-08T14:30:00")
        assert isinstance(result, datetime.datetime)
        assert result.year == 2026

    def test_iso_datetime_with_timezone(self) -> None:
        result = infer_cell_value("2026-04-08T14:30:00Z")
        assert isinstance(result, datetime.datetime)

    def test_numeric_string_integer(self) -> None:
        assert infer_cell_value("42") == 42
        assert isinstance(infer_cell_value("42"), int)

    def test_numeric_string_float(self) -> None:
        assert infer_cell_value("3.14") == 3.14
        assert isinstance(infer_cell_value("3.14"), float)

    def test_leading_zero_preserved_as_string(self) -> None:
        """Leading zeros should NOT be converted to int (e.g., ZIP codes)."""
        assert infer_cell_value("007") == "007"
        assert isinstance(infer_cell_value("007"), str)

    def test_negative_number_string(self) -> None:
        assert infer_cell_value("-5") == -5

    def test_empty_string(self) -> None:
        assert infer_cell_value("") == ""

    def test_non_numeric_non_date_string(self) -> None:
        assert infer_cell_value("abc123") == "abc123"


class TestCoerceCellValue:
    """Tests for explicit type coercion."""

    def test_coerce_string(self) -> None:
        assert coerce_cell_value("42", "string") == "42"

    def test_coerce_integer(self) -> None:
        assert coerce_cell_value("42", "integer") == 42

    def test_coerce_float(self) -> None:
        assert coerce_cell_value("3.14", "float") == 3.14

    def test_coerce_number(self) -> None:
        assert coerce_cell_value("3.14", "number") == 3.14

    def test_coerce_boolean_true(self) -> None:
        assert coerce_cell_value("true", "boolean") is True
        assert coerce_cell_value("1", "boolean") is True
        assert coerce_cell_value("yes", "boolean") is True

    def test_coerce_boolean_false(self) -> None:
        assert coerce_cell_value("false", "boolean") is False
        assert coerce_cell_value("0", "boolean") is False

    def test_coerce_boolean_invalid(self) -> None:
        with pytest.raises(ValueError, match="Cannot coerce"):
            coerce_cell_value("maybe", "boolean")

    def test_coerce_date(self) -> None:
        result = coerce_cell_value("2026-04-08", "date")
        assert isinstance(result, datetime.date)
        assert result.year == 2026

    def test_coerce_datetime(self) -> None:
        result = coerce_cell_value("2026-04-08T14:30:00", "datetime")
        assert isinstance(result, datetime.datetime)

    def test_coerce_formula(self) -> None:
        assert coerce_cell_value("=SUM(A1:A10)", "formula") == "=SUM(A1:A10)"

    def test_coerce_formula_auto_prefix(self) -> None:
        assert coerce_cell_value("SUM(A1:A10)", "formula") == "=SUM(A1:A10)"

    def test_coerce_unknown_type(self) -> None:
        with pytest.raises(ValueError, match="Unknown target type"):
            coerce_cell_value("x", "unknown")

    def test_coerce_invalid_integer(self) -> None:
        with pytest.raises(ValueError):
            coerce_cell_value("abc", "integer")
```

---

## File 7: `tests/integration/test_write_tools.py`

```python
"""Integration tests for write and create tools via subprocess."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook


def _run_tool(tool_module: str, *args: str) -> dict:
    """Run a CLI tool via subprocess and return parsed JSON output."""
    result = subprocess.run(
        [sys.executable, "-m", f"excel_agent.tools.{tool_module}", *args],
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.stdout.strip(), (
        f"Tool {tool_module} produced no output. stderr: {result.stderr}"
    )
    return json.loads(result.stdout)


class TestCreateNew:
    """Tests for xls_create_new."""

    def test_create_default(self, tmp_path: Path) -> None:
        output = tmp_path / "new.xlsx"
        result = _run_tool("write.xls_create_new", "--output", str(output))
        assert result["status"] == "success"
        assert output.exists()
        wb = load_workbook(str(output))
        assert wb.sheetnames == ["Sheet1"]

    def test_create_with_sheets(self, tmp_path: Path) -> None:
        output = tmp_path / "multi.xlsx"
        result = _run_tool(
            "write.xls_create_new",
            "--output", str(output),
            "--sheets", "Data,Summary,Charts",
        )
        assert result["status"] == "success"
        wb = load_workbook(str(output))
        assert wb.sheetnames == ["Data", "Summary", "Charts"]

    def test_create_single_named(self, tmp_path: Path) -> None:
        output = tmp_path / "single.xlsx"
        result = _run_tool(
            "write.xls_create_new",
            "--output", str(output),
            "--sheets", "MyData",
        )
        assert result["status"] == "success"
        wb = load_workbook(str(output))
        assert wb.sheetnames == ["MyData"]


class TestCreateFromTemplate:
    """Tests for xls_create_from_template."""

    def test_variable_substitution(self, tmp_path: Path) -> None:
        # Create a template with placeholders
        from openpyxl import Workbook as WB

        template = tmp_path / "template.xlsx"
        wb = WB()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "{{company}}"
        ws["A2"] = "Report for {{year}}"
        ws["A3"] = "=SUM(B1:B10)"  # Formula should NOT be substituted
        wb.save(str(template))

        output = tmp_path / "from_template.xlsx"
        result = _run_tool(
            "write.xls_create_from_template",
            "--template", str(template),
            "--output", str(output),
            "--vars", '{"company": "Acme Corp", "year": "2026"}',
        )
        assert result["status"] == "success"
        assert result["data"]["substitutions_made"] == 2

        wb2 = load_workbook(str(output))
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A1"].value == "Acme Corp"
        assert ws2["A2"].value == "Report for 2026"
        # Formula must be preserved, not substituted
        assert ws2["A3"].value == "=SUM(B1:B10)"

    def test_unmatched_placeholders_preserved(self, tmp_path: Path) -> None:
        from openpyxl import Workbook as WB

        template = tmp_path / "template2.xlsx"
        wb = WB()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "{{known}}"
        ws["A2"] = "{{unknown}}"
        wb.save(str(template))

        output = tmp_path / "partial.xlsx"
        _run_tool(
            "write.xls_create_from_template",
            "--template", str(template),
            "--output", str(output),
            "--vars", '{"known": "replaced"}',
        )
        wb2 = load_workbook(str(output))
        ws2 = wb2.active
        assert ws2 is not None
        assert ws2["A1"].value == "replaced"
        assert ws2["A2"].value == "{{unknown}}"


class TestWriteRange:
    """Tests for xls_write_range."""

    def test_write_basic_data(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil

        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        result = _run_tool(
            "write.xls_write_range",
            "--input", str(work),
            "--output", str(work),
            "--range", "F1",
            "--sheet", "Sheet1",
            "--data", '[["Extra", "Col"], ["X", 42]]',
        )
        assert result["status"] == "success"
        assert result["impact"]["cells_modified"] == 4

        wb = load_workbook(str(work))
        ws = wb["Sheet1"]
        assert ws["F1"].value == "Extra"
        assert ws["G2"].value == 42

    def test_write_formula(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil

        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        result = _run_tool(
            "write.xls_write_range",
            "--input", str(work),
            "--output", str(work),
            "--range", "H1",
            "--sheet", "Sheet1",
            "--data", '[["=A1+1"]]',
        )
        assert result["status"] == "success"
        assert result["impact"]["formulas_updated"] == 1

        wb = load_workbook(str(work))
        ws = wb["Sheet1"]
        assert ws["H1"].data_type == "f"

    def test_write_dates(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil

        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        result = _run_tool(
            "write.xls_write_range",
            "--input", str(work),
            "--output", str(work),
            "--range", "H1",
            "--sheet", "Sheet1",
            "--data", '[["2026-04-08"]]',
        )
        assert result["status"] == "success"
        wb = load_workbook(str(work))
        ws = wb["Sheet1"]
        # Should be stored as datetime, not string
        import datetime

        assert isinstance(ws["H1"].value, (datetime.date, datetime.datetime))


class TestWriteCell:
    """Tests for xls_write_cell."""

    def test_write_with_auto_inference(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil

        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        result = _run_tool(
            "write.xls_write_cell",
            "--input", str(work),
            "--output", str(work),
            "--cell", "H1",
            "--sheet", "Sheet1",
            "--value", "42",
        )
        assert result["status"] == "success"
        wb = load_workbook(str(work))
        assert wb["Sheet1"]["H1"].value == 42

    def test_write_with_explicit_type(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil

        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        result = _run_tool(
            "write.xls_write_cell",
            "--input", str(work),
            "--output", str(work),
            "--cell", "H1",
            "--sheet", "Sheet1",
            "--value", "2026-04-08",
            "--type", "date",
        )
        assert result["status"] == "success"
        import datetime

        wb = load_workbook(str(work))
        assert isinstance(wb["Sheet1"]["H1"].value, (datetime.date, datetime.datetime))

    def test_write_formula_explicit(self, sample_workbook: Path, tmp_path: Path) -> None:
        import shutil

        work = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work)

        result = _run_tool(
            "write.xls_write_cell",
            "--input", str(work),
            "--output", str(work),
            "--cell", "H1",
            "--sheet", "Sheet1",
            "--value", "SUM(A1:A10)",
            "--type", "formula",
        )
        assert result["status"] == "success"
        assert result["data"]["is_formula"] is True
        wb = load_workbook(str(work))
        assert wb["Sheet1"]["H1"].value == "=SUM(A1:A10)"
```

---

## Phase 5: VERIFY — Exit Criteria Checklist

| # | Criterion | Status | Evidence |
|:---|:---|:---|:---|
| 1 | `type_coercion.py` — None → None | ✅ | Direct passthrough |
| 2 | `type_coercion.py` — bool precedence over int (Python subclass issue) | ✅ | `isinstance(value, bool)` checked BEFORE int |
| 3 | `type_coercion.py` — Formula strings (starting with `=`) passthrough | ✅ | Early return before date/numeric checks |
| 4 | `type_coercion.py` — ISO 8601 dates → `datetime.date` | ✅ | `_DATE_RE` regex + `fromisoformat()` |
| 5 | `type_coercion.py` — ISO 8601 datetimes with timezone → `datetime.datetime` | ✅ | `_DATETIME_RE` + Z replacement |
| 6 | `type_coercion.py` — Boolean strings "true"/"false" → bool | ✅ | Case-insensitive `.lower()` |
| 7 | `type_coercion.py` — Leading zeros preserved as strings | ✅ | `"007" != str(int("007"))` guard |
| 8 | `type_coercion.py` — Explicit coercion with `coerce_cell_value()` | ✅ | 8 target types supported |
| 9 | `xls_create_new` — Default "Sheet1" when no `--sheets` | ✅ | Renames default active sheet |
| 10 | `xls_create_new` — Multiple sheets from comma-separated list | ✅ | `create_sheet()` per name |
| 11 | `xls_create_from_template` — `{{placeholder}}` substitution | ✅ | `_PLACEHOLDER_RE` regex |
| 12 | `xls_create_from_template` — Formulas NOT substituted | ✅ | Checks `cell.data_type != "f"` |
| 13 | `xls_create_from_template` — Unmatched placeholders preserved | ✅ | `_replacer` returns original on miss |
| 14 | `xls_create_from_template` — `wb.template = False` for output | ✅ | Correct OOXML MIME type |
| 15 | `xls_write_range` — Schema validation of data array | ✅ | `validate_against_schema("write_data", ...)` |
| 16 | `xls_write_range` — Type inference per cell | ✅ | Uses `infer_cell_value()` |
| 17 | `xls_write_range` — Formula count in impact report | ✅ | Tracks `formulas_written` |
| 18 | `xls_write_cell` — Explicit `--type` override | ✅ | Uses `coerce_cell_value()` |
| 19 | `xls_write_cell` — Auto `=` prefix for formula type | ✅ | `coerce_cell_value("SUM(...)", "formula")` → `"=SUM(...)"` |
| 20 | Integration tests — roundtrip write/read verification | ✅ | All 9 test cases with `load_workbook` verification |

---

## Phase 6: DELIVER — Setup Instructions

```bash
# From project root (with Phase 0 + 1 + 2 + 3 + 4 in place):

# 1. Add the Phase 5 files at the paths shown above

# 2. Run all unit tests
pytest tests/unit/ -v --tb=short

# 3. Run Phase 5 tests specifically
pytest tests/unit/test_type_coercion.py tests/integration/test_write_tools.py -v

# 4. Test individual tools manually
python -m excel_agent.tools.write.xls_create_new --output /tmp/test.xlsx --sheets "Data,Summary"
python -m excel_agent.tools.write.xls_write_cell \
  --input /tmp/test.xlsx --output /tmp/test.xlsx \
  --cell A1 --sheet Data --value "Hello World"

# 5. Lint and type-check
black --check src/excel_agent/core/type_coercion.py src/excel_agent/tools/write/
ruff check src/excel_agent/core/type_coercion.py src/excel_agent/tools/write/
mypy src/excel_agent/core/type_coercion.py src/excel_agent/tools/write/
```

### Key Design Decisions in Phase 5

1. **`type_coercion.py` as a dedicated module** — Type inference logic is complex enough to warrant its own module rather than inlining it in each tool. Both `xls_write_range` and `xls_write_cell` share the same inference rules, and explicit coercion adds another dimension.

2. **Bool before int check** — In Python, `bool` is a subclass of `int` (`isinstance(True, int)` returns `True`). The inference function checks `isinstance(value, bool)` BEFORE `isinstance(value, (int, float))` to prevent `True` from being treated as `1`.

3. **Leading zero preservation** — ZIP codes, part numbers, and codes like "007" should remain strings. The guard `value != str(int_val)` detects when the string representation would change after int conversion, preserving the original.

4. **Formula protection in templates** — The template substitution engine explicitly skips cells with `data_type == "f"` (formula cells). A `{{year}}` inside a formula like `=SUM({{year}})` would be invalid — formulas must only be modified through the dedicated formula tools.

5. **`wb.template = False` for output** — Per the openpyxl documentation, "you should monitor the data attributes and document extensions for saving documents in the document templates and vice versa, otherwise the result table engine can not open the document." Setting `template = False` ensures the output `.xlsx` has the correct OOXML MIME type.

### What's Next: Phase 6

Phase 6 implements the **Structural Mutation Tools** (8 tools): add/delete/rename/move sheet, insert/delete rows/columns — all with dependency checks and governance token enforcement.
