"""Type inference and explicit coercion for cell values.

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
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_DATETIME_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$"
)


def infer_cell_value(value: Any) -> Any:
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


def coerce_cell_value(value: str, target_type: str) -> Any:
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
