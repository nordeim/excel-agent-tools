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
