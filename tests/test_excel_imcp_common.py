from decimal import Decimal

from app.infrastructure.extraction.imcp.common import (
    to_decimal,
)


def test_to_decimal_ignores_excel_error_values() -> None:
    assert to_decimal("#ДЕЛ/0!") is None
    assert to_decimal("#DIV/0!") is None
    assert to_decimal("#Н/Д") is None


def test_to_decimal_parses_regular_numeric_values() -> None:
    assert to_decimal("123,45") == Decimal("123.45")
    assert to_decimal(" 1 234,56 руб. ") == Decimal("1234.56")


def test_to_decimal_limits_precision_to_two_decimal_places() -> None:
    assert to_decimal("13572602.073333332") == Decimal("13572602.07")
