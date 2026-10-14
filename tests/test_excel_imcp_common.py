from decimal import Decimal

from imcp_parser import parse_vat_info_from_cell_text, to_decimal
from imcp_parser.entities import VatInfo


VAT_RATE_PERCENT = 20


def test_to_decimal_ignores_excel_error_values() -> None:
    assert to_decimal("#ДЕЛ/0!") is None
    assert to_decimal("#DIV/0!") is None
    assert to_decimal("#Н/Д") is None


def test_to_decimal_parses_regular_numeric_values() -> None:
    assert to_decimal("123,45") == Decimal("123.45")
    assert to_decimal(" 1 234,56 руб. ") == Decimal("1234.56")


def test_to_decimal_limits_precision_to_two_decimal_places() -> None:
    assert to_decimal("13572602.073333332") == Decimal("13572602.07")


def test_parse_vat_rate_percent_from_cell_text_returns_rate_after_nds() -> None:
    assert parse_vat_info_from_cell_text("в т.ч. НДС 20 %") == VatInfo(
        rate=20, included=True
    )


def test_parse_vat_rate_percent_from_cell_text_returns_rate_before_nds() -> None:
    assert parse_vat_info_from_cell_text("ставка 0% НДС") == VatInfo(
        rate=0, included=None
    )


def test_parse_vat_rate_percent_from_cell_text_returns_zero_for_no_vat_text() -> None:
    assert parse_vat_info_from_cell_text("цена без НДС") == VatInfo(
        rate=None, included=False
    )
    assert parse_vat_info_from_cell_text("НДС не начисляется") == VatInfo(
        rate=0, included=None
    )


def test_parse_vat_rate_percent_from_cell_text_returns_none_for_unknown_text() -> None:
    assert parse_vat_info_from_cell_text("составляет, руб.") is None
    assert parse_vat_info_from_cell_text("   ") is None
