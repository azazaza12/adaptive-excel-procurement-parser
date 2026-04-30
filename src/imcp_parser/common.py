import re
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet


_NUM_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
_MIN_ANCHOR_WORD_LEN = 2
_TWO_DECIMAL_PLACES = Decimal("0.01")
_TWO_DECIMAL_PLACES_EXPONENT = -2


def norm_text(value: str) -> str:
    normalized = value.casefold().replace("ё", "е")
    normalized = normalized.replace("\xa0", " ")
    normalized = re.sub(r"[^\w]+", " ", normalized, flags=re.UNICODE)
    return " ".join(normalized.split())


def parse_number(value: object) -> int | None:
    if isinstance(value, int):
        return value if value > 0 else None
    if isinstance(value, float):
        int_value = int(value)
        if value == int_value and int_value > 0:
            return int_value
        return None
    if isinstance(value, str):
        stripped = value.strip().replace(".", "")
        if stripped.isdigit():
            parsed = int(stripped)
            return parsed if parsed > 0 else None
    return None


def match_anchor(cell_value: str, anchor: str) -> bool:
    if not isinstance(cell_value, str):
        return False

    cell_norm = norm_text(cell_value)
    anchor_words = [
        word for word in norm_text(anchor).split() if len(word) >= _MIN_ANCHOR_WORD_LEN
    ]
    if not anchor_words:
        return False
    cell_words = cell_norm.split()

    for anchor_word in anchor_words:
        if not any(word.startswith(anchor_word) for word in cell_words):
            return False

    return True


def to_decimal(value: Any) -> Decimal | None:
    result: Decimal | None = None

    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, int | float):
        result = Decimal(str(value))
    elif isinstance(value, str):
        stripped = value.strip()
        if stripped:
            if stripped.startswith("#"):
                return None
            stripped = stripped.replace("\xa0", " ").replace(" ", "")
            match = _NUM_RE.search(stripped)
            if match:
                number = match.group(0).replace(",", ".")
                try:
                    result = Decimal(number)
                except InvalidOperation:
                    result = None

    if result is None:
        return None
    if result.as_tuple().exponent < _TWO_DECIMAL_PLACES_EXPONENT:
        return result.quantize(_TWO_DECIMAL_PLACES, rounding=ROUND_HALF_UP)
    return result


def to_stripped_str(value: Any):
    if not isinstance(value, str):
        return None
    stripped = value.strip()
    return stripped or None
