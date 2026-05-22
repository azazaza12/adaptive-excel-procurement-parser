import re
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any

from app.domain.entities.imcp.result import VatInfo


_NUM_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
_CELL_TEXT_VAT_RATE_RE = re.compile(r"(?i)НДС\D*(\d+)\s*%?")
_CELL_TEXT_VAT_RATE_BEFORE_NDS_RE = re.compile(r"(?i)(\d+)\s*%\s*НДС")
_NO_VAT_PATTERNS = (
    re.compile(r"(?i)\bндс\s+не\s+предусмотрен\b"),
    re.compile(r"(?i)\bндс\s+не\s+облагается\b"),
    re.compile(r"(?i)\bне\s+облагается\s+ндс\b"),
    re.compile(r"(?i)\bндс\s+не\s+начисляется\b"),
)
_INCLUDE_VAT_PATTERNS = (
    re.compile(r"(?i)\bс\s+ндс\b"),
    re.compile(r"(?i)\bв\s+том\s+числе\s+ндс\b"),
    re.compile(r"(?i)\bв\s*т\.?\s*ч\.?\s*ндс\b"),
    re.compile(r"(?i)\bндс\s+включ[её]н\b"),
    re.compile(r"(?i)\bс\s+уч[её]том\s+ндс\b"),
    re.compile(r"(?i)\bвключая\s+ндс\b"),
)
_NOT_INCLUDE_VAT_PATTERNS = (
    re.compile(r"(?i)\bбез\s+уч[её]та\s+ндс\b"),
    re.compile(r"(?i)\bндс\s+не\s+включ[её]н\b"),
    re.compile(r"(?i)\bбез\s+ндс\b"),
)
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
    exp = result.as_tuple().exponent
    if isinstance(exp, int) and exp < _TWO_DECIMAL_PLACES_EXPONENT:
        return result.quantize(_TWO_DECIMAL_PLACES, rounding=ROUND_HALF_UP)
    return result


def to_stripped_str(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    stripped = value.strip()
    return stripped or None


def parse_vat_info_from_cell_text(value: Any) -> VatInfo | None:
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None

    vat_rate = _extract_vat_rate(text)
    if vat_rate is not None:
        included_vat = _extract_vat_included(text)
        return VatInfo(rate=vat_rate, included=included_vat)
    vat_rate, included_vat = _extract_no_vat(text)
    if vat_rate is None and included_vat is None:
        return None
    return VatInfo(rate=vat_rate, included=included_vat)


def _extract_vat_rate(text: str) -> int | None:
    match = _CELL_TEXT_VAT_RATE_RE.search(text)
    if match is not None:
        return int(match.group(1))

    match = _CELL_TEXT_VAT_RATE_BEFORE_NDS_RE.search(text)
    if match is not None:
        return int(match.group(1))
    return None


def _extract_vat_included(text: str) -> bool | None:
    for pattern in _NOT_INCLUDE_VAT_PATTERNS:
        match = pattern.search(text)
        if match is not None:
            return False
    for pattern in _INCLUDE_VAT_PATTERNS:
        match = pattern.search(text)
        if match is not None:
            return True
    return None


def _extract_no_vat(text: str) -> tuple[int | None, bool | None]:
    for pattern in _NO_VAT_PATTERNS:
        match = pattern.search(text)
        if match is not None:
            return 0, None
    included_vat = _extract_vat_included(text)
    return None, included_vat
