from abc import ABC, abstractmethod
from typing import Any

from app.infrastructure.extraction.imcp.common import norm_text, to_decimal


class CellConverter(ABC):
    @abstractmethod
    def convert(self, value: Any) -> Any:
        pass


class DecimalCellConverter(CellConverter):
    def convert(self, value: Any) -> Any:
        return to_decimal(value)


class StringCellConverter(CellConverter):
    def convert(self, value: Any) -> Any:
        if isinstance(value, str):
            stripped = value.strip()
            if stripped:
                return stripped
        return None


class NormStringCellConverter(CellConverter):
    def convert(self, value: Any) -> Any:
        if isinstance(value, str):
            normalized = norm_text(value)
            if normalized:
                return normalized
        return None


class CompositeCellConverter(CellConverter):
    def __init__(self, *converters: CellConverter) -> None:
        self._converters = converters

    def convert(self, value: Any) -> Any:
        current_value = value
        for converter in self._converters:
            current_value = converter.convert(current_value)
            if current_value is None:
                return None
        return current_value


class IgnoreCellConverter(CellConverter):
    def convert(self, value: Any) -> Any:
        _ = value
        return None
