from collections.abc import Callable
from enum import StrEnum
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet


class SheetType(StrEnum):
    GOODS = "goods"
    WORKS_SERVICES = "works_services"


class SheetExtractorRegistry:
    def __init__(self) -> None:
        self._extractors: dict[SheetType, Callable[[Worksheet], Any | None]] = {}

    def register_extractor(
        self,
        sheet_type: SheetType,
        extractor: Callable[[Worksheet], Any | None],
    ) -> None:
        self._extractors[sheet_type] = extractor

    def extract(self, sheet_type: SheetType, ws: Worksheet) -> Any | None:
        extractor = self._extractors.get(sheet_type)
        if extractor is None:
            msg = f"Extractor is not registered for sheet type: {sheet_type}"
            raise ValueError(msg)
        return extractor(ws)
