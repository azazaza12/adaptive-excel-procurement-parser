from collections.abc import Sequence
from typing import Any

from openpyxl.workbook.workbook import Workbook

from app.infrastructure.extraction.excel_imcp.base_sheet import BaseSheet

SheetClass = type[BaseSheet[Any]]


class SheetFactory:
    def __init__(
        self,
        sheet_classes: Sequence[SheetClass],
        *,
        raise_on_duplicates: bool = False,
    ) -> None:
        self._sheet_classes = tuple(sheet_classes)
        self._raise_on_duplicates = raise_on_duplicates

    def create_all(self, workbook: Workbook) -> list[BaseSheet[Any]]:
        sheets: list[BaseSheet[Any]] = []
        for ws in workbook.worksheets:
            for sheet_class in self._sheet_classes:
                if not sheet_class.is_me(ws):
                    continue
                sheets.append(sheet_class(ws))
        return sheets

    def create_unique(self, workbook: Workbook) -> dict[SheetClass, BaseSheet[Any]]:
        matched_sheets = self.create_all(workbook)
        unique_sheets: dict[SheetClass, BaseSheet[Any]] = {}

        for sheet in matched_sheets:
            sheet_class = type(sheet)
            if sheet_class not in unique_sheets:
                unique_sheets[sheet_class] = sheet
                continue
            if self._raise_on_duplicates:
                msg = f"Multiple sheets matched {sheet_class.__name__}"
                raise ValueError(msg)

        return unique_sheets
