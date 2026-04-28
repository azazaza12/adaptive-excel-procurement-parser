from io import BytesIO

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.domain.entities.imcp.result import IMCPResult
from app.domain.exceptions import (
    ExcelIMCPNotFoundError,
)
from app.domain.service.excel_imcp_extractor import IExcelIMCPExtractor
from app.infrastructure.extraction.imcp.base_sheet import BaseSheet


def _build_sheet(ws: Worksheet) -> BaseSheet | None:
    for sheet_cls in BaseSheet.registered():
        if sheet_cls.is_me(ws):
            return sheet_cls(ws)
    return None


class OpenpyxlExcelIMCPExtractor(IExcelIMCPExtractor):
    def extract_imcp(self, *, content: bytes) -> IMCPResult:
        wb = load_workbook(BytesIO(content), data_only=True, read_only=False)

        try:
            result = IMCPResult()
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                sheet = _build_sheet(ws)
                if sheet is not None:
                    sheet.extract()
            return result
        finally:
            wb.close()
