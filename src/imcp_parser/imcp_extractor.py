from io import BytesIO

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.domain.entities.imcp.result import IMCPResult
from app.domain.exceptions import (
    IMCPNotFoundError,
)
from app.domain.service.excel_imcp_extractor import IExcelIMCPExtractor
from app.infrastructure.extraction.imcp.base_sheet import BaseSheet
from app.infrastructure.extraction.imcp.result_updater import update_result


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
            has_supported_data = False
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                sheet = _build_sheet(ws)
                if sheet is None:
                    continue
                parsed = sheet.extract()
                if parsed is not None:
                    update_result(parsed, result)
                    has_supported_data = True
            if not has_supported_data:
                raise IMCPNotFoundError
            return result
        finally:
            wb.close()
