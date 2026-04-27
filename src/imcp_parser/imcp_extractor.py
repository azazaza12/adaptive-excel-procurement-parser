from io import BytesIO

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from app.domain.entities.excel.result import ExcelIMCPResult
from app.domain.exceptions import (
    ExcelIMCPNotFoundError,
)
from app.domain.service.excel_imcp_extractor import IExcelIMCPExtractor
from app.infrastructure.extraction.excel_imcp.works_sheet import WorksSheet


def _find_works_sheet(wb: Workbook) -> WorksSheet | None:
    for sheet_name in wb.sheetnames:
        worksheet = wb[sheet_name]
        if WorksSheet.is_me(worksheet):
            return WorksSheet(worksheet)
    return None


class OpenpyxlExcelIMCPExtractor(IExcelIMCPExtractor):
    def extract_imcp(self, *, content: bytes) -> ExcelIMCPResult:
        wb = load_workbook(BytesIO(content), data_only=True, read_only=False)

        try:
            result = ExcelIMCPResult()
            works_sheet = _find_works_sheet(wb)
            if works_sheet is not None:
                works_parse_result = works_sheet.extract()
                result.works_imcp_avg = works_parse_result.imcp_avg
                result.works_imcp_limit = works_parse_result.imcp_limit
                result.works_orgs_amounts = works_parse_result.orgs_amounts

            if not result:
                msg = "No initial max contract price values found in xlsx"
                raise ExcelIMCPNotFoundError(msg)
            return result
        finally:
            wb.close()
