from openpyxl.worksheet.worksheet import Worksheet

from app.infrastructure.extraction.imcp.models import WorksSheetParseResult
from app.infrastructure.extraction.imcp.works_sheet import WorksSheet


def extract_works_sheet(
    ws: Worksheet,
) -> WorksSheetParseResult:
    return WorksSheet(ws).extract()
