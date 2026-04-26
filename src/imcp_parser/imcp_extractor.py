from io import BytesIO

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.domain.entities.excel.result import ExcelIMCPResult
from app.domain.exceptions import (
    ExcelIMCPNotFoundError,
)
from app.domain.service.excel_imcp_extractor import IExcelIMCPExtractor
from app.infrastructure.extraction.excel_imcp.common import (
    match_anchor,
    norm_text,
    select_sheet,
)
from app.infrastructure.extraction.excel_imcp.extractor_registry import (
    SheetExtractorRegistry,
    SheetType,
)
from app.infrastructure.extraction.excel_imcp.goods_sheet_extractor import (
    extract_goods_sheet,
)
from app.infrastructure.extraction.excel_imcp.works_services_sheet_extractor import (
    extract_works_sheet,
)


_IMCP_ANCHOR_AVG = "начальн максимальн цен договор средн арифметическ"
_IMCP_SHEET_GOODS = "Расчет при закупке товаров"
_IMCP_SHEET_WORKS_SERVICES = "Расчет при закупке работ, услуг"
_GOODS_MARKERS = ("товар",)
_WORKS_SERVICES_MARKERS = ("работ", "услуг")


class OpenpyxlExcelIMCPExtractor(IExcelIMCPExtractor):
    def __init__(self) -> None:
        self._imcp_anchor_avg = _IMCP_ANCHOR_AVG
        self._imcp_sheet_goods = _IMCP_SHEET_GOODS
        self._imcp_sheet_works_services = _IMCP_SHEET_WORKS_SERVICES
        self._extractor_factory = SheetExtractorRegistry()
        self._extractor_factory.register_extractor(
            SheetType.GOODS,
            extract_goods_sheet,
        )
        self._extractor_factory.register_extractor(
            SheetType.WORKS_SERVICES,
            extract_works_sheet,
        )

    def extract_imcp(self, *, content: bytes) -> ExcelIMCPResult:
        wb = load_workbook(BytesIO(content), data_only=True, read_only=False)

        try:
            result = ExcelIMCPResult()

            ws_goods = select_sheet(
                wb,
                self._imcp_sheet_goods,
                self._imcp_anchor_avg,
            )
            if ws_goods is not None and not self._is_goods_sheet_candidate(ws_goods):
                ws_goods = None
            if ws_goods is not None:
                goods_parse_result = self._extractor_factory.extract(
                    SheetType.GOODS,
                    ws_goods,
                )
                result.goods_imcp_avg = goods_parse_result.imcp_avg
                result.goods_orgs_amounts = goods_parse_result.orgs_amounts

            ws_ws = select_sheet(
                wb,
                self._imcp_sheet_works_services,
                self._imcp_anchor_avg,
                excluded_sheet_names={ws_goods.title} if ws_goods is not None else None,
            )
            if ws_ws is not None and not self._is_works_sheet_candidate(ws_ws):
                ws_ws = None
            if ws_ws is not None:
                works_parse_result = self._extractor_factory.extract(
                    SheetType.WORKS_SERVICES,
                    ws_ws,
                )
                result.works_imcp_avg = works_parse_result.imcp_avg
                result.works_imcp_limit = works_parse_result.imcp_limit
                result.works_orgs_amounts = works_parse_result.orgs_amounts

            if not result:
                msg = "No initial max contract price values found in xlsx"
                raise ExcelIMCPNotFoundError(msg)
            return result
        finally:
            wb.close()

    def _is_goods_sheet_candidate(self, ws: Worksheet) -> bool:
        by_title = self._classify_sheet_title(ws.title)
        if by_title is not None:
            return by_title == SheetType.GOODS
        return self._classify_sheet_cells(ws) == SheetType.GOODS

    def _is_works_sheet_candidate(self, ws: Worksheet) -> bool:
        by_title = self._classify_sheet_title(ws.title)
        if by_title is not None:
            return by_title == SheetType.WORKS_SERVICES
        return self._classify_sheet_cells(ws) == SheetType.WORKS_SERVICES

    def _classify_sheet_title(self, sheet_title: str) -> SheetType | None:
        title_norm = norm_text(sheet_title)
        has_goods_marker = self._contains_anchor(title_norm, _GOODS_MARKERS)
        has_works_marker = self._contains_anchor(title_norm, _WORKS_SERVICES_MARKERS)
        if has_works_marker:
            return SheetType.WORKS_SERVICES
        if has_goods_marker:
            return SheetType.GOODS
        return None

    def _classify_sheet_cells(self, ws: Worksheet) -> SheetType:
        has_goods_marker = False
        has_works_marker = False
        for row in ws.iter_rows(values_only=True):
            for value in row:
                if not isinstance(value, str):
                    continue
                normalized_value = norm_text(value)
                if self._contains_anchor(normalized_value, _WORKS_SERVICES_MARKERS):
                    has_works_marker = True
                    break
                if self._contains_anchor(normalized_value, _GOODS_MARKERS):
                    has_goods_marker = True
            if has_works_marker:
                break

        if has_works_marker:
            return SheetType.WORKS_SERVICES
        if has_goods_marker:
            return SheetType.GOODS
        return SheetType.GOODS

    def _contains_anchor(self, value: str, anchors: tuple[str, ...]) -> bool:
        return any(match_anchor(value, anchor) for anchor in anchors)
