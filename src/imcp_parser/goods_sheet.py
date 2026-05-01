from dataclasses import dataclass
from openpyxl.worksheet.worksheet import Worksheet

from app.infrastructure.extraction.imcp.models import GoodsSheetParseResult
from app.infrastructure.extraction.imcp.base_sheet import BaseSheet
from app.infrastructure.extraction.imcp.base_abstractions import MatchAnchor, AnyOfAnchors


_INDEX_HEADER_MATCHER = AnyOfAnchors(
    MatchAnchor("№ п п"),
    MatchAnchor("номер"),
)
_TITLE_ANCHORS = (
    MatchAnchor("товар"),
)
_ITEM_NAME_HEADER_MATHER = MatchAnchor("наименова товар")
_GOODS_PRICE_BLOCK_PARENT_ANCHORS = MatchAnchor("цена единицы товара",)
_AMOUNT_HEADER_MATHER = MatchAnchor("нмц товар")
_UNIT_IMCP_ANCHORS = MatchAnchor("нмц единиц",)
_UNIT_OF_MEASURE_ANCHORS = MatchAnchor("ед изм",)
_QUANTITY_ANCHORS = AnyOfAnchors(MatchAnchor("кол"), MatchAnchor("всего"))
_ITEM_IMCP_ANCHORS = AnyOfAnchors(MatchAnchor("нмц товар"), MatchAnchor("нмц"))
_IMCP_ANCHOR_AVG = MatchAnchor("начальн максимальн цен договор средн арифметическ")


@dataclass(slots=True)
class GoodsHeader:
    start_row_idx: int
    end_row_idx: int
    index_col_idx: int
    item_name_col_idx: int
    price_block_start_col_idx: int
    unit_imcp_col_idx: int
    unit_of_measure_col_idx: int
    quantity_col_idx: int
    item_imcp_col_idx: int
    item_imcp_avg_col_idx: int


@dataclass(slots=True)
class SourceHeader:
    col_idx: int
    name: str


#TODO поиск границ блока цен и извлечение заголовков
class GoodsSheet(BaseSheet):

    @classmethod
    def is_me(cls, ws: Worksheet) -> bool:
        sheet = cls(ws)
        if sheet._has_title_anchors(_TITLE_ANCHORS):
            return True
        return sheet._has_table_header((_AMOUNT_HEADER_MATHER, _ITEM_NAME_HEADER_MATHER))

    def extract(self) -> GoodsSheetParseResult | None:
        header = None
        index_header = self._find_header_by(_INDEX_HEADER_MATCHER)
        item_name_header = self._find_header_by(_ITEM_NAME_HEADER_MATHER)
        price_block_header = self._find_header_by(_GOODS_PRICE_BLOCK_PARENT_ANCHORS)
        unit_imcp_header = self._find_header_by(_UNIT_IMCP_ANCHORS)
        unit_of_measure_anchor = self._find_header_by(_UNIT_OF_MEASURE_ANCHORS)
        quantity_header = self._find_header_by(_QUANTITY_ANCHORS)
        item_imcp_header = self._find_header_by(_ITEM_IMCP_ANCHORS)
        item_imcp_avg_header = self._find_header_by(_IMCP_ANCHOR_AVG)

        if index_header and item_name_header and price_block_header and unit_imcp_header and quantity_header and item_imcp_header and item_imcp_avg_header:
            price_block_end_col_idx = unit_imcp_header.col_idx - 1
            source_headers = self._extract_source_headers_in_block(
                header_row_idx=price_block_header.row_idx,
                start_col_idx=price_block_header.col_idx,
                end_col_idx=price_block_end_col_idx,
            )

        return None

    def _extract_source_headers_in_block(
        self,
        *,
        header_row_idx: int,
        start_col_idx: int,
        end_col_idx: int,
    ) -> list[SourceHeader]:
        source_headers: list[SourceHeader] = []

        for col_idx in range(start_col_idx, end_col_idx + 1):
            raw_value = self.ws.cell(row=header_row_idx + 1, column=col_idx).value
            name = raw_value.strip() if isinstance(raw_value, str) else ""
            if not name:
                continue
            source_headers.append(SourceHeader(col_idx=col_idx, name=name))

        return source_headers









