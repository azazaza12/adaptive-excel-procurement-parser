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

        return None









