from openpyxl.worksheet.worksheet import Worksheet
from app.infrastructure.extraction.imcp.base_sheet import BaseSheet
from app.infrastructure.extraction.imcp.base_abstractions import MatchAnchor


_TITLE_ANCHORS = (
    MatchAnchor("товар"),
)
_AMOUNT_HEADER_MATHER = MatchAnchor("нмц товар")
_ITEM_HEADER_MATHER = MatchAnchor("наим товар")
class GoodsSheet(BaseSheet):

    @classmethod
    def is_me(cls, ws: Worksheet) -> bool:
        sheet = cls(ws)
        if sheet._has_title_anchors(_TITLE_ANCHORS):
            return True
        return sheet._has_table_header((_AMOUNT_HEADER_MATHER, _ITEM_HEADER_MATHER))

    def extract(self) -> GoodsSheetParseResult | None:


