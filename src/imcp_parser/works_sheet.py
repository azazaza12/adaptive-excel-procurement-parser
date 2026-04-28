from dataclasses import dataclass
from decimal import Decimal

from openpyxl.worksheet.worksheet import Worksheet

from app.domain.entities.imcp.result import WorksOrgsAmountsResult, IMCPResult
from app.domain.entities.imcp.works_result import WorksOrgAmountResult
from app.infrastructure.extraction.imcp.base_abstractions import (
    BaseList,
    BaseColumn,
    DecimalConverter,
    MatchAnchor,
    StringConverter,
)
from app.infrastructure.extraction.imcp.base_sheet import BaseSheet
from app.infrastructure.extraction.imcp.common import (
    advance_empty_rows_series,
    extract_amount_by_anchor,
    find_anchor_row_idx,
    find_header_columns,
    is_stop_row_by_anchor_match,
    norm_text,
)
from app.infrastructure.extraction.imcp.models import WorksSheetParseResult
from app.infrastructure.extraction.imcp.common import norm_text, match_anchor


_WORKS_SHEET_TITLE_ANCHORS = ("работ", "услуг")
_ORG_HEADER_ANCHOR = "организац"
_AMOUNT_HEADER_ANCHOR = "стоимост выполнен работ оказан услуг"
_STOP_ROW_ANCHORS = (
    "начальн максимальн цен договор",
    "инструкц",
)
_IMCP_ANCHOR_AVG = "начальн максимальн цен договор средн арифметическ"
_IMCP_ANCHOR_LIMIT = "начальн максимальн цен договор лимит финансирован"
_MAX_EMPTY_ROWS_IN_SERIES = 3




@dataclass(slots=True)
class WorksRow:
    name: str | None
    amount: Decimal
    id: int



def _is_amount_header(value: object) -> bool:
    if not isinstance(value, str):
        return False
    return match_anchor(value, "стоимост") and (match_anchor(value, "работ") or match_anchor(value, "услуг"))


def _is_org_header(value: object) -> bool:
    if not isinstance(value, str):
        return False
    return match_anchor(value, "организац")

def _find_works_header(ws: Worksheet) -> tuple[int, int] | None:
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        org_col_idx: int | None = None
        amount_col_idx: int | None = None

        for cell in row:
            value = cell.value

            if org_col_idx is None and is_org_header(value):
                org_col_idx = cell.column
                header_row_idx = row_idx
                continue
            if amount_col_idx is None and is_amount_header(value):
                amount_col_idx = cell.column
                continue
    return header_row_idx, org_col_idx, amount_col_idx


def _extract_works_rows(ws: Worksheet, header: tuple[int, int, int]) -> list[WorksRow]
    rows: list[WorksRow] = []
    header_row_idx, org_col_idx, amount_col_idx = header
    empty_rows_in_series = 0
    for row_idx in range(header_row_idx + 1, self.ws.max_row + 1):
        name = StringConverter().convert(self.ws.cell(row_idx, org_col_idx).value)
        is_empty_name = name is None
        empty_rows_in_series, should_stop = advance_empty_rows_series(
            empty_rows_in_series,
            is_empty=is_empty_name,
            max_empty_rows_in_series=_MAX_EMPTY_ROWS_IN_SERIES,
        )
        if is_empty_name:
            if should_stop:
                break
            continue


        if is_stop_row_by_anchor_match(name, _STOP_ROW_ANCHORS):
            break

        amount = DecimalConverter().convert(self.ws.cell(row_idx, amount_col_idx).value)
        if amount is None:
            continue

        rows.append(WorksRow(name=name, amount=amount, id=row_idx))
    return rows



class WorksSheet(BaseSheet[WorksSheetParseResult]):
    @classmethod
    def is_me(cls, ws: Worksheet) -> bool:
        if cls._has_works_title(ws):
            return True
        return cls._has_works_table_header(ws)

    def extract(self) -> WorksSheetParseResult:
        header = _find_works_header(self.ws)
        rows: list[WorksRow] = []
        if header is not None:
            rows = _extract_works_rows(self.ws, header)
        #TODO row to orgs_amounts
        return WorksSheetParseResult(
            imcp_avg=extract_amount_by_anchor(self.ws, _IMCP_ANCHOR_AVG),
            imcp_limit=extract_amount_by_anchor(self.ws, _IMCP_ANCHOR_LIMIT),
            orgs_amounts=orgs_amounts,
        )

    @staticmethod
    def _has_works_title(ws: Worksheet) -> bool:
        title_norm = norm_text(ws.title)
        return all(anchor in title_norm for anchor in _WORKS_SHEET_TITLE_ANCHORS)

    @staticmethod
    def _has_works_table_header(ws: Worksheet) -> bool:
        header = find_header_columns(
            ws,
            {
                "item_col_idx": _ORG_HEADER_ANCHOR,
                "amount": _AMOUNT_HEADER_ANCHOR,
            },
        )
        if header is not None:
            return True
        return find_anchor_row_idx(ws, _IMCP_ANCHOR_LIMIT) is not None

    def apply_to_result(self, target: IMCPResult, extracted: WorksSheetParseResult) -> None:
        target.works_imcp_avg = extracted.imcp_avg
        target.works_imcp_limit = extracted.imcp_limit
        target.works_orgs_amounts = extracted.orgs_amounts
