from dataclasses import dataclass

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
class _WorksRow:
    name: str | None
    amount: object


class WorksSheet(BaseSheet[WorksSheetParseResult]):
    @classmethod
    def is_me(cls, ws: Worksheet) -> bool:
        if cls._has_works_title(ws):
            return True
        return cls._has_works_table_header(ws)

    def extract(self) -> WorksSheetParseResult:
        header = find_header_columns(
            self.ws,
            {
                "item_col_idx": _ORG_HEADER_ANCHOR,
                "amount": _AMOUNT_HEADER_ANCHOR,
            },
        )

        orgs_amounts: WorksOrgsAmountsResult | None = None
        if header is not None:
            header_row_idx, _ = header
            org_column = BaseColumn(
                self.ws,
                MatchAnchor(_ORG_HEADER_ANCHOR),
                converter=StringConverter(),
                header_row_idx=header_row_idx,
            )
            amount_column = BaseColumn(
                self.ws,
                MatchAnchor(_AMOUNT_HEADER_ANCHOR),
                converter=DecimalConverter(),
                header_row_idx=header_row_idx,
            )
            works_list = BaseList(
                org_column,
                amount_column,
                row_builder=lambda name, amount: _WorksRow(name=name, amount=amount),
            )
            organizations: list[WorksOrgAmountResult] = []
            empty_rows_in_series = 0
            for row_idx in range(header_row_idx + 1, self.ws.max_row + 1):
                row = works_list.build_row(row_idx)
                if row is None or row.name is None:
                    empty_rows_in_series, should_stop = advance_empty_rows_series(
                        empty_rows_in_series,
                        is_empty=True,
                        max_empty_rows_in_series=_MAX_EMPTY_ROWS_IN_SERIES,
                    )
                    if should_stop:
                        break
                    continue

                empty_rows_in_series, _ = advance_empty_rows_series(
                    empty_rows_in_series,
                    is_empty=False,
                    max_empty_rows_in_series=_MAX_EMPTY_ROWS_IN_SERIES,
                )

                if is_stop_row_by_anchor_match(row.name, _STOP_ROW_ANCHORS):
                    break
                if row.amount is None:
                    continue

                organizations.append(
                    WorksOrgAmountResult(
                        id=len(organizations) + 1,
                        name=row.name,
                        amount=row.amount,
                    )
                )

            if organizations:
                orgs_amounts = WorksOrgsAmountsResult(organizations=organizations)

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
