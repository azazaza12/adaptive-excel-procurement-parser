from dataclasses import dataclass
from decimal import Decimal

from openpyxl.worksheet.worksheet import Worksheet

from app.domain.entities.imcp.result import IMCPResult
from app.domain.entities.imcp.works_result import WorksOrgAmountResult
from app.infrastructure.extraction.imcp.base_abstractions import (
    AnyOfAnchors,
    BaseAnchor,
    MatchAnchor,
)
from app.infrastructure.extraction.imcp.base_sheet import BaseSheet
from app.infrastructure.extraction.imcp.common import (
    advance_empty_rows_series,
    find_anchor_row_idx,
    norm_text,
)
from app.infrastructure.extraction.imcp.models import WorksSheetParseResult
from app.infrastructure.extraction.imcp.common import to_decimal, to_stripped_str
from enum import Enum


_WORKS_SHEET_TITLE_ANCHORS = ("работ", "услуг")
_MAX_EMPTY_ROWS_IN_SERIES = 3

_ORG_HEADER_MATCHER = MatchAnchor("организац")
_AMOUNT_HEADER_WORKS_OR_SERVICES_MATCHER = AnyOfAnchors(
    MatchAnchor("стоимост работ"),
    MatchAnchor("стоимост услуг"),
)
_IMCP_ROW_MATCHER = MatchAnchor("начальн максимальн цен договор")
_IMCP_AVG_ROW_MATCHER = MatchAnchor("начальн максимальн цен договор средн арифметическ")
_IMCP_LIMIT_ROW_MATCHER = MatchAnchor("лимит финансирован")
_INSTRUCTION_ROW_MATCHER = MatchAnchor("инструкц")
_STOP_ROW_MATCHER = AnyOfAnchors(_IMCP_ROW_MATCHER, _INSTRUCTION_ROW_MATCHER)


@dataclass(slots=True)
class WorksHeader:
    row_idx: int
    org_col_idx: int
    amount_col_idx: int


def _find_works_header(ws: Worksheet) -> WorksHeader | None:
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        org_col_idx: int | None = None
        amount_col_idx: int | None = None
        for cell in row:
            value = cell.value

            if org_col_idx is None and _ORG_HEADER_MATCHER.matches(value):
                org_col_idx = cell.column
            if amount_col_idx is None and _AMOUNT_HEADER_WORKS_OR_SERVICES_MATCHER.matches(value):
                amount_col_idx = cell.column
        if org_col_idx is not None and amount_col_idx is not None:
            return WorksHeader(
                row_idx=row_idx,
                org_col_idx=org_col_idx,
                amount_col_idx=amount_col_idx,
            )
    return None


def _row_has_anchor(ws: Worksheet, *, row_idx: int, anchor: BaseAnchor) -> bool:
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(row=row_idx, column=col_idx).value
        if isinstance(value, str) and anchor.matches(value):
            return True
    return False


def _extract_amount_from_row(
    ws: Worksheet, *, row_idx: int, amount_col_idx: int
) -> Decimal | None:
    amount_from_header_column = to_decimal(
        ws.cell(row=row_idx, column=amount_col_idx).value
    )
    if amount_from_header_column is not None:
        return amount_from_header_column

    for col_idx in range(ws.max_column, 0, -1):
        amount = to_decimal(ws.cell(row=row_idx, column=col_idx).value)
        if amount is not None:
            return amount
    return None


def _extract_org_amount_rows_until_imcp(
    ws: Worksheet, header: WorksHeader
) -> tuple[list[WorksOrgAmountResult], int | None]:
    organizations: list[WorksOrgAmountResult] = []
    empty_rows_in_series = 0
    imcp_row_idx: int | None = None

    for row_idx in range(header.row_idx + 1, ws.max_row + 1):
        if _row_has_anchor(ws, row_idx=row_idx, anchor=_IMCP_ROW_MATCHER):
            imcp_row_idx = row_idx
            break

        name = to_stripped_str(
            ws.cell(row=row_idx, column=header.org_col_idx).value
        )
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


        if _STOP_ROW_MATCHER.matches(name):
            break

        amount = to_decimal(
            ws.cell(row=row_idx, column=header.amount_col_idx).value
        )
        if amount is None:
            continue

        organizations.append(
            WorksOrgAmountResult(
                id=len(organizations) + 1,
                name=name,
                amount=amount,
            )
        )

    return organizations, imcp_row_idx


def _extract_imcp_from_imcp_block(
    ws: Worksheet,
    *,
    imcp_row_idx: int | None,
    amount_col_idx: int,
) -> tuple[Decimal | None, Decimal | None]:
    imcp_avg: Decimal | None = None
    imcp_limit: Decimal | None = None
    if imcp_row_idx is None:
        return imcp_avg, imcp_limit

    for row_idx in range(imcp_row_idx, ws.max_row + 1):
        if imcp_avg is None and _row_has_anchor(ws, row_idx=row_idx, anchor=_IMCP_AVG_ROW_MATCHER):
            imcp_avg = _extract_amount_from_row(
                ws,
                row_idx=row_idx,
                amount_col_idx=amount_col_idx,
            )

        if imcp_limit is None and _row_has_anchor(
            ws, row_idx=row_idx, anchor=_IMCP_LIMIT_ROW_MATCHER
        ):
            imcp_limit = _extract_amount_from_row(
                ws,
                row_idx=row_idx,
                amount_col_idx=amount_col_idx,
            )

        if imcp_avg is not None and imcp_limit is not None:
            break
        if _row_has_anchor(ws, row_idx=row_idx, anchor=_INSTRUCTION_ROW_MATCHER):
            break

    return imcp_avg, imcp_limit


def _extract_org_amount_rows_and_imcp(
    ws: Worksheet, header: WorksHeader
) -> tuple[list[WorksOrgAmountResult], Decimal | None, Decimal | None]:
    organizations, nmcd_row_idx = _extract_org_amount_rows_until_imcp(ws, header)
    imcp_avg, imcp_limit = _extract_imcp_from_imcp_block(
        ws,
        imcp_row_idx=nmcd_row_idx,
        amount_col_idx=header.amount_col_idx,
    )
    return organizations, imcp_avg, imcp_limit



class WorksSheet(BaseSheet[WorksSheetParseResult]):
    @classmethod
    def is_me(cls, ws: Worksheet) -> bool:
        if cls._has_works_title(ws):
            return True
        return cls._has_works_table_header(ws)

    def extract(self) -> WorksSheetParseResult:
        header = None
        org_header = self._find_header_by(_ORG_HEADER_MATCHER)
        amount_header = self._find_header_by(_AMOUNT_HEADER_WORKS_OR_SERVICES_MATCHER)
        if org_header and amount_header and org_header.row_idx == amount_header.row_idx:
            header = WorksHeader(row_idx=org_header.row_idx, org_col_idx=org_header.col_idx, amount_col_idx=amount_header.col_idx)

        org_amount_rows: list[WorksOrgAmountResult] = []
        imcp_avg: Decimal | None = None
        imcp_limit: Decimal | None = None
        if header is not None:
            org_amount_rows, imcp_avg, imcp_limit = _extract_org_amount_rows_and_imcp(
                self.ws, header
            )
        orgs_amounts = org_amount_rows or None
        return WorksSheetParseResult(
            imcp_avg=imcp_avg,
            imcp_limit=imcp_limit,
            orgs_amounts=orgs_amounts,
        )

    def _has_works_title(self) -> bool:
        title_norm = norm_text(self.ws.title)
        return all(anchor in title_norm for anchor in _WORKS_SHEET_TITLE_ANCHORS)

    def _has_works_table_header(self) -> bool:
        if self._find_header_by(_ORG_HEADER_MATCHER) is not None and self._find_header_by(_AMOUNT_HEADER_WORKS_OR_SERVICES_MATCHER) is not None:
            return True
        return False

    def apply_to_result(self, target: IMCPResult, extracted: WorksSheetParseResult) -> None:
        target.works_imcp_avg = extracted.imcp_avg
        target.works_imcp_limit = extracted.imcp_limit
        target.works_orgs_amounts = extracted.orgs_amounts
