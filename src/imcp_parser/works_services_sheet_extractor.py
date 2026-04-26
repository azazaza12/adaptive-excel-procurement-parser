from openpyxl.worksheet.worksheet import Worksheet

from app.domain.entities.excel.result import WorksOrgsAmountsResult
from app.domain.entities.excel.works_services_result import WorksOrgAmountResult
from app.infrastructure.extraction.excel_imcp.cell_converters import (
    DecimalCellConverter,
    StringCellConverter,
)
from app.infrastructure.extraction.excel_imcp.common import (
    extract_amount_by_anchor,
    find_metric_column_idx,
)
from app.infrastructure.extraction.excel_imcp.dynamic_table_extractor import (
    DynamicRowData,
)
from app.infrastructure.extraction.excel_imcp.extractor_builder import ExtractorBuilder
from app.infrastructure.extraction.excel_imcp.models import (
    ParsingContext,
    WorksSheetParseResult,
)


_ORG_HEADER_ANCHOR = "организац"
_AMOUNT_HEADER_ANCHOR = "стоимост выполнен работ оказан услуг"
_STOP_ROW_ANCHORS = (
    "начальн максимальн цен договор",
    "инструкц",
)
_IMCP_ANCHOR_AVG = "начальн максимальн цен договор средн арифметическ"
_IMCP_ANCHOR_LIMIT = "начальн максимальн цен договор лимит финансирован"
_IMCP_ANCHOR_LIMIT_FALLBACK = "лимит финансирован"
_MAX_EMPTY_ROWS_IN_SERIES = 3
_DECIMAL_CONVERTER = DecimalCellConverter()
_STRING_CONVERTER = StringCellConverter()


def _extract_item_name(
    ws: Worksheet, row_idx: int, context: ParsingContext
) -> str | None:
    value = ws.cell(row=row_idx, column=context.header.item_col_idx).value
    return _STRING_CONVERTER.convert(value)


def _extract_metric(
    ws: Worksheet,
    row_idx: int,
    metric_name: str,
    col_idx: int,
    context: ParsingContext,
) -> object:
    _ = metric_name
    _ = context
    value = ws.cell(row=row_idx, column=col_idx).value
    return _DECIMAL_CONVERTER.convert(value)


def _is_valid_works_row(row_data: DynamicRowData) -> bool:
    return row_data.name is not None and row_data.metrics.get("amount") is not None


def _aggregate_works(
    context: ParsingContext,
    row_data_list: list[DynamicRowData],
) -> WorksOrgsAmountsResult | None:
    _ = context
    organizations: list[WorksOrgAmountResult] = []
    for row_data in row_data_list:
        name = row_data.name
        amount = row_data.metrics.get("amount")
        if name is None or amount is None:
            continue
        organizations.append(
            WorksOrgAmountResult(
                id=len(organizations) + 1,
                name=name,
                amount=amount,
            )
        )
    if not organizations:
        return None
    return WorksOrgsAmountsResult(organizations=organizations)


def extract_works_sheet(
    ws: Worksheet,
) -> WorksSheetParseResult:
    amount_col_idx = find_metric_column_idx(
        ws,
        item_anchor=_ORG_HEADER_ANCHOR,
        metric_anchor=_AMOUNT_HEADER_ANCHOR,
    )
    extractor = (
        ExtractorBuilder[WorksOrgsAmountsResult]()
        .with_header_anchors(
            {
                "item_col_idx": _ORG_HEADER_ANCHOR,
                "amount": _AMOUNT_HEADER_ANCHOR,
            }
        )
        .with_stop_row_anchors(*_STOP_ROW_ANCHORS)
        .with_max_empty_rows_in_series(_MAX_EMPTY_ROWS_IN_SERIES)
        .with_item_name_extractor(_extract_item_name)
        .with_metric_extractor(_extract_metric)
        .with_is_valid_row(_is_valid_works_row)
        .with_aggregator(_aggregate_works)
        .build()
    )
    imcp_limit = extract_amount_by_anchor(
        ws,
        _IMCP_ANCHOR_LIMIT,
        preferred_col_idx=amount_col_idx,
    )
    if imcp_limit is None:
        imcp_limit = extract_amount_by_anchor(
            ws,
            _IMCP_ANCHOR_LIMIT_FALLBACK,
            preferred_col_idx=amount_col_idx,
        )
    return WorksSheetParseResult(
        imcp_avg=extract_amount_by_anchor(
            ws,
            _IMCP_ANCHOR_AVG,
            preferred_col_idx=amount_col_idx,
        ),
        imcp_limit=imcp_limit,
        orgs_amounts=extractor.extract(ws),
    )
