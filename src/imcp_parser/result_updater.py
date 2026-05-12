from functools import singledispatch

from app.domain.entities.imcp.result import IMCPResult
from app.domain.exceptions import UnsupportedSheetResultError
from app.infrastructure.extraction.imcp.models import (
    GoodsSheetParseResult,
    WorksSheetParseResult,
)


@singledispatch
def update_result(parsed: object, _result: IMCPResult) -> None:
    msg = (
        f"Unsupported parsed result type: {type(parsed).__module__}."
        f"{type(parsed).__qualname__}"
    )
    raise UnsupportedSheetResultError(msg)


@update_result.register(WorksSheetParseResult)
def _(parsed: WorksSheetParseResult, result: IMCPResult) -> None:
    result.works_imcp_avg = parsed.imcp_avg
    result.works_imcp_limit = parsed.imcp_limit
    result.works_orgs_amounts = parsed.orgs_amounts
    result.works_amount_header_vat_info = parsed.amount_header_vat_info
    result.works_imcp_avg_row_vat_info = parsed.imcp_avg_row_vat_info


@update_result.register(GoodsSheetParseResult)
def _(parsed: GoodsSheetParseResult, result: IMCPResult) -> None:
    result.goods_data = parsed.goods_data
    result.goods_imcp_avg = parsed.imcp_avg
    result.goods_vat_info = parsed.vat_rate_info
