from dataclasses import dataclass
from decimal import Decimal

from .entities import VatInfo
from .entities.goods_result import GoodsData
from .entities.works_result import WorksOrgAmountResult


@dataclass(slots=True)
class WorksSheetParseResult:
    imcp_avg: Decimal | None = None
    imcp_limit: Decimal | None = None
    orgs_amounts: list[WorksOrgAmountResult] | None = None
    amount_header_vat_info: VatInfo | None = None
    imcp_avg_row_vat_info: VatInfo | None = None


@dataclass(slots=True)
class GoodsSheetParseResult:
    goods_data: GoodsData | None = None
    imcp_avg: Decimal | None = None
    vat_rate_info: VatInfo | None = None
