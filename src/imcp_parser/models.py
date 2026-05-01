from dataclasses import dataclass
from decimal import Decimal

from app.domain.entities.imcp.goods_result import GoodsData
from app.domain.entities.imcp.works_result import WorksOrgAmountResult


@dataclass(slots=True)
class WorksSheetParseResult:
    imcp_avg: Decimal | None = None
    imcp_limit: Decimal | None = None
    orgs_amounts: list[WorksOrgAmountResult] | None = None


@dataclass(slots=True)
class GoodsSheetParseResult:
    goods_data: GoodsData | None = None
    imcp_avg: Decimal | None = None
