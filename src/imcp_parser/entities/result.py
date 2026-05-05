from decimal import Decimal

from pydantic import BaseModel

from app.domain.entities.imcp.goods_result import (
    GoodsOrgsAmountsResult,
)
from app.domain.entities.imcp.works_result import (
    WorksOrgAmountResult,
)


class IMCPResult(BaseModel):
    goods_imcp_avg: Decimal | None = None
    works_imcp_avg: Decimal | None = None
    works_imcp_limit: Decimal | None = None
    goods_orgs_amounts: GoodsOrgsAmountsResult | None = None
    works_orgs_amounts: list[WorksOrgAmountResult] | None = None
