from decimal import Decimal

from pydantic import BaseModel

from app.domain.entities.excel.goods_result import (
    GoodsOrgsAmountsResult,
)
from app.domain.entities.excel.works_services_result import (
    WorksOrgsAmountsResult,
)


class ExcelIMCPResult(BaseModel):
    goods_imcp_avg: Decimal | None = None
    works_imcp_avg: Decimal | None = None
    works_imcp_limit: Decimal | None = None
    goods_orgs_amounts: GoodsOrgsAmountsResult | None = None
    works_orgs_amounts: WorksOrgsAmountsResult | None = None
