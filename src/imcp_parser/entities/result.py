from dataclasses import dataclass
from decimal import Decimal

from pydantic import BaseModel

from app.domain.entities.imcp.goods_result import GoodsData
from app.domain.entities.imcp.works_result import (
    WorksOrgAmountResult,
)


@dataclass
class VatInfo:
    rate: int | None
    included: bool | None


class IMCPResult(BaseModel):
    goods_imcp_avg: Decimal | None = None
    goods_vat_info: VatInfo | None = None
    works_imcp_avg: Decimal | None = None
    works_imcp_limit: Decimal | None = None
    goods_data: GoodsData | None = None
    works_orgs_amounts: list[WorksOrgAmountResult] | None = None
    works_amount_header_vat_info: VatInfo | None = None
    works_imcp_avg_row_vat_info: VatInfo | None = None
