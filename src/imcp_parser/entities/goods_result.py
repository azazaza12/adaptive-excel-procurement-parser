from decimal import Decimal

from pydantic import BaseModel


class GoodsOrgResult(BaseModel):
    id: int
    name: str


class GoodsItemResult(BaseModel):
    id: int
    name: str
    unit_of_measure: str | None = None
    unit_imcp: Decimal | None = None
    quantity: Decimal | None = None
    item_imcp: Decimal | None = None


class GoodsOrgAmountResult(BaseModel):
    organization_id: int
    item_id: int
    amount: Decimal


class GoodsOrgsAmountsResult(BaseModel):
    organizations: list[GoodsOrgResult]
    goods_items: list[GoodsItemResult]
    cells: list[GoodsOrgAmountResult]
