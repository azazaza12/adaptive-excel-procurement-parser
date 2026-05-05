from decimal import Decimal

from pydantic import BaseModel


class Source(BaseModel):
    id: int
    name: str


class GoodsItem(BaseModel):
    id: int
    name: str
    unit_of_measure: str | None = None
    unit_imcp: Decimal | None = None
    quantity: Decimal | None = None
    item_imcp: Decimal | None = None


class PriceCell(BaseModel):
    source_id: int
    item_id: int
    price: Decimal


class GoodsData(BaseModel):
    sources: list[Source]
    goods_items: list[GoodsItem]
    price_cells: list[PriceCell]
