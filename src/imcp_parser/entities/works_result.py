from decimal import Decimal

from pydantic import BaseModel


class WorksOrgAmountResult(BaseModel):
    id: int
    name: str
    amount: Decimal


class WorksOrgsAmountsResult(BaseModel):
    organizations: list[WorksOrgAmountResult]
