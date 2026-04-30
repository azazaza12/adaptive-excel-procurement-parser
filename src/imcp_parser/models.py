from dataclasses import dataclass
from decimal import Decimal

from app.domain.entities.imcp.works_result import WorksOrgAmountResult


@dataclass(slots=True)
class WorksSheetParseResult:
    imcp_avg: Decimal | None = None
    imcp_limit: Decimal | None = None
    orgs_amounts: list[WorksOrgAmountResult] | None = None
