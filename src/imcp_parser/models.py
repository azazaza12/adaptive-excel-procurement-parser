from dataclasses import dataclass, field
from decimal import Decimal

from app.domain.entities.imcp.goods_result import GoodsOrgsAmountsResult
from app.domain.entities.imcp.works_result import WorksOrgsAmountsResult


type HeaderAnchorPattern = str | tuple[str, ...]


@dataclass(slots=True)
class HeaderColumns:
    header_row_idx: int
    item_col_idx: int
    metric_columns: dict[str, int] = field(default_factory=dict)


@dataclass(slots=True)
class OrganizationInfo:
    col_idx: int
    name: str
    id: int


@dataclass(slots=True)
class ParsingContext:
    header: HeaderColumns
    organizations: list[OrganizationInfo] = field(default_factory=list)
    empty_rows_in_series: int = 0


@dataclass(slots=True)
class ExtractorConfig:
    header_anchors: dict[str, HeaderAnchorPattern]
    stop_row_anchors: tuple[str, ...] = ()
    stop_column_anchors: tuple[str, ...] = ()
    org_header_anchors: tuple[str, ...] = ()
    header_scan_depth: int = 2
    price_block_parent_anchors: tuple[str, ...] = ()
    price_block_service_subheaders: tuple[str, ...] = ()
    max_empty_rows_in_series: int = 3


@dataclass(slots=True)
class GoodsSheetParseResult:
    imcp_avg: Decimal | None = None
    orgs_amounts: GoodsOrgsAmountsResult | None = None


@dataclass(slots=True)
class WorksSheetParseResult:
    imcp_avg: Decimal | None = None
    imcp_limit: Decimal | None = None
    orgs_amounts: WorksOrgsAmountsResult | None = None
