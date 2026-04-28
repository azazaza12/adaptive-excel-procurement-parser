from __future__ import annotations

from copy import deepcopy

from app.infrastructure.extraction.imcp.dynamic_table_extractor import (
    AddOrgAmountHandler,
    Aggregator,
    DynamicTableExtractor,
    EmptyRowFactory,
    IsValidRowPredicate,
    ItemNameExtractor,
    MetricExtractor,
    OrgAmountExtractor,
    SetFieldHandler,
    SetItemNameHandler,
)
from app.infrastructure.extraction.imcp.models import (
    ExtractorConfig,
    HeaderAnchorPattern,
)


class ExtractorBuilder[ResultT]:
    def __init__(self) -> None:
        self._header_anchors: dict[str, HeaderAnchorPattern] = {}
        self._stop_row_anchors: tuple[str, ...] = ()
        self._stop_column_anchors: tuple[str, ...] = ()
        self._org_header_anchors: tuple[str, ...] = ()
        self._header_scan_depth = 2
        self._price_block_parent_anchors: tuple[str, ...] = ()
        self._price_block_service_subheaders: tuple[str, ...] = ()
        self._max_empty_rows_in_series = 3

        self._item_name_extractor: ItemNameExtractor | None = None
        self._metric_extractor: MetricExtractor | None = None
        self._org_amount_extractor: OrgAmountExtractor | None = None
        self._is_valid_row: IsValidRowPredicate | None = None
        self._aggregate: Aggregator[ResultT] | None = None
        self._set_field: SetFieldHandler | None = None
        self._add_org_amount: AddOrgAmountHandler | None = None
        self._set_item_name: SetItemNameHandler | None = None
        self._empty_row_factory: EmptyRowFactory | None = None

    def with_header_anchors(
        self,
        anchors: dict[str, HeaderAnchorPattern],
    ) -> ExtractorBuilder[ResultT]:
        self._header_anchors = deepcopy(anchors)
        return self

    def with_header_anchor(
        self,
        field_name: str,
        anchor: HeaderAnchorPattern,
    ) -> ExtractorBuilder[ResultT]:
        self._header_anchors[field_name] = anchor
        return self

    def with_stop_row_anchors(self, *anchors: str) -> ExtractorBuilder[ResultT]:
        self._stop_row_anchors = tuple(anchors)
        return self

    def with_stop_column_anchors(self, *anchors: str) -> ExtractorBuilder[ResultT]:
        self._stop_column_anchors = tuple(anchors)
        return self

    def with_org_header_anchors(self, *anchors: str) -> ExtractorBuilder[ResultT]:
        self._org_header_anchors = tuple(anchors)
        return self

    def with_header_scan_depth(self, depth: int) -> ExtractorBuilder[ResultT]:
        self._header_scan_depth = depth
        return self

    def with_price_block_parent_anchors(
        self,
        *anchors: str,
    ) -> ExtractorBuilder[ResultT]:
        self._price_block_parent_anchors = tuple(anchors)
        return self

    def with_price_block_service_subheaders(
        self,
        *subheaders: str,
    ) -> ExtractorBuilder[ResultT]:
        self._price_block_service_subheaders = tuple(subheaders)
        return self

    def with_max_empty_rows_in_series(
        self,
        max_empty_rows_in_series: int,
    ) -> ExtractorBuilder[ResultT]:
        self._max_empty_rows_in_series = max_empty_rows_in_series
        return self

    def with_item_name_extractor(
        self,
        extractor: ItemNameExtractor,
    ) -> ExtractorBuilder[ResultT]:
        self._item_name_extractor = extractor
        return self

    def with_metric_extractor(
        self,
        extractor: MetricExtractor,
    ) -> ExtractorBuilder[ResultT]:
        self._metric_extractor = extractor
        return self

    def with_org_amount_extractor(
        self,
        extractor: OrgAmountExtractor,
    ) -> ExtractorBuilder[ResultT]:
        self._org_amount_extractor = extractor
        return self

    def with_is_valid_row(
        self,
        predicate: IsValidRowPredicate,
    ) -> ExtractorBuilder[ResultT]:
        self._is_valid_row = predicate
        return self

    def with_aggregator(
        self, aggregator: Aggregator[ResultT]
    ) -> ExtractorBuilder[ResultT]:
        self._aggregate = aggregator
        return self

    def with_set_field_handler(
        self,
        handler: SetFieldHandler,
    ) -> ExtractorBuilder[ResultT]:
        self._set_field = handler
        return self

    def with_add_org_amount_handler(
        self,
        handler: AddOrgAmountHandler,
    ) -> ExtractorBuilder[ResultT]:
        self._add_org_amount = handler
        return self

    def with_set_item_name_handler(
        self,
        handler: SetItemNameHandler,
    ) -> ExtractorBuilder[ResultT]:
        self._set_item_name = handler
        return self

    def with_empty_row_factory(
        self,
        factory: EmptyRowFactory,
    ) -> ExtractorBuilder[ResultT]:
        self._empty_row_factory = factory
        return self

    def build(self) -> DynamicTableExtractor[ResultT]:
        if not self._header_anchors:
            msg = "header_anchors is required"
            raise ValueError(msg)

        config = ExtractorConfig(
            header_anchors=deepcopy(self._header_anchors),
            stop_row_anchors=self._stop_row_anchors,
            stop_column_anchors=self._stop_column_anchors,
            org_header_anchors=self._org_header_anchors,
            header_scan_depth=self._header_scan_depth,
            price_block_parent_anchors=self._price_block_parent_anchors,
            price_block_service_subheaders=self._price_block_service_subheaders,
            max_empty_rows_in_series=self._max_empty_rows_in_series,
        )

        return DynamicTableExtractor(
            config,
            item_name_extractor=self._item_name_extractor,
            metric_extractor=self._metric_extractor,
            org_amount_extractor=self._org_amount_extractor,
            is_valid_row=self._is_valid_row,
            aggregate=self._aggregate,
            set_field=self._set_field,
            add_org_amount=self._add_org_amount,
            set_item_name=self._set_item_name,
            empty_row_factory=self._empty_row_factory,
        )
