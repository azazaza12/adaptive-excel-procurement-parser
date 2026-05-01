from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import ClassVar, TypeVar

from openpyxl.worksheet.worksheet import Worksheet

from app.infrastructure.extraction.imcp.base_abstractions import BaseAnchor


SheetResultT = TypeVar("SheetResultT")


@dataclass(slots=True)
class HeaderCell:
    row_idx: int
    col_idx: int


class BaseSheet[SheetResultT](ABC):
    _registry: ClassVar[list[type[BaseSheet]]] = []

    def __init__(self, ws: Worksheet) -> None:
        self.ws = ws

    def __init_subclass__(cls, **kwargs: object) -> None:
        super().__init_subclass__(**kwargs)

        if getattr(cls, "__abstractmethods__", False):
            return
        BaseSheet._registry.append(cls)

    @classmethod
    def registered(cls) -> tuple[type[BaseSheet], ...]:
        return tuple(BaseSheet._registry)

    @classmethod
    @abstractmethod
    def is_me(cls, ws: Worksheet) -> bool:
        pass

    @abstractmethod
    def extract(self) -> SheetResultT:
        pass

    def row_has_anchor(self, *, row_idx: int, anchor: BaseAnchor) -> bool:
        for col_idx in range(1, self.ws.max_column + 1):
            value = self.ws.cell(row=row_idx, column=col_idx).value
            if anchor.matches(value):
                return True
        return False

    def _find_header_by(self, anchor: BaseAnchor) -> HeaderCell | None:
        for row_idx, row in enumerate(self.ws.iter_rows(), start=1):
            for cell in row:
                value = cell.value
                if anchor.matches(value):
                    return HeaderCell(row_idx=row_idx, col_idx=cell.column)
        return None

    def _has_title_anchors(self, title_anchors: tuple[BaseAnchor, ...]) -> bool:
        return all(anchor.matches(self.ws.title) for anchor in title_anchors)

    def _has_table_header(self, anchors: tuple[BaseAnchor, ...]) -> bool:
        headers = [self._find_header_by(anchor) for anchor in anchors]
        return bool(all(header is not None for header in headers))
