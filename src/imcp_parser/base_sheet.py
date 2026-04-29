from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Generic, TypeVar, ClassVar

from openpyxl.worksheet.worksheet import Worksheet
from app.domain.entities.imcp.result import IMCPResult
from app.infrastructure.extraction.imcp.base_abstractions import BaseAnchor

SheetResultT = TypeVar("SheetResultT")

@dataclass
class HeaderCell(slots=True):
    row_idx: int
    col_idx: int


class BaseSheet(ABC, Generic[SheetResultT]):
    _registry: ClassVar[list[type["BaseSheet"]]] = []


    def __init__(self, ws: Worksheet) -> None:
        self.ws = ws

    
    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        if getattr(cls, "__abstractmethods__", False):
            return
        BaseSheet._registry.append(cls)

    
    @classmethod
    def registered(cls) -> tuple[type["BaseSheet"], ...]:
        return tuple(BaseSheet._registry)

    @classmethod
    @abstractmethod
    def is_me(cls, ws: Worksheet) -> bool:
        pass

    @abstractmethod
    def extract(self) -> SheetResultT:
        pass


    def _find_header_by(self, anchor: BaseAnchor) -> HeaderCell | None:
        for row_idx, row in enumerate(self.ws.iter_rows(), start=1):
            for cell in row:
                value = cell.value
                if anchor.matches(value):
                    return HeaderCell(row_idx=row_idx, col_idx=cell.column)
        return None

    @abstractmethod
    def apply_to_result(self, target: IMCPResult, extracted: SheetResultT) -> None:
        pass

    

