from abc import ABC, abstractmethod
from typing import Generic, TypeVar, ClassVar

from openpyxl.worksheet.worksheet import Worksheet

SheetResultT = TypeVar("SheetResultT")


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
    @abstractmethod
    def is_me(cls, ws: Worksheet) -> bool:
        pass

    @abstractmethod
    def extract(self) -> SheetResultT:
        pass

    
    @abstractmethod
    def apply_to_result(self, target: ExcelIMCPResult, extracted: SheetResultT,) - None
    pass

