from abc import ABC, abstractmethod
from typing import Generic, TypeVar

from openpyxl.worksheet.worksheet import Worksheet

SheetResultT = TypeVar("SheetResultT")


class BaseSheet(ABC, Generic[SheetResultT]):
    def __init__(self, ws: Worksheet) -> None:
        self.ws = ws

    @classmethod
    @abstractmethod
    def is_me(cls, ws: Worksheet) -> bool:
        pass

    @abstractmethod
    def extract(self) -> SheetResultT:
        pass
