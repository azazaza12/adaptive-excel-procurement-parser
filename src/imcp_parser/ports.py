from abc import ABC, abstractmethod

from .entities import IMCPResult


class IExcelIMCPExtractor(ABC):
    @abstractmethod
    def extract_imcp(self, *, content: bytes) -> IMCPResult: ...
