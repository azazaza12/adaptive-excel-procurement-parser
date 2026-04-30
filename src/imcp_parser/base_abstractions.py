from abc import ABC, abstractmethod
from typing import Any

from app.infrastructure.extraction.imcp.common import norm_text

_MIN_ANCHOR_WORD_LEN = 1


class BaseAnchor(ABC):
    @abstractmethod
    def matches(self, value: str) -> bool:
        pass


class MatchAnchor(BaseAnchor):
    def __init__(self, anchor: str) -> None:
        self._anchor = anchor

    def matches(self, value: Any) -> bool:
        if not isinstance(value, str):
            return False
        cell_norm = norm_text(value)
        anchor_words = [
            word
            for word in norm_text(self._anchor).split()
            if len(word) >= _MIN_ANCHOR_WORD_LEN
        ]
        if not anchor_words:
            return False
        cell_words = cell_norm.split()
        for anchor_word in anchor_words:
            if not any(word.startswith(anchor_word) for word in cell_words):
                return False
        return True


class ExactAnchor(BaseAnchor):
    def __init__(self, anchor: str) -> None:
        self._anchor_norm = norm_text(anchor)

    def matches(self, value: str) -> bool:
        if not isinstance(value, str):
            return False
        return norm_text(value) == self._anchor_norm


class AnyOfAnchors(BaseAnchor):
    def __init__(self, *anchors: BaseAnchor) -> None:
        self._anchors = anchors

    def matches(self, value: Any) -> bool:
        if not isinstance(value, str):
            return False
        return any(anchor.matches(value) for anchor in self._anchors)
