from app.infrastructure.extraction.imcp.base_abstractions import (
    AnyOfAnchors,
    ExactAnchor,
    MatchAnchor,
)


def test_match_anchor_matches_words_by_prefix() -> None:
    anchor = MatchAnchor("наименовани товар")

    assert anchor.matches("Наименование товара")
    assert anchor.matches("наименования товаров и услуг")
    assert not anchor.matches("Единица измерения")


def test_exact_anchor_matches_normalized_value() -> None:
    anchor = ExactAnchor("  Стоимость работ ")
    assert anchor.matches("стоимость работ")
    assert not anchor.matches("Стоимость услуг")


def test_any_of_anchors_matches_when_one_anchor_matches() -> None:
    anchor = AnyOfAnchors(MatchAnchor("наименование"), MatchAnchor("цена"))

    assert anchor.matches("Цена предложения, руб.")
    assert not anchor.matches("Ед. изм.")
