from decimal import Decimal
from io import BytesIO

from openpyxl import Workbook

from app.domain.entities.imcp import IMCPResult, VatInfo
from app.domain.entities.imcp.goods_result import (
    GoodsData,
    GoodsItem,
    Organization,
    PriceCell,
)
from app.domain.entities.imcp.works_result import WorksOrgAmountResult
from app.infrastructure.extraction.imcp import OpenpyxlExcelIMCPExtractor
from app.infrastructure.extraction.imcp.imcp_extractor import (
    _is_placeholder,
    _normalize_name,
    _normalize_organization_names,
    _pick_normalized_name,
)


_ANCHOR_AVG = (
    "Начальная (максимальная) цена договора (среднее арифметическое значение) "
    "составляет, руб. в т.ч. НДС 20%"
)
_ANCHOR_LIMIT = (
    "Начальная (максимальная) цена договора (с учетом лимита финансирования) "
    "составляет, руб."
)
_SHEET_WORKS_SERVICES = "Расчет при закупке работ, услуг"


def _wb_to_bytes(wb: Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_extract_initial_max_contract_price_happy_path_two_sheets() -> None:
    wb = Workbook()

    ws_other = wb.active
    ws_other.title = "Данные листа"
    ws_other["A1"] = "Служебная информация"

    ws_ws = wb.create_sheet(_SHEET_WORKS_SERVICES)
    ws_ws["A1"] = "№ п/п"
    ws_ws["B1"] = "Организации"
    ws_ws["C1"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["A2"] = "1"
    ws_ws["B2"] = "ООО Альфа"
    ws_ws["C2"] = "12 331 372,00"
    ws_ws["A3"] = "2"
    ws_ws["B3"] = "ООО Бета"
    ws_ws["C3"] = "12 000 000,00"
    ws_ws["B4"] = _ANCHOR_AVG
    ws_ws["C4"] = Decimal("12331372.00")
    ws_ws["B5"] = _ANCHOR_LIMIT
    ws_ws["C5"] = "12 000 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg is None
    assert result.goods_data is None
    assert result.works_imcp_avg == Decimal("12331372.00")
    assert result.works_imcp_limit == Decimal("12000000.00")
    assert result.works_orgs_amounts is not None


def test_extract_works_org_amounts_keeps_order_and_amounts() -> None:
    wb = Workbook()

    ws_other = wb.active
    ws_other.title = "Прочий лист"
    ws_other["A1"] = "Служебно"

    ws_ws = wb.create_sheet(_SHEET_WORKS_SERVICES)
    ws_ws["A1"] = "№ п/п"
    ws_ws["B1"] = "Организации"
    ws_ws["C1"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["A2"] = "1"
    ws_ws["B2"] = "ООО Гамма Сервис"
    ws_ws["C2"] = "13 594 116,00"
    ws_ws["A3"] = "2"
    ws_ws["B3"] = "ООО ДЕЛЬТАСТРОЙ"
    ws_ws["C3"] = "12 000 000,00"
    ws_ws["A4"] = "3"
    ws_ws["B4"] = "ИП Иванов И.И."
    ws_ws["C4"] = "11 400 000,00"
    ws_ws["B5"] = _ANCHOR_AVG
    ws_ws["C5"] = "13 000 000,00"
    ws_ws["B6"] = _ANCHOR_LIMIT
    ws_ws["C6"] = "12 000 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.works_orgs_amounts is not None
    organizations = result.works_orgs_amounts
    assert [org.name for org in organizations] == [
        "ООО Гамма Сервис",
        "ООО ДЕЛЬТАСТРОЙ",
        "ИП Иванов И.И.",
    ]
    assert [org.id for org in organizations] == [1, 2, 3]
    assert [org.amount for org in organizations] == [
        Decimal("13594116.00"),
        Decimal("12000000.00"),
        Decimal("11400000.00"),
    ]


def test_extract_imcp_keeps_works_only_partial_contract() -> None:
    wb = Workbook()
    ws_other = wb.active
    ws_other.title = "Данные листа"

    ws_ws = wb.create_sheet("Лист с расчетом")
    ws_ws["A1"] = "№ п/п"
    ws_ws["B1"] = "Организации"
    ws_ws["C1"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["A2"] = "1"
    ws_ws["B2"] = "ООО Альфа"
    ws_ws["C2"] = "1 111 000,00"
    ws_ws["B3"] = _ANCHOR_AVG
    ws_ws["C3"] = "2 000 000,00"
    ws_ws["B4"] = _ANCHOR_LIMIT
    ws_ws["C4"] = "1 900 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg is None
    assert result.goods_data is None
    assert result.works_imcp_avg == Decimal("2000000.00")
    assert result.works_imcp_limit == Decimal("1900000.00")
    assert result.works_orgs_amounts is not None
    assert [org.name for org in result.works_orgs_amounts] == ["ООО Альфа"]


def test_extract_imcp_limit_with_short_fallback_anchor() -> None:
    wb = Workbook()
    ws_other = wb.active
    ws_other.title = "Данные листа"

    ws_ws = wb.create_sheet(_SHEET_WORKS_SERVICES)
    ws_ws["A1"] = "№ п/п"
    ws_ws["B1"] = "Организации"
    ws_ws["C1"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["A2"] = "1"
    ws_ws["B2"] = "ООО Альфа"
    ws_ws["C2"] = "1 111 000,00"
    ws_ws["B3"] = _ANCHOR_AVG
    ws_ws["C3"] = "2 000 000,00"
    ws_ws["B4"] = "Лимит финансирования"
    ws_ws["C4"] = "1 900 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.works_imcp_avg == Decimal("2000000.00")
    assert result.works_imcp_limit == Decimal("1900000.00")


def test_extract_imcp_ignores_non_works_sheet_for_works_result() -> None:
    wb = Workbook()
    ws_non_works = wb.active
    ws_non_works.title = "Расчет при закупке товаро"
    ws_non_works["B3"] = _ANCHOR_AVG
    ws_non_works["C3"] = "1000,00"

    ws_ws = wb.create_sheet("Расчет при закупке работ и услуг")
    ws_ws["A1"] = "№ п/п"
    ws_ws["B1"] = "Организации"
    ws_ws["C1"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["A2"] = "1"
    ws_ws["B2"] = "ООО Бета"
    ws_ws["C2"] = "1400,00"
    ws_ws["B3"] = _ANCHOR_AVG
    ws_ws["C3"] = "2000,00"
    ws_ws["B4"] = _ANCHOR_LIMIT
    ws_ws["C4"] = "1500,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg is None
    assert result.works_imcp_avg == Decimal("2000.00")
    assert result.works_imcp_limit == Decimal("1500.00")
    assert result.works_orgs_amounts is not None
    assert [org.name for org in result.works_orgs_amounts] == ["ООО Бета"]


def test_normalize_name_basic_cleanup() -> None:
    name = " ООО Ёшкин  кот"
    normalized_name = _normalize_name(name)
    gold_normalized_name = "ооо ешкин кот"
    assert normalized_name == gold_normalized_name


def test_is_placeholder_detect_empties_and_templates() -> None:
    assert _is_placeholder("") is True
    assert _is_placeholder("-") is True
    assert _is_placeholder("орг-я") is True
    assert _is_placeholder("ООО Эпсилон") is False
    assert _is_placeholder("Организация") is True
    assert _is_placeholder("Контрагент") is True
    assert _is_placeholder("Самозанятый Иванов") is False
    assert _is_placeholder("Организаци") is True
    assert _is_placeholder("Организация 3") is True
    assert _is_placeholder("Орг-ция 2") is True
    assert _is_placeholder("контрагент 1") is True

    assert _is_placeholder('ООО "Эгис"') is False
    assert _is_placeholder("ИП Иванов Иван Иванович") is False


def test_pick_normalize_org_name() -> None:
    names = ["Орг-ция 1", 'ООО "НВК Монтаж"']
    assert _pick_normalized_name(names) == 'ООО "НВК Монтаж"'

    names = ["Орг-ция 2", 'ООО "Эгис"']
    assert _pick_normalized_name(names) == 'ООО "Эгис"'

    names = ["организация 3", 'ООО "ЮГ Поволжье"']
    assert _pick_normalized_name(names) == 'ООО "ЮГ Поволжье"'

    names = ["фильтрационные технологии"]
    assert _pick_normalized_name(names) == "фильтрационные технологии"

    names = ["ИП Иванов Иван Иванович"]
    assert _pick_normalized_name(names) == "ИП Иванов Иван Иванович"

    names = ["поставщик 1"]
    assert _pick_normalized_name(names) == "поставщик 1"


def test_normalize_organization_name_sync_goods_and_works_sheet() -> None:
    result = IMCPResult(
        goods_imcp_avg=Decimal(1000),
        goods_vat_info=VatInfo(rate=20, included=True),
        works_imcp_avg=Decimal(1000),
        works_imcp_limit=Decimal(1200),
        goods_data=GoodsData(
            organizations=[
                Organization(
                    id=1,
                    name="Орг-ция 1",
                ),
                Organization(
                    id=2,
                    name="Орг-ция 2",
                ),
                Organization(
                    id=3,
                    name="Орг-ция 3",
                ),
            ],
            goods_items=[
                GoodsItem(
                    id=1,
                    name="ТО АПС",
                    unit_of_measure="месяц",
                    unit_imcp=Decimal(100),
                    quantity=Decimal(1),
                    item_imcp=Decimal(100),
                )
            ],
            price_cells=[
                PriceCell(organization_id=1, item_id=1, price=Decimal(100)),
                PriceCell(organization_id=2, item_id=1, price=Decimal(100)),
                PriceCell(organization_id=3, item_id=1, price=Decimal(100)),
            ],
        ),
        works_orgs_amounts=[
            WorksOrgAmountResult(id=1, name="ООО Дзета Монтаж", amount=Decimal(1000)),
            WorksOrgAmountResult(id=2, name="ООО Эта", amount=Decimal(1000)),
            WorksOrgAmountResult(id=3, name='ООО "ЮГ Поволжье"', amount=Decimal(1000)),
        ],
    )

    normalized = _normalize_organization_names(result)

    assert normalized.goods_data is not None
    assert normalized.works_orgs_amounts is not None
    assert [org.name for org in normalized.goods_data.organizations] == [
        "ООО Дзета Монтаж",
        "ООО Эта",
        'ООО "ЮГ Поволжье"',
    ]
    assert [org.name for org in normalized.works_orgs_amounts] == [
        "ООО Дзета Монтаж",
        "ООО Эта",
        'ООО "ЮГ Поволжье"',
    ]
