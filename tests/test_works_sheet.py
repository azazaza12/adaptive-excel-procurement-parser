from decimal import Decimal

from app.infrastructure.extraction.imcp.works_sheet import WorksSheet
from openpyxl import Workbook


_ANCHOR_AVG = (
    "Начальная (максимальная) цена договора (среднее арифметическое значение) "
    "составляет, руб. в т.ч. НДС 20%"
)
_ANCHOR_LIMIT = (
    "Начальная (максимальная) цена договора (с учетом лимита финансирования) "
    "составляет, руб."
)


def test_works_sheet_is_me_by_title() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Расчет при закупке работ, услуг"

    assert WorksSheet.is_me(ws)


def test_works_sheet_is_me_by_header_fallback() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист с расчетом"
    ws["B5"] = "Организации"
    ws["C5"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."

    assert WorksSheet.is_me(ws)


def test_works_sheet_is_not_me_for_goods_like_sheet() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Расчет при закупке товаров"
    ws["B3"] = _ANCHOR_AVG
    ws["C3"] = "1000,00"

    assert not WorksSheet.is_me(ws)


def test_works_sheet_extract_keeps_order_amounts_and_imcp_values() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист с расчетом"
    ws["A1"] = "№ п/п"
    ws["B1"] = "Организации"
    ws["C1"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws["A2"] = "1"
    ws["B2"] = "ООО Гамма Сервис"
    ws["C2"] = "13 594 116,00"
    ws["A3"] = "2"
    ws["B3"] = "ООО ДЕЛЬТАСТРОЙ"
    ws["C3"] = "12 000 000,00"
    ws["A4"] = "3"
    ws["B4"] = "ИП Иванов И.И."
    ws["C4"] = "11 400 000,00"
    ws["B5"] = _ANCHOR_AVG
    ws["C5"] = "13 000 000,00"
    ws["B6"] = _ANCHOR_LIMIT
    ws["C6"] = "12 000 000,00"

    result = WorksSheet(ws).extract()

    assert result.imcp_avg == Decimal("13000000.00")
    assert result.imcp_limit == Decimal("12000000.00")
    assert result.orgs_amounts is not None
    assert [org.name for org in result.orgs_amounts] == [
        "ООО Гамма Сервис",
        "ООО ДЕЛЬТАСТРОЙ",
        "ИП Иванов И.И.",
    ]
    assert [org.id for org in result.orgs_amounts] == [1, 2, 3]
    assert [org.amount for org in result.orgs_amounts] == [
        Decimal("13594116.00"),
        Decimal("12000000.00"),
        Decimal("11400000.00"),
    ]
