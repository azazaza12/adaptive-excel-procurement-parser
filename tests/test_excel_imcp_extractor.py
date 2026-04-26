from decimal import Decimal
from io import BytesIO

from app.infrastructure.extraction.excel_imcp_extractor import (
    OpenpyxlExcelIMCPExtractor,
)
from openpyxl import Workbook


_ANCHOR_AVG = (
    "Начальная (максимальная) цена договора (среднее арифметическое значение) "
    "составляет, руб. в т.ч. НДС 20%"
)
_ANCHOR_LIMIT = (
    "Начальная (максимальная) цена договора (с учетом лимита финансирования) "
    "составляет, руб."
)
_SHEET_GOODS = "Расчет при закупке товаров"
_SHEET_WORKS_SERVICES = "Расчет при закупке товаров, услуг"
_EXPECTED_GOODS_ITEMS_COUNT = 2
_EXPECTED_ORGS_COUNT = 3
_EXPECTED_CELLS_COUNT = 3


def _wb_to_bytes(wb: Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_extract_initial_max_contract_price_happy_path_two_sheets() -> None:
    wb = Workbook()

    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS
    ws_goods["B7"] = _ANCHOR_AVG
    ws_goods["C7"] = "12 331 372,00"

    ws_ws = wb.create_sheet(_SHEET_WORKS_SERVICES)
    ws_ws["B5"] = _ANCHOR_AVG
    ws_ws["C5"] = Decimal("12331372.00")
    ws_ws["B6"] = _ANCHOR_LIMIT
    ws_ws["C6"] = "12 000 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg == Decimal("12331372.00")
    assert result.works_imcp_avg == Decimal("12331372.00")
    assert result.works_imcp_limit == Decimal("12000000.00")


def test_extract_initial_max_contract_price_fallback_sheet_name_by_anchor() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = "Расчет при закупке товаро"  # typo
    ws_goods["B3"] = _ANCHOR_AVG
    ws_goods["D3"] = "1000,50"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg == Decimal("1000.50")


def test_extract_goods_org_amounts_matrix() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C2"] = "Орг-пред 1"
    ws_goods["D2"] = "Орг-пред 2"
    ws_goods["E2"] = "Орг-пред 3"
    ws_goods["F1"] = "нмц единицы товара"
    ws_goods["G1"] = "Ед. изм."
    ws_goods["H1"] = "Кол."
    ws_goods["I1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО АПС"
    ws_goods["C3"] = "701 595"
    ws_goods["D3"] = "681 220"
    ws_goods["E3"] = "651 933"
    ws_goods["F3"] = "678 249,33"
    ws_goods["G3"] = "месяц"
    ws_goods["H3"] = "12"
    ws_goods["I3"] = "8 138 992,00"

    ws_goods["B4"] = "ТО СОУЭ"
    ws_goods["C4"] = "260238"
    ws_goods["D4"] = "290 580"
    ws_goods["E4"] = "267780"

    ws_goods["B6"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"

    ws_goods["B8"] = _ANCHOR_AVG
    ws_goods["C8"] = "12 331 372,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.name for org in result.goods_orgs_amounts.organizations] == [
        "Орг-пред 1",
        "Орг-пред 2",
        "Орг-пред 3",
    ]
    assert [org.id for org in result.goods_orgs_amounts.organizations] == [
        1,
        2,
        3,
    ]
    assert len(result.goods_orgs_amounts.goods_items) == _EXPECTED_GOODS_ITEMS_COUNT

    first_item = result.goods_orgs_amounts.goods_items[0]
    assert first_item.id == 1
    assert first_item.name == "ТО АПС"
    assert first_item.unit_of_measure == "месяц"
    assert first_item.unit_imcp == Decimal("678249.33")
    assert first_item.quantity == Decimal(12)
    assert first_item.item_imcp == Decimal("8138992.00")

    item_1_cells = [
        cell for cell in result.goods_orgs_amounts.cells if cell.item_id == 1
    ]
    assert len(item_1_cells) == _EXPECTED_CELLS_COUNT
    assert item_1_cells[0].organization_id == 1
    assert item_1_cells[0].amount == Decimal(701595)
    assert item_1_cells[2].organization_id == _EXPECTED_ORGS_COUNT
    assert item_1_cells[2].amount == Decimal(651933)


def test_extract_goods_org_amounts_skips_non_numeric_and_partial_rows() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C2"] = "Орг-пред 1"
    ws_goods["D2"] = "Орг-пред 2"
    ws_goods["E1"] = (
        "НМЦ единицы товара, в руб. с НДС (среднее арифметическое значение)"
    )
    ws_goods["F1"] = "Ед. изм."
    ws_goods["G1"] = "Кол."
    ws_goods["H1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО СКУД"
    ws_goods["C3"] = "8 500"
    ws_goods["D3"] = "N/A"
    ws_goods["E3"] = "9 100"
    ws_goods["F3"] = "месяц"
    ws_goods["G3"] = "12"
    ws_goods["H3"] = "109200"

    ws_goods["B4"] = "ТО СОТС"
    ws_goods["C4"] = None
    ws_goods["D4"] = "13 400"
    ws_goods["E4"] = "13 400"
    ws_goods["F4"] = "месяц"
    ws_goods["G4"] = "12"
    ws_goods["H4"] = "160800"

    ws_goods["B6"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B7"] = _ANCHOR_AVG
    ws_goods["C7"] = "100 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.id for org in result.goods_orgs_amounts.organizations] == [
        1,
        2,
    ]
    assert len(result.goods_orgs_amounts.goods_items) == _EXPECTED_GOODS_ITEMS_COUNT
    first_item = result.goods_orgs_amounts.goods_items[0]
    assert first_item.unit_imcp == Decimal(9100)
    second_item = result.goods_orgs_amounts.goods_items[1]
    assert second_item.item_imcp == Decimal(160800)
    item_1_cells = [
        cell for cell in result.goods_orgs_amounts.cells if cell.item_id == 1
    ]
    assert len(item_1_cells) == 1
    assert {cell.organization_id for cell in item_1_cells} == {1}


def test_extract_goods_item_imcp_with_short_nmc_header_fallback() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C2"] = "Орг-пред 1"
    ws_goods["D1"] = "НМЦ единицы товара"
    ws_goods["E1"] = "Ед. изм."
    ws_goods["F1"] = "Кол."
    ws_goods["G1"] = "НМЦ"

    ws_goods["B3"] = "ТО ВПС"
    ws_goods["C3"] = "1000"
    ws_goods["D3"] = "123,45"
    ws_goods["E3"] = "шт."
    ws_goods["F3"] = "2"
    ws_goods["G3"] = "246,90"

    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B6"] = _ANCHOR_AVG
    ws_goods["C6"] = "1 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    first_item = result.goods_orgs_amounts.goods_items[0]
    assert first_item.unit_imcp == Decimal("123.45")
    assert first_item.item_imcp == Decimal("246.90")


def test_extract_goods_org_amounts_accepts_supplier_headers_under_price_parent() -> (
    None
):
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C2"] = "Поставщик 1"
    ws_goods["D2"] = "Поставщик 2"
    ws_goods["E1"] = "Цена единицы товара, руб."
    ws_goods["F1"] = "Ед. изм."
    ws_goods["G1"] = "Кол."
    ws_goods["H1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО АПС"
    ws_goods["C3"] = "701 595"
    ws_goods["D3"] = "681 220"
    ws_goods["E3"] = "678 249,33"

    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B6"] = _ANCHOR_AVG
    ws_goods["C6"] = "1 000 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.name for org in result.goods_orgs_amounts.organizations] == [
        "Поставщик 1",
        "Поставщик 2",
    ]


def test_extract_goods_org_amounts_accepts_typical_org_headers() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C2"] = "ООО Йота"
    ws_goods["D2"] = "АО Тета"
    ws_goods["E2"] = "ИП Иванов"
    ws_goods["F1"] = "Цена единицы товара, руб."
    ws_goods["G1"] = "Ед. изм."
    ws_goods["H1"] = "Кол."
    ws_goods["I1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО СКУД"
    ws_goods["C3"] = "8 500"
    ws_goods["D3"] = "10 200"
    ws_goods["E3"] = "8 337"
    ws_goods["F3"] = "9 012,33"
    ws_goods["G3"] = "месяц"
    ws_goods["H3"] = "12"
    ws_goods["I3"] = "108147,96"

    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B6"] = _ANCHOR_AVG
    ws_goods["C6"] = "100 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.name for org in result.goods_orgs_amounts.organizations] == [
        "ООО Йота",
        "АО Тета",
        "ИП Иванов",
    ]
    assert len(result.goods_orgs_amounts.cells) == _EXPECTED_CELLS_COUNT


def test_extract_goods_org_amounts_accepts_org_n_under_price_parent() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C1"] = "Цена единицы товара, руб."
    ws_goods.merge_cells("C1:E1")
    ws_goods["C2"] = "Орг-ция 1"
    ws_goods["D2"] = "Орг-ция 2"
    ws_goods["E2"] = "Орг-ция 3"
    ws_goods["F1"] = "НМЦ единицы товара"
    ws_goods["G1"] = "Ед. изм."
    ws_goods["H1"] = "Кол."
    ws_goods["I1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО АПС"
    ws_goods["C3"] = "10 000"
    ws_goods["D3"] = "11 000"
    ws_goods["E3"] = "9 500"
    ws_goods["F3"] = "10 166,67"
    ws_goods["G3"] = "месяц"
    ws_goods["H3"] = "3"
    ws_goods["I3"] = "30 500,01"

    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B6"] = _ANCHOR_AVG
    ws_goods["C6"] = "100 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.name for org in result.goods_orgs_amounts.organizations] == [
        "Орг-ция 1",
        "Орг-ция 2",
        "Орг-ция 3",
    ]


def test_extract_goods_org_amounts_accepts_company_names_under_price_parent() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["D1"] = "Цена единицы товара"
    ws_goods.merge_cells("D1:F1")
    ws_goods["D2"] = "ООО Йота"
    ws_goods["E2"] = "АО Тета"
    ws_goods["F2"] = "ИП Иванов"
    ws_goods["G1"] = "НМЦ единицы товара"
    ws_goods["H1"] = "Ед. изм."
    ws_goods["I1"] = "Кол."
    ws_goods["J1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО СКУД"
    ws_goods["D3"] = "8 500"
    ws_goods["E3"] = "10 200"
    ws_goods["F3"] = "8 337"
    ws_goods["G3"] = "9 012,33"
    ws_goods["H3"] = "месяц"
    ws_goods["I3"] = "12"
    ws_goods["J3"] = "108147,96"

    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B6"] = _ANCHOR_AVG
    ws_goods["C6"] = "100 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.name for org in result.goods_orgs_amounts.organizations] == [
        "ООО Йота",
        "АО Тета",
        "ИП Иванов",
    ]


def test_extract_goods_org_amounts_accepts_non_org_names_under_price_parent() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C1"] = "Цена единицы товара, руб."
    ws_goods.merge_cells("C1:D1")
    ws_goods["C2"] = "фильтрационные технологии"
    ws_goods["D2"] = "инженерные решения"
    ws_goods["E1"] = "НМЦ единицы товара"
    ws_goods["F1"] = "Ед. изм."
    ws_goods["G1"] = "Кол."
    ws_goods["H1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО СОТС"
    ws_goods["C3"] = "13 400"
    ws_goods["D3"] = "13 500"
    ws_goods["E3"] = "13 450"
    ws_goods["F3"] = "месяц"
    ws_goods["G3"] = "12"
    ws_goods["H3"] = "161400"

    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B6"] = _ANCHOR_AVG
    ws_goods["C6"] = "100 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.name for org in result.goods_orgs_amounts.organizations] == [
        "фильтрационные технологии",
        "инженерные решения",
    ]


def test_extract_goods_org_amounts_fallback_without_price_parent() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS

    ws_goods["B1"] = "Наименование товара"
    ws_goods["C1"] = "ООО Альфа"
    ws_goods["D1"] = "ООО Бета"
    ws_goods["E1"] = "НМЦ единицы товара"
    ws_goods["F1"] = "Ед. изм."
    ws_goods["G1"] = "Кол."
    ws_goods["H1"] = "НМЦ товара"

    ws_goods["B3"] = "ТО АПС"
    ws_goods["C3"] = "101"
    ws_goods["D3"] = "102"
    ws_goods["E3"] = "101,5"
    ws_goods["F3"] = "шт"
    ws_goods["G3"] = "2"
    ws_goods["H3"] = "203"

    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"
    ws_goods["B6"] = _ANCHOR_AVG
    ws_goods["C6"] = "100 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_orgs_amounts is not None
    assert [org.name for org in result.goods_orgs_amounts.organizations] == [
        "ООО Альфа",
        "ООО Бета",
    ]


def test_extract_works_org_amounts_keeps_order_and_amounts() -> None:
    wb = Workbook()

    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS
    ws_goods["B7"] = _ANCHOR_AVG
    ws_goods["C7"] = "12 331 372,00"

    ws_ws = wb.create_sheet(_SHEET_WORKS_SERVICES)
    ws_ws["A1"] = "№ п/п"
    ws_ws["B1"] = "Организации"
    ws_ws["C1"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["B2"] = "ООО Гамма Сервис"
    ws_ws["C2"] = "13 594 116,00"
    ws_ws["B3"] = "ООО ДЕЛЬТАСТРОЙ"
    ws_ws["C3"] = "12 000 000,00"
    ws_ws["B4"] = "ИП Иванов И.И."
    ws_ws["C4"] = "11 400 000,00"
    ws_ws["B5"] = _ANCHOR_AVG
    ws_ws["C5"] = "13 000 000,00"
    ws_ws["B6"] = _ANCHOR_LIMIT
    ws_ws["C6"] = "12 000 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.works_orgs_amounts is not None
    organizations = result.works_orgs_amounts.organizations
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


def test_extract_imcp_keeps_goods_only_partial_contract() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = _SHEET_GOODS
    ws_goods["B7"] = _ANCHOR_AVG
    ws_goods["C7"] = "1 234 567,89"
    ws_goods["B1"] = "Наименование товара"
    ws_goods["C2"] = "ООО Йота"
    ws_goods["F1"] = "Цена единицы товара, руб."
    ws_goods["G1"] = "Ед. изм."
    ws_goods["H1"] = "Кол."
    ws_goods["I1"] = "НМЦ товара"
    ws_goods["B3"] = "ТО АПС"
    ws_goods["C3"] = "1000"
    ws_goods["F3"] = "1200"
    ws_goods["G3"] = "шт"
    ws_goods["H3"] = "2"
    ws_goods["I3"] = "2400"
    ws_goods["B5"] = "ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ РАСЧЕТА НМЦД"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg == Decimal("1234567.89")
    assert result.goods_orgs_amounts is not None
    assert result.works_imcp_avg is None
    assert result.works_imcp_limit is None
    assert result.works_orgs_amounts is None


def test_extract_imcp_classifies_unnamed_sheet_by_goods_cells() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = "Лист 1"
    ws_goods["B1"] = "Наименование товара"
    ws_goods["C2"] = "ООО Йота"
    ws_goods["F1"] = "Цена единицы товара, руб."
    ws_goods["G1"] = "Ед. изм."
    ws_goods["H1"] = "Кол."
    ws_goods["I1"] = "НМЦ товара"
    ws_goods["B3"] = "ТО АПС"
    ws_goods["C3"] = "1000"
    ws_goods["F3"] = "1200"
    ws_goods["G3"] = "шт"
    ws_goods["H3"] = "2"
    ws_goods["I3"] = "2400"
    ws_goods["B8"] = _ANCHOR_AVG
    ws_goods["I8"] = "12 331 372,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg == Decimal("12331372.00")
    assert result.goods_orgs_amounts is not None
    assert len(result.goods_orgs_amounts.goods_items) == 1
    assert result.works_imcp_avg is None
    assert result.works_imcp_limit is None
    assert result.works_orgs_amounts is None


def test_extract_imcp_keeps_works_only_partial_contract() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = "Данные листа"

    ws_ws = wb.create_sheet("Лист с расчетом")
    ws_ws["B1"] = _ANCHOR_AVG
    ws_ws["C1"] = "2 000 000,00"
    ws_ws["B2"] = _ANCHOR_LIMIT
    ws_ws["C2"] = "1 900 000,00"
    ws_ws["A5"] = "№ п/п"
    ws_ws["B5"] = "Организации"
    ws_ws["C5"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["B6"] = "ООО Альфа"
    ws_ws["C6"] = "1 111 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg is None
    assert result.goods_orgs_amounts is None
    assert result.works_imcp_avg == Decimal("2000000.00")
    assert result.works_imcp_limit == Decimal("1900000.00")
    assert result.works_orgs_amounts is not None
    assert [org.name for org in result.works_orgs_amounts.organizations] == [
        "ООО Альфа"
    ]


def test_extract_imcp_limit_with_short_fallback_anchor() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = "Данные листа"

    ws_ws = wb.create_sheet(_SHEET_WORKS_SERVICES)
    ws_ws["B1"] = _ANCHOR_AVG
    ws_ws["C1"] = "2 000 000,00"
    ws_ws["B2"] = "Лимит финансирования"
    ws_ws["C2"] = "1 900 000,00"
    ws_ws["A5"] = "№ п/п"
    ws_ws["B5"] = "Организации"
    ws_ws["C5"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["B6"] = "ООО Альфа"
    ws_ws["C6"] = "1 111 000,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.works_imcp_avg == Decimal("2000000.00")
    assert result.works_imcp_limit == Decimal("1900000.00")


def test_extract_imcp_fallback_does_not_reuse_goods_sheet_for_works() -> None:
    wb = Workbook()
    ws_goods = wb.active
    ws_goods.title = "Расчет при закупке товаро"  # typo, fallback by anchor
    ws_goods["B3"] = _ANCHOR_AVG
    ws_goods["C3"] = "1000,00"

    ws_ws = wb.create_sheet("Расчет при закупке товаров и работ")
    ws_ws["B2"] = _ANCHOR_AVG
    ws_ws["C2"] = "2000,00"
    ws_ws["B3"] = _ANCHOR_LIMIT
    ws_ws["C3"] = "1500,00"
    ws_ws["A6"] = "№ п/п"
    ws_ws["B6"] = "Организации"
    ws_ws["C6"] = "Стоимость выполнения работ / оказания услуг, в т.ч. НДС 20 %, руб."
    ws_ws["B7"] = "ООО Бета"
    ws_ws["C7"] = "1400,00"

    extractor = OpenpyxlExcelIMCPExtractor()
    result = extractor.extract_imcp(content=_wb_to_bytes(wb))

    assert result.goods_imcp_avg == Decimal("1000.00")
    assert result.works_imcp_avg == Decimal("2000.00")
    assert result.works_imcp_limit == Decimal("1500.00")
    assert result.works_orgs_amounts is not None
    assert [org.name for org in result.works_orgs_amounts.organizations] == ["ООО Бета"]
