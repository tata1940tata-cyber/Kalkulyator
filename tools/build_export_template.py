# -*- coding: utf-8 -*-
"""Сборка Excel-шаблона выгрузки данных для калькулятора юнит-экономики."""

import csv
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "Шаблон_выгрузки_данных.xlsx"

C_DARK = "1F3864"
C_MID = "2F5597"
C_LIGHT = "D9E2F3"
C_ALT = "F2F6FC"
C_INPUT = "FFF2CC"
C_TOTAL = "DDEBF7"
C_GRAY = "808080"

THIN = Side(style="thin", color="C9D3E3")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_TITLE = Font(name="Calibri", size=16, bold=True, color=C_DARK)
F_SUB = Font(name="Calibri", size=10, italic=True, color=C_GRAY)
F_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_BODY = Font(name="Calibri", size=10)
F_BOLD = Font(name="Calibri", size=10, bold=True)
F_SMALL = Font(name="Calibri", size=9, color="404040")
F_SECTION = Font(name="Calibri", size=12, bold=True, color=C_DARK)
AL_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
AL_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

RUB = '#,##0.00" ₽"'
PCT = '0.0%'
INT = '0'


def header_cell(cell, text, comment=None):
    cell.value = text
    cell.font = F_HDR
    cell.fill = PatternFill("solid", fgColor=C_DARK)
    cell.alignment = AL_CENTER
    cell.border = BORDER
    if comment:
        cell.comment = Comment(comment, "data-model")
        cell.comment.width = 320
        cell.comment.height = 120


def sheet_title(ws, text, sub, last_col):
    ws.merge_cells(f"A1:{last_col}1")
    cell = ws["A1"]
    cell.value = text
    cell.font = F_TITLE
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A2:{last_col}2")
    cell = ws["A2"]
    cell.value = sub
    cell.font = F_SUB
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 32


# ---------------------------------------------------------------- метаданные
SKU_COLUMNS = [
    ("offer_id", "ID продавца", "Артикул продавца (offer_id). Ключ связи с API и учётной системой."),
    ("market_sku", "SKU Маркета", "Идентификатор товара на Маркете (sku)."),
    ("name", "Название", "Название карточки товара."),
    ("category_id", "ID категории", "Код категории Маркета. По нему выбирается ставка комиссии."),
    ("category_name", "Категория", "Название категории для читаемости."),
    ("brand", "Бренд", "Бренд товара."),
    ("scheme", "Схема", "FBY / FBS / DBS. Определяет набор тарифов."),
    ("weight_g", "Вес, г", "Вес единицы товара в граммах."),
    ("length_mm", "Длина, мм", "Длина упаковки в мм."),
    ("width_mm", "Ширина, мм", "Ширина упаковки в мм."),
    ("height_mm", "Высота, мм", "Высота упаковки в мм."),
    ("volume_l", "Объём, л", "Объём упаковки в литрах = Д × Ш × В / 1 000 000."),
    ("cogs_rub", "Себестоимость, ₽", "Закупка или производство единицы без возмещаемого НДС."),
    ("packaging_rub", "Упаковка, ₽", "Коробка, наполнитель, маркировка на единицу."),
    ("delivery_to_warehouse_rub", "Доставка до склада, ₽", "Транспортные расходы на единицу."),
    ("storage_days", "Срок хранения, дней", "Плановый средний срок хранения на складе."),
    ("return_rate", "Доля возвратов", "Доля заказов, которые вернулись. 0.18 = 18%."),
    ("current_price_rub", "Текущая цена, ₽", "Актуальная цена продажи."),
    ("min_price_rub", "Минимальная цена, ₽", "Ниже этой цены не продаём (MAP, договорённости)."),
    ("vat_included", "Цена с НДС", "true/false — включён ли НДС в цену."),
    ("comment", "Комментарий", "Свободное поле для заметок."),
]

PROMO_COLUMNS = [
    ("promo_id", "ID акции", "Идентификатор акции на Маркете."),
    ("promo_name", "Название акции", "Название как в кабинете."),
    ("promo_type", "Тип акции", "DISCOUNT — скидка %, FIXED_PRICE — фиксированная цена, PROMO_CODE — промокод."),
    ("offer_id", "ID продавца", "Товар-участник (связь с листом «Товары»)."),
    ("start_date", "Начало", "Дата начала в формате ГГГГ-ММ-ДД."),
    ("end_date", "Окончание", "Дата окончания в формате ГГГГ-ММ-ДД."),
    ("seller_discount_rate", "Скидка продавца", "Доля скидки за счёт продавца: 0.15 = 15%."),
    ("promo_price_rub", "Цена в акции, ₽", "Заполняется для FIXED_PRICE. Иначе считается от текущей цены."),
    ("platform_subsidy_rate", "Субсидия площадки", "Доля скидки, которую компенсирует Маркет."),
    ("participation_fee_rub", "Взнос за участие, ₽", "Фиксированная плата за участие, если есть."),
    ("drr_rate", "ДРР в акции", "Доля рекламы в выручке на период акции."),
    ("comment", "Комментарий", "Условия акции, ограничения."),
]

FIELD_DICT = [
    ("Настройки", "tax.regime", "Налоговый режим", "строка", "—", "да", "Бухгалтерия",
     "USN_INCOME — УСН «доходы», USN_PROFIT — УСН «доходы минус расходы», OSNO."),
    ("Настройки", "tax.rate", "Ставка налога", "доля", "%", "да", "Бухгалтерия", "6% для УСН «доходы»."),
    ("Настройки", "tax.vat_rate", "Ставка НДС", "доля", "%", "да", "Бухгалтерия", "0 / 5 / 7 / 20%."),
    ("Настройки", "tax.vat_in_price", "НДС в цене", "логический", "—", "да", "Бухгалтерия",
     "true, если цена в кабинете указана с НДС."),
    ("Настройки", "targets.margin_rate", "Целевая маржинальность", "доля", "%", "да", "Бизнес",
     "Цель по марже: 0.20 = 20% от выручки."),
    ("Настройки", "targets.drr_rate", "Целевой ДРР", "доля", "%", "да", "Маркетинг",
     "Плановая доля рекламы в выручке."),
    ("Настройки", "targets.min_margin_rub", "Минимальная маржа, ₽", "число", "₽", "нет", "Бизнес",
     "Порог, ниже которого позиция считается проблемной."),
    ("Настройки", "rounding.mode", "Округление цены", "строка", "—", "да", "Бизнес",
     "none / up_1 / up_10 / up_50 / psych_99."),
    ("Настройки", "defaults.storage_days", "Срок хранения по умолчанию", "число", "дни", "нет", "Планирование",
     "Используется, если у товара не заполнен storage_days."),
    ("Настройки", "defaults.return_rate", "Доля возвратов по умолчанию", "доля", "%", "нет", "Аналитика",
     "Используется, если у товара не заполнена доля возвратов."),
    ("Настройки", "defaults.other_rate", "Прочие расходы", "доля", "%", "нет", "Бизнес",
     "Брак, ПО, зарплаты и прочее — % от выручки."),
    ("Настройки", "api.campaign_id", "ID кампании", "число", "—", "да", "API", "Кампания Маркета."),
    ("Настройки", "api.business_id", "ID бизнеса", "число", "—", "нет", "API", "Для методов бизнеса."),
    ("Настройки", "api.sync_interval_minutes", "Период синхронизации", "число", "мин", "нет", "Архитектура",
     "Как часто обновлять тарифы и цены."),
    ("Товары", "offer_id", "ID продавца", "строка", "—", "да", "API / выгрузка", "Ключ товара."),
    ("Товары", "market_sku", "SKU Маркета", "число", "—", "да", "API", "Идентификатор на Маркете."),
    ("Товары", "category_id", "ID категории", "число", "—", "да", "API", "Определяет комиссию."),
    ("Товары", "scheme", "Схема работы", "строка", "—", "да", "API / кабинет", "FBY / FBS / DBS."),
    ("Товары", "volume_l", "Объём упаковки", "число", "л", "да", "API / карточка", "Для тарифа логистики и хранения."),
    ("Товары", "cogs_rub", "Себестоимость", "число", "₽", "да", "Учётная система", "Основа расчёта маржи."),
    ("Товары", "packaging_rub", "Упаковка", "число", "₽", "да", "Учётная система", "Стоимость на единицу."),
    ("Товары", "delivery_to_warehouse_rub", "Доставка до склада", "число", "₽", "нет", "Логистика", "На единицу."),
    ("Товары", "storage_days", "Срок хранения", "число", "дни", "нет", "Планирование", "Средний срок до продажи."),
    ("Товары", "return_rate", "Доля возвратов", "доля", "%", "нет", "Аналитика", "0.18 = 18%."),
    ("Товары", "current_price_rub", "Текущая цена", "число", "₽", "да", "API / кабинет", "Цена на витрине."),
    ("Товары", "min_price_rub", "Минимальная цена", "число", "₽", "нет", "Бизнес", "Ограничение снизу."),
    ("Товары", "vat_included", "Цена с НДС", "логический", "—", "нет", "Бухгалтерия", "Признак НДС в цене."),
    ("Тарифы", "commission_by_category", "Комиссия по категориям", "словарь", "%", "да", "Кабинет → Тарифы",
     "Ставка комиссии за продажу по каждой категории."),
    ("Тарифы", "logistics_tiers", "Логистика по объёму", "таблица", "₽", "да", "Кабинет → Тарифы",
     "Ступени объёма и цена за единицу."),
    ("Тарифы", "last_mile", "Последняя миля", "объект", "% / ₽", "да", "Кабинет → Тарифы",
     "Ставка от цены + минимум и максимум."),
    ("Тарифы", "acquiring", "Эквайринг", "объект", "%", "да", "Договор", "0, если включён в комиссию."),
    ("Тарифы", "storage", "Хранение", "объект", "₽/л/сутки", "да", "Кабинет → Тарифы", "Тариф за литр в сутки."),
    ("Тарифы", "returns", "Возвраты", "объект", "₽", "да", "Кабинет → Тарифы",
     "Стоимость возврата и возвращается ли комиссия."),
    ("Тарифы", "utilization", "Утилизация", "объект", "₽", "нет", "Кабинет → Тарифы", "Плата за утилизацию единицы."),
    ("Тарифы", "fines", "Штрафы", "объект", "%", "нет", "Кабинет → Тарифы", "Плановая доля штрафов от выручки."),
    ("Акции", "promo_id", "ID акции", "строка", "—", "да", "API / кабинет", "Ключ акции."),
    ("Акции", "promo_type", "Тип акции", "строка", "—", "да", "API / кабинет", "DISCOUNT / FIXED_PRICE / PROMO_CODE."),
    ("Акции", "offer_id", "ID продавца", "строка", "—", "да", "API", "Товар-участник."),
    ("Акции", "seller_discount_rate", "Скидка продавца", "доля", "%", "нет", "Кабинет", "За счёт продавца."),
    ("Акции", "promo_price_rub", "Цена в акции", "число", "₽", "нет", "Кабинет", "Для фиксированной цены."),
    ("Акции", "drr_rate", "ДРР в акции", "доля", "%", "нет", "Маркетинг", "Повышенный ДРР на период акции."),
    ("Акции", "participation_fee_rub", "Взнос за участие", "число", "₽", "нет", "Кабинет", "Фиксированная плата."),
]

RESULT_CAPTIONS = {
    "offer_id": "ID продавца", "name": "Товар", "scenario": "Сценарий",
    "price_rub": "Цена, ₽", "revenue_rub": "Выручка, ₽",
    "commission_rub": "Комиссия, ₽", "logistics_rub": "Логистика, ₽",
    "last_mile_rub": "Последняя миля, ₽", "acquiring_rub": "Эквайринг, ₽",
    "storage_rub": "Хранение, ₽", "returns_rub": "Возвраты, ₽",
    "utilization_rub": "Утилизация, ₽", "fines_rub": "Штрафы, ₽",
    "delivery_rub": "Доставка до склада, ₽", "drr_rub": "Реклама / ДРР, ₽",
    "tax_rub": "Налоги, ₽", "other_rub": "Прочее, ₽",
    "cogs_rub": "Себестоимость, ₽", "packaging_rub": "Упаковка, ₽",
    "total_costs_rub": "Итого расходов, ₽", "margin_rub": "Маржа, ₽",
    "margin_rate": "Маржинальность", "roi": "ROI к себестоимости", "flags": "Проверки",
}

wb = Workbook()

# ---------------------------------------------------------------- инструкция
ws = wb.active
ws.title = "Инструкция"
ws.sheet_properties.tabColor = C_MID
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 110
ws.sheet_view.showGridLines = False
ws.merge_cells("A1:B1")
ws["A1"] = "Шаблон выгрузки данных для калькулятора юнит-экономики"
ws["A1"].font = F_TITLE
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 32

lines = [
    ("Зачем это файл", F_SECTION),
    ("Шаблон задаёт полный набор данных, который нужен калькулятору: тарифы площадки, характеристики товаров, "
     "цены, акции и настройки расчёта. Заполненный файл можно загружать в калькулятор вместо API.", F_BODY),
    ("", F_BODY),
    ("Как заполнять", F_SECTION),
    ("1. Лист «Настройки» — налоговый режим, целевая маржинальность, целевой ДРР, правило округления.", F_BODY),
    ("2. Лист «Товары» — по строке на SKU: категория, схема, объём, себестоимость, текущая и минимальная цена.", F_BODY),
    ("3. Листы «Комиссии» и «Логистика» — тарифы площадки по вашим категориям и ступеням объёма.", F_BODY),
    ("4. Лист «Прочие тарифы» — последняя миля, эквайринг, хранение, возвраты, утилизация, штрафы.", F_BODY),
    ("5. Лист «Акции» — акции, в которые планируем входить, и цены входа.", F_BODY),
    ("6. Лист «Результаты (эталон)» — пример расчёта, с которым сверяется боевая реализация.", F_BODY),
    ("", F_BODY),
    ("Правила", F_SECTION),
    ("•  Строка 1 на листах данных — технические имена полей. Не переименовывать: по ним работает импорт. "
     "В примечании к каждой ячейке — расшифровка.", F_BODY),
    ("•  Проценты записываются долями: 15% = 0.15, а не 15.", F_BODY),
    ("•  Даты — в формате ГГГГ-ММ-ДД. Деньги — в рублях, разделитель — точка.", F_BODY),
    ("•  Обязательные поля нельзя оставлять пустыми: калькулятор остановит расчёт и покажет ошибку.", F_BODY),
    ("•  Тарифы привязаны к дате: при изменении условий создавайте новую версию, а не правьте старую.", F_BODY),
    ("", F_BODY),
    ("Что уточнить перед запуском", F_SECTION),
    ("•  Схема работы по каждому SKU: FBY, FBS или DBS.", F_BODY),
    ("•  Налоговый режим и ставка НДС.", F_BODY),
    ("•  Целевая маржинальность и целевой ДРР — от них зависит расчёт новой цены.", F_BODY),
    ("•  Правило округления цены: до 10 ₽, до 50 ₽ или «психологические» 99.", F_BODY),
    ("•  Кто отвечает за актуализацию тарифов и как часто они меняются.", F_BODY),
]
row = 3
for text, font in lines:
    ws.merge_cells(f"A{row}:B{row}")
    cell = ws[f"A{row}"]
    cell.value = text
    cell.font = font
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[row].height = 18 if len(text) < 100 else 34
    row += 1

# ---------------------------------------------------------------- словарь полей
ws = wb.create_sheet("Словарь полей")
ws.sheet_properties.tabColor = "7F7F7F"
sheet_title(ws, "Словарь полей", "Полный перечень данных, которые нужны калькулятору, с источниками.", "H")
widths = {"A": 14, "B": 26, "C": 30, "D": 12, "E": 12, "F": 12, "G": 22, "H": 70}
for col, width in widths.items():
    ws.column_dimensions[col].width = width
head = ["Лист", "Поле", "Название", "Тип", "Ед.", "Обяз.", "Источник", "Описание"]
for i, text in enumerate(head, start=1):
    header_cell(ws.cell(row=4, column=i), text)
for i, item in enumerate(FIELD_DICT):
    row = 5 + i
    for j, value in enumerate(item, start=1):
        cell = ws.cell(row=row, column=j, value=value)
        cell.font = F_BODY
        cell.alignment = AL_WRAP
        cell.border = BORDER
        if i % 2:
            cell.fill = PatternFill("solid", fgColor=C_ALT)
    ws.row_dimensions[row].height = 30
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:H{4 + len(FIELD_DICT)}"

# ---------------------------------------------------------------- настройки
ws = wb.create_sheet("Настройки")
ws.sheet_properties.tabColor = "C00000"
sheet_title(ws, "Настройки расчёта", "Значения в колонке B — входные, их можно менять.", "D")
settings = json.load(open(DATA / "settings.json", encoding="utf-8"))

for col, width in {"A": 34, "B": 20, "C": 12, "D": 62}.items():
    ws.column_dimensions[col].width = width
for i, text in enumerate(["Параметр", "Значение", "Ед.", "Комментарий"], start=1):
    header_cell(ws.cell(row=4, column=i), text)

settings_rows = [
    ("Налоговый режим", settings["tax"]["regime"], "—", "USN_INCOME / USN_PROFIT / OSNO"),
    ("Ставка налога", settings["tax"]["rate"], "%", "6% для УСН «доходы»"),
    ("Ставка НДС", settings["tax"]["vat_rate"], "%", "0, если не платите НДС"),
    ("НДС включён в цену", str(settings["tax"]["vat_in_price"]).lower(), "—", "true/false"),
    ("Целевая маржинальность", settings["targets"]["margin_rate"], "%", "Цель по марже от выручки"),
    ("Целевой ДРР", settings["targets"]["drr_rate"], "%", "Плановая доля рекламы"),
    ("Минимальная маржа", settings["targets"]["min_margin_rub"], "₽", "Порог проблемной позиции"),
    ("Округление цены", settings["rounding"]["mode"], "—", "none / up_1 / up_10 / up_50 / psych_99"),
    ("Шаг округления", settings["rounding"]["step"], "₽", "Используется для режима up_N"),
    ("Срок хранения по умолчанию", settings["defaults"]["storage_days"], "дни", "Если у товара не задан"),
    ("Доля возвратов по умолчанию", settings["defaults"]["return_rate"], "%", "Если у товара не задана"),
    ("Прочие расходы", settings["defaults"]["other_rate"], "%", "Брак, ПО, зарплаты"),
    ("ID кампании", settings["api"]["campaign_id"], "—", "Для синхронизации по API"),
    ("ID бизнеса", settings["api"]["business_id"], "—", "Для методов бизнеса"),
    ("Период синхронизации", settings["api"]["sync_interval_minutes"], "мин", "Частота обновления"),
]
for i, (label, value, unit, comment) in enumerate(settings_rows):
    row = 5 + i
    ws.cell(row=row, column=1, value=label).font = F_BODY
    cell = ws.cell(row=row, column=2, value=value)
    cell.font = F_BOLD
    cell.alignment = AL_CENTER
    cell.fill = PatternFill("solid", fgColor=C_INPUT)
    ws.cell(row=row, column=3, value=unit).alignment = AL_CENTER
    ws.cell(row=row, column=4, value=comment).font = F_SMALL
    for col in range(1, 5):
        ws.cell(row=row, column=col).border = BORDER
    ws.row_dimensions[row].height = 20
ws.freeze_panes = "A5"


def dump_table(ws, columns, rows, start_row=4, formats=None, comments=True):
    formats = formats or {}
    for i, (key, caption, comment) in enumerate(columns, start=1):
        cell = ws.cell(row=start_row, column=i)
        header_cell(cell, key, comment if comments else None)
        ws.column_dimensions[get_column_letter(i)].width = max(12, min(30, len(caption) + 6))
    for i, row_data in enumerate(rows):
        row = start_row + 1 + i
        for j, (key, _caption, _comment) in enumerate(columns, start=1):
            value = row_data.get(key, "")
            if key in formats and value not in ("", None):
                try:
                    value = float(value)
                except ValueError:
                    pass
            cell = ws.cell(row=row, column=j, value=value)
            cell.font = F_BODY
            cell.alignment = AL_WRAP
            cell.border = BORDER
            if key in formats:
                cell.number_format = formats[key]
            if i % 2:
                cell.fill = PatternFill("solid", fgColor=C_ALT)
        ws.row_dimensions[row].height = 26
    ws.freeze_panes = f"A{start_row + 1}"
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(columns))}{start_row + len(rows)}"


def read_csv(name):
    with open(DATA / name, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


# ---------------------------------------------------------------- товары
ws = wb.create_sheet("Товары")
ws.sheet_properties.tabColor = "548235"
sheet_title(ws, "Товары (SKU)", "По строке на товар. Жёлтым выделены обязательные для заполнения поля.", "U")
dump_table(
    ws, SKU_COLUMNS, read_csv("skus.csv"),
    formats={"volume_l": '0.00', "cogs_rub": RUB, "packaging_rub": RUB,
             "delivery_to_warehouse_rub": RUB, "return_rate": PCT,
             "current_price_rub": RUB, "min_price_rub": RUB},
)

# ---------------------------------------------------------------- комиссии
ws = wb.create_sheet("Комиссии")
ws.sheet_properties.tabColor = "BF8F00"
sheet_title(ws, "Комиссии за продажу", "Ставки по категориям и схемам работы. Проценты — долями.", "E")
tariffs = json.load(open(DATA / "tariffs.json", encoding="utf-8"))
cat_names = {row["category_id"]: row["category_name"] for row in read_csv("skus.csv")}
comm_rows = []
for scheme, cfg in tariffs["schemes"].items():
    for cat_id, rate in cfg["commission_by_category"].items():
        comm_rows.append({"scheme": scheme, "category_id": cat_id,
                          "category_name": cat_names.get(cat_id, ""),
                          "commission_rate": rate, "comment": ""})
    comm_rows.append({"scheme": scheme, "category_id": "default",
                      "category_name": "Прочие категории",
                      "commission_rate": cfg["default_commission"],
                      "comment": "Ставка по умолчанию, если категории нет в списке"})

comm_cols = [
    ("scheme", "Схема", "FBY / FBS / DBS."),
    ("category_id", "ID категории", "Код категории Маркета или default."),
    ("category_name", "Категория", "Название для читаемости."),
    ("commission_rate", "Комиссия", "Доля от цены: 0.15 = 15%."),
    ("comment", "Комментарий", "Источник и дата проверки."),
]
for i, (key, caption, comment) in enumerate(comm_cols, start=1):
    header_cell(ws.cell(row=4, column=i), key, comment)
for i, row_data in enumerate(comm_rows):
    row = 5 + i
    for j, (key, _c, _m) in enumerate(comm_cols, start=1):
        cell = ws.cell(row=row, column=j, value=row_data[key])
        cell.font = F_BODY
        cell.border = BORDER
        cell.alignment = AL_WRAP
        if key == "commission_rate":
            cell.number_format = PCT
            cell.alignment = AL_CENTER
        if i % 2:
            cell.fill = PatternFill("solid", fgColor=C_ALT)
for col, width in {"A": 10, "B": 14, "C": 26, "D": 14, "E": 44}.items():
    ws.column_dimensions[col].width = width
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:E{4 + len(comm_rows)}"

# ---------------------------------------------------------------- логистика
ws = wb.create_sheet("Логистика")
ws.sheet_properties.tabColor = "7030A0"
sheet_title(ws, "Логистика по объёму", "Ступени объёма упаковки и стоимость за единицу.", "E")
log_rows = []
for scheme, cfg in tariffs["schemes"].items():
    for tier in cfg["logistics_tiers"]:
        log_rows.append({"scheme": scheme,
                         "volume_from_l": tier["volume_from_l"],
                         "volume_to_l": tier["volume_to_l"],
                         "price_rub": tier["price_rub"],
                         "comment": "свыше 50 л — фиксированная ставка" if tier["volume_to_l"] is None else ""})
log_cols = [
    ("scheme", "Схема", "FBY / FBS / DBS."),
    ("volume_from_l", "Объём от, л", "Нижняя граница ступени."),
    ("volume_to_l", "Объём до, л", "Верхняя граница включительно; пусто — без ограничения."),
    ("price_rub", "Стоимость, ₽", "За единицу товара."),
    ("comment", "Комментарий", "Условия применения."),
]
for i, (key, caption, comment) in enumerate(log_cols, start=1):
    header_cell(ws.cell(row=4, column=i), key, comment)
for i, row_data in enumerate(log_rows):
    row = 5 + i
    for j, (key, _c, _m) in enumerate(log_cols, start=1):
        value = row_data[key]
        cell = ws.cell(row=row, column=j, value=value)
        cell.font = F_BODY
        cell.border = BORDER
        cell.alignment = AL_CENTER if key != "comment" else AL_WRAP
        if key == "price_rub":
            cell.number_format = RUB
        if i % 2:
            cell.fill = PatternFill("solid", fgColor=C_ALT)
for col, width in {"A": 10, "B": 14, "C": 14, "D": 15, "E": 40}.items():
    ws.column_dimensions[col].width = width
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:E{4 + len(log_rows)}"

# ---------------------------------------------------------------- прочие тарифы
ws = wb.create_sheet("Прочие тарифы")
ws.sheet_properties.tabColor = "2E75B6"
sheet_title(ws, "Прочие тарифы", "Последняя миля, эквайринг, хранение, возвраты, утилизация, штрафы.", "E")
other_rows = []
for scheme, cfg in tariffs["schemes"].items():
    other_rows += [
        {"scheme": scheme, "parameter": "last_mile.rate", "value": cfg["last_mile"]["rate"], "unit": "% от цены", "comment": "Доля от цены товара"},
        {"scheme": scheme, "parameter": "last_mile.min_rub", "value": cfg["last_mile"]["min_rub"], "unit": "₽", "comment": "Минимум за заказ"},
        {"scheme": scheme, "parameter": "last_mile.max_rub", "value": cfg["last_mile"]["max_rub"], "unit": "₽", "comment": "Максимум за заказ"},
        {"scheme": scheme, "parameter": "acquiring.rate", "value": cfg["acquiring"]["rate"], "unit": "% от суммы", "comment": "0, если включён в комиссию"},
        {"scheme": scheme, "parameter": "storage.rate_per_liter_day", "value": cfg["storage"]["rate_per_liter_day"], "unit": "₽/л/сутки", "comment": "Тариф хранения"},
        {"scheme": scheme, "parameter": "storage.free_days", "value": cfg["storage"]["free_days"], "unit": "дни", "comment": "Льготный период"},
        {"scheme": scheme, "parameter": "returns.fee_per_return_rub", "value": cfg["returns"]["fee_per_return_rub"], "unit": "₽", "comment": "Стоимость одного возврата"},
        {"scheme": scheme, "parameter": "utilization.fee_per_unit_rub", "value": cfg["utilization"]["fee_per_unit_rub"], "unit": "₽", "comment": "За утилизированную единицу"},
        {"scheme": scheme, "parameter": "fines.rate", "value": cfg["fines"]["rate"], "unit": "% от выручки", "comment": "Плановые штрафы"},
        {"scheme": scheme, "parameter": "delivery_to_warehouse_rub", "value": cfg["delivery_to_warehouse_rub"], "unit": "₽", "comment": "Если не задано у товара"},
    ]
other_cols = [
    ("scheme", "Схема", "FBY / FBS / DBS."),
    ("parameter", "Параметр", "Технический путь параметра в tariffs.json."),
    ("value", "Значение", "Ставка или сумма."),
    ("unit", "Ед.", "Единица измерения."),
    ("comment", "Комментарий", "Как применяется."),
]
for i, (key, caption, comment) in enumerate(other_cols, start=1):
    header_cell(ws.cell(row=4, column=i), key, comment)
for i, row_data in enumerate(other_rows):
    row = 5 + i
    for j, (key, _c, _m) in enumerate(other_cols, start=1):
        cell = ws.cell(row=row, column=j, value=row_data[key])
        cell.font = F_BODY
        cell.border = BORDER
        cell.alignment = AL_WRAP
        if key == "parameter":
            cell.font = Font(name="Consolas", size=9)
        if i % 2:
            cell.fill = PatternFill("solid", fgColor=C_ALT)
for col, width in {"A": 10, "B": 34, "C": 14, "D": 16, "E": 40}.items():
    ws.column_dimensions[col].width = width
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:E{4 + len(other_rows)}"

# ---------------------------------------------------------------- акции
ws = wb.create_sheet("Акции")
ws.sheet_properties.tabColor = "C00000"
sheet_title(ws, "Акции и цены входа", "Акции, в которые планируем входить, и условия участия.", "L")
dump_table(
    ws, PROMO_COLUMNS, read_csv("promos.csv"),
    formats={"seller_discount_rate": PCT, "promo_price_rub": RUB,
             "platform_subsidy_rate": PCT, "participation_fee_rub": RUB,
             "drr_rate": PCT},
)

# ---------------------------------------------------------------- результаты
ws = wb.create_sheet("Результаты (эталон)")
ws.sheet_properties.tabColor = "548235"
sheet_title(
    ws, "Эталонные расчёты",
    "Сформировано reference_calc.py. Используйте для сверки боевой реализации: цифры должны совпадать.",
    "X",
)
with open(DATA / "expected_results.csv", encoding="utf-8-sig", newline="") as fh:
    reader = csv.DictReader(fh)
    result_rows = list(reader)
    result_cols = reader.fieldnames

for i, key in enumerate(result_cols, start=1):
    cell = ws.cell(row=4, column=i)
    header_cell(cell, RESULT_CAPTIONS.get(key, key), f"Техническое имя: {key}")
    ws.column_dimensions[get_column_letter(i)].width = 16 if key not in ("name", "flags") else 34

for i, row_data in enumerate(result_rows):
    row = 5 + i
    for j, key in enumerate(result_cols, start=1):
        value = row_data.get(key, "")
        if key.endswith("_rub") or key in ("margin_rate", "roi"):
            try:
                value = float(value)
            except ValueError:
                pass
        cell = ws.cell(row=row, column=j, value=value)
        cell.font = F_BODY
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if key.endswith("_rub"):
            cell.number_format = RUB
        elif key in ("margin_rate", "roi"):
            cell.number_format = PCT
        if row_data.get("margin_rub") not in ("", None):
            try:
                if float(row_data["margin_rub"]) < 0:
                    cell.fill = PatternFill("solid", fgColor="FFC7CE")
            except ValueError:
                pass
    ws.row_dimensions[row].height = 22
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:{get_column_letter(len(result_cols))}{4 + len(result_rows)}"

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved:", OUT.name)
