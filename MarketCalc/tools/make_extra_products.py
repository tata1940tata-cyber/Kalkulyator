# -*- coding: utf-8 -*-
"""Генерация файла «доп.товары для загрузки.xlsx» для калькулятора МАРКЕТ."""

import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent.parent
APP = ROOT / "MarketCalc" / "app"
OUT = ROOT / "доп.товары для загрузки.xlsx"
OUT_COPY = ROOT / "MarketCalc" / "release" / "доп.товары для загрузки.xlsx"

HEADERS = [
    "offer_id", "market_sku", "name", "category_id", "category_name", "scheme",
    "weight_g", "length_mm", "width_mm", "height_mm", "volume_l", "cogs_rub",
    "packaging_rub", "delivery_to_warehouse_rub", "storage_days", "return_rate",
    "current_price_rub", "min_price_rub", "commission_rate", "logistics_rub",
    "last_mile_rate", "acquiring_rate", "drr_rate", "tax_rate", "other_rate",
]

CATEGORIES = {
    "101": "Бытовая техника",
    "102": "Обувь",
    "103": "Красота и косметика",
    "104": "Товары для дома",
    "105": "Детские товары",
}

# артикул, sku, название, категория, бренд, вес г, Д×Ш×В мм, цена, себестоимость,
# упаковка, доставка, дней хранения, возвраты, мин. цена
PRODUCTS = [
    ("DOP-001", 2001000001, "Чайник электрический 1,7 л", "101", "Aurora", 1100, (220, 220, 280), 1990, 830, 25, 25, 45, 0.03, 1690),
    ("DOP-002", 2001000002, "Утюг паровой 2400 Вт", "101", "Aurora", 1400, (320, 170, 210), 3490, 1550, 35, 30, 45, 0.04, 2930),
    ("DOP-003", 2001000003, "Пылесос вертикальный беспроводной", "101", "Voltix", 2900, (300, 250, 1150), 12990, 6100, 70, 50, 60, 0.05, 10990),
    ("DOP-004", 2001000004, "Мультиварка 5 л, 12 программ", "101", "Aurora", 4200, (350, 300, 320), 7490, 3350, 60, 40, 50, 0.04, 6290),

    ("DOP-005", 2001000005, "Кроссовки женские белые, 38", "102", "Stride", 800, (330, 230, 120), 4290, 1750, 25, 25, 45, 0.18, 3590),
    ("DOP-006", 2001000006, "Ботинки мужские зимние, 43", "102", "Stride", 1300, (360, 280, 150), 6990, 3050, 35, 30, 60, 0.15, 5890),
    ("DOP-007", 2001000007, "Сандалии детские, 30", "102", "Stride", 400, (250, 180, 100), 1890, 690, 15, 20, 40, 0.12, 1590),
    ("DOP-008", 2001000008, "Тапочки домашние унисекс", "102", "Stride", 300, (300, 200, 100), 990, 330, 12, 20, 30, 0.08, 790),

    ("DOP-009", 2001000009, "Шампунь профессиональный 500 мл", "103", "GlowLab", 560, (80, 80, 230), 1290, 410, 15, 20, 60, 0.04, 1090),
    ("DOP-010", 2001000010, "Крем для лица с гиалуроном 50 мл", "103", "GlowLab", 120, (60, 60, 90), 1690, 470, 10, 20, 45, 0.05, 1390),
    ("DOP-011", 2001000011, "Парфюмерная вода 50 мл", "103", "GlowLab", 260, (90, 60, 160), 4990, 2050, 20, 25, 90, 0.06, 4290),
    ("DOP-012", 2001000012, "Набор кистей для макияжа, 12 шт", "103", "GlowLab", 350, (260, 180, 60), 2390, 740, 18, 20, 60, 0.03, 1990),

    ("DOP-013", 2001000013, "Набор кастрюль 6 предметов", "104", "HomeLine", 5600, (450, 350, 300), 5990, 2650, 70, 45, 45, 0.04, 4990),
    ("DOP-014", 2001000014, "Органайзер для кухни навесной", "104", "HomeLine", 700, (400, 300, 150), 1490, 430, 20, 25, 45, 0.03, 1190),
    ("DOP-015", 2001000015, "Постельное бельё 2-спальное, сатин", "104", "HomeLine", 1400, (350, 250, 120), 3290, 1280, 30, 25, 40, 0.05, 2690),
    ("DOP-016", 2001000016, "Коврик придверный 60×90 см", "104", "HomeLine", 1600, (650, 150, 150), 1190, 360, 20, 30, 50, 0.06, 990),

    ("DOP-017", 2001000017, "Конструктор, 500 деталей", "105", "KidsJoy", 900, (400, 300, 250), 2490, 780, 35, 30, 60, 0.05, 1990),
    ("DOP-018", 2001000018, "Кукла с аксессуарами, 30 см", "105", "KidsJoy", 420, (300, 150, 350), 1790, 540, 25, 25, 45, 0.06, 1490),
    ("DOP-019", 2001000019, "Развивающий коврик для малышей", "105", "KidsJoy", 1100, (700, 500, 100), 2890, 930, 30, 30, 50, 0.04, 2390),
    ("DOP-020", 2001000020, "Самокат детский трёхколёсный", "105", "KidsJoy", 2200, (700, 200, 250), 3990, 1480, 45, 40, 60, 0.07, 3290),
]

PROMOS = [
    ("DOP-P1", "Скидка дня", "DISCOUNT", "DOP-001", "2026-11-10", "2026-11-17", 0.08, None, 0.14, "Скидка за счёт продавца"),
    ("DOP-P2", "Чёрная пятница", "DISCOUNT", "DOP-005", "2026-11-20", "2026-11-30", 0.12, None, 0.15, "Обувь, повышенный ДРР"),
    ("DOP-P3", "Чёрная пятница", "FIXED_PRICE", "DOP-010", "2026-11-20", "2026-11-30", None, 1390, 0.12, "Фиксированная цена"),
    ("DOP-P4", "Распродажа склада", "DISCOUNT", "DOP-015", "2026-11-05", "2026-11-25", 0.25, None, 0.10, "Проверка низкой маржи"),
    ("DOP-P5", "Новинки", "DISCOUNT", "DOP-017", "2026-11-01", "2026-11-30", 0.06, None, 0.16, "Вывод новинки"),
    ("DOP-P6", "Красота недели", "DISCOUNT", "DOP-012", "2026-11-08", "2026-11-15", 0.10, None, 0.12, "Набор кистей"),
]

DARK = "111827"
THIN = Side(style="thin", color="D0D7E5")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header(ws, row, titles, widths):
    for index, title in enumerate(titles, start=1):
        cell = ws.cell(row=row, column=index, value=title)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=DARK)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.row_dimensions[row].height = 30


def intro(ws, title, note, last_col):
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=DARK)
    ws.row_dimensions[1].height = 26
    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = note
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="808080")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Товары"
    intro(ws, "Доп. товары для загрузки — 20 SKU",
          "Готовый набор для кнопки «Загрузить данные». Категории 101–105 соответствуют демо-тарифам. "
          "Можно менять названия, цены и характеристики — структуру колонок не трогайте.",
          "Y")
    header(ws, 4, HEADERS,
           [12, 14, 38, 10, 20, 8, 9, 9, 9, 9, 9, 12, 11, 14, 10, 10, 12, 11, 11, 11, 11, 11, 9, 9, 9])

    for row_index, product in enumerate(PRODUCTS, start=5):
        (offer_id, sku, name, cat, brand, weight, dims, price, cogs,
         pack, delivery, days, returns, min_price) = product
        volume = round(dims[0] * dims[1] * dims[2] / 1_000_000, 2)
        values = [
            offer_id, sku, name, cat, CATEGORIES[cat], "FBY",
            weight, dims[0], dims[1], dims[2], volume, cogs,
            pack, delivery, days, returns, price, min_price,
            None, None, None, None, None, None, None,
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=col, value=value)
            cell.font = Font(name="Calibri", size=10)
            cell.border = BORDER
            if col == 16:
                cell.number_format = '0.0%'
            elif col in (12, 13, 14, 17, 18):
                cell.number_format = '#,##0 ₽'
            elif col == 11:
                cell.number_format = '0.00'
        ws.row_dimensions[row_index].height = 18

    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:Y{4 + len(PRODUCTS)}"

    # --- акции
    ws2 = wb.create_sheet("Акции")
    intro(ws2, "Акции для новых SKU",
          "Необязательный лист: если он есть, калькулятор посчитает маржу по ценам входа. "
          "Ненужные строки можно удалить.", "J")
    header(ws2, 4, ["promo_id", "promo_name", "promo_type", "offer_id", "start_date",
                    "end_date", "seller_discount_rate", "promo_price_rub", "drr_rate", "comment"],
           [12, 24, 14, 12, 12, 12, 14, 14, 9, 38])
    for row_index, promo in enumerate(PROMOS, start=5):
        for col, value in enumerate(promo, start=1):
            cell = ws2.cell(row=row_index, column=col, value=value)
            cell.font = Font(name="Calibri", size=10)
            cell.border = BORDER
            if col in (7, 9):
                cell.number_format = '0.0%'
            if col == 8:
                cell.number_format = '#,##0 ₽'
        ws2.row_dimensions[row_index].height = 18
    ws2.freeze_panes = "A5"

    # --- инструкция
    guide = wb.create_sheet("Инструкция")
    guide.column_dimensions["A"].width = 4
    guide.column_dimensions["B"].width = 105
    guide.sheet_view.showGridLines = False
    guide.merge_cells("A1:B1")
    guide["A1"] = "Как загрузить этот файл"
    guide["A1"].font = Font(name="Calibri", size=16, bold=True, color=DARK)
    guide.row_dimensions[1].height = 30
    lines = [
        ("1. Откройте приложение МАРКЕТ и нажмите «Загрузить данные».", False),
        ("2. Выберите файл «доп.товары для загрузки.xlsx».", False),
        ("3. В приложении появятся 20 товаров; тарифы останутся прежними.", False),
        ("4. Акции с листа «Акции» подтянутся автоматически.", False),
        ("", False),
        ("Состав набора", True),
        ("• Бытовая техника — 4 SKU (категория 101).", False),
        ("• Обувь — 4 SKU (категория 102).", False),
        ("• Красота и косметика — 4 SKU (категория 103).", False),
        ("• Товары для дома — 4 SKU (категория 104).", False),
        ("• Детские товары — 4 SKU (категория 105).", False),
        ("", False),
        ("Что можно менять", True),
        ("• Названия, цены, себестоимость, габариты и долю возвратов.", False),
        ("• Строку заголовков (4-я строка) и названия колонок менять не нужно.", False),
        ("• Проценты записываются долями: 0.15 = 15%.", False),
        ("", False),
        (f"Файл сформирован {datetime.now().strftime('%d.%m.%Y')}. Цены демонстрационные.", False),
    ]
    row = 3
    for text, bold in lines:
        guide.merge_cells(f"A{row}:B{row}")
        cell = guide[f"A{row}"]
        cell.value = text
        cell.font = Font(name="Calibri", size=11, bold=bold, color="7A5B00" if bold else "333333")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        guide.row_dimensions[row].height = 20 if len(text) < 95 else 30
        row += 1

    wb.save(OUT)
    OUT_COPY.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_COPY)

    # --- проверка расчётом
    bundle = json.loads((APP / "data" / "default_bundle.json").read_text(encoding="utf-8"))
    scheme = bundle["tariffs"]["schemes"]["FBY"]
    settings = bundle["settings"]

    def tier(volume):
        for t in scheme["logistics_tiers"]:
            if t["volume_to_l"] is None or volume <= t["volume_to_l"]:
                return t["price_rub"]
        return 220

    print(f"{'SKU':10} {'цена':>8} {'расход':>9} {'маржа':>8} {'%':>7}")
    bad = 0
    for offer_id, sku, name, cat, brand, weight, dims, price, cogs, pack, delivery, days, returns, min_price in PRODUCTS:
        volume = round(dims[0] * dims[1] * dims[2] / 1_000_000, 2)
        commission = price * scheme["commission_by_category"][cat]
        logistics = tier(volume)
        lm = min(500, max(25, price * 0.055))
        storage = 0.15 * volume * days
        ret = returns * 60
        util = 30 * returns
        fines = price * 0.007
        drr = price * 0.10
        tax = price * 0.06
        other = price * 0.01
        total = (commission + logistics + lm + storage + ret + util + fines
                 + delivery + drr + tax + cogs + pack + other)
        margin = price - total
        rate = margin / price * 100
        print(f"{offer_id:10} {price:8.0f} {total:9.2f} {margin:8.2f} {rate:6.1f}%  {name}")
        if margin < 0:
            bad += 1
    print("убыточных:", bad)
    print("saved:", OUT)
    print("saved:", OUT_COPY)


if __name__ == "__main__":
    main()
