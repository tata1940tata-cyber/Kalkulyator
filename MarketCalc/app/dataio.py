# -*- coding: utf-8 -*-
"""Данные приложения: хранение, шаблон прайса, экспорт расчёта."""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

APP_NAME = "MarketCalc"
APP_TITLE = "МАРКЕТ — калькулятор юнит-экономики"

if getattr(sys, "frozen", False):
    ROOT = Path(getattr(sys, "_MEIPASS"))
else:
    ROOT = Path(__file__).resolve().parent

WEB_DIR = ROOT / "web"
BUNDLED_DATA = ROOT / "data" / "default_bundle.json"
DATA_FILE_NAME = "data.json"

DARK = "111827"
ACCENT = "FFD43B"
THIN = Side(style="thin", color="D0D7E5")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor=DARK)
HDR_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BODY = Font(name="Calibri", size=10)
MONEY = '#,##0.00" ₽"'
PCT = '0.0%'


def user_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    directory = Path(base) / APP_NAME
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def stored_path() -> Path:
    return user_data_dir() / DATA_FILE_NAME


def default_bundle() -> dict:
    with open(BUNDLED_DATA, encoding="utf-8") as fh:
        return json.load(fh)


def load_bundle() -> dict:
    path = stored_path()
    if path.exists():
        try:
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
            if data.get("skus"):
                return data
        except (OSError, ValueError):
            pass
    return default_bundle()


def save_bundle(bundle: dict) -> None:
    bundle.setdefault("meta", {})
    bundle["meta"]["saved_at"] = datetime.now().isoformat(timespec="seconds")
    with open(stored_path(), "w", encoding="utf-8") as fh:
        json.dump(bundle, fh, ensure_ascii=False, indent=2)


def reset_bundle() -> dict:
    path = stored_path()
    if path.exists():
        path.unlink()
    return default_bundle()


# ------------------------------------------------------------------ шаблон
SKU_HEADERS = [
    "offer_id", "market_sku", "name", "category_id", "category_name", "scheme",
    "weight_g", "length_mm", "width_mm", "height_mm", "volume_l", "cogs_rub",
    "packaging_rub", "delivery_to_warehouse_rub", "storage_days", "return_rate",
    "current_price_rub", "min_price_rub", "commission_rate", "logistics_rub",
    "last_mile_rate", "acquiring_rate", "drr_rate", "tax_rate", "other_rate",
]

SKU_CAPTIONS = [
    "Артикул продавца", "SKU Маркета", "Название", "ID категории", "Категория", "Схема",
    "Вес, г", "Длина, мм", "Ширина, мм", "Высота, мм", "Объём, л", "Себестоимость, ₽",
    "Упаковка, ₽", "Доставка до склада, ₽", "Хранение, дней", "Доля возвратов, %",
    "Текущая цена, ₽", "Мин. цена, ₽", "Комиссия, %", "Логистика, ₽",
    "Последняя миля, %", "Эквайринг, %", "ДРР, %", "Налог, %", "Прочее, %",
]


def _header_row(ws, row, headers, widths=None):
    for index, title in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=index, value=title)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    if widths:
        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(index)].width = width


def _sheet_intro(ws, title, note, last_col):
    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=13, bold=True, color=DARK)
    ws.row_dimensions[1].height = 24
    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = note
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="808080")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 26


def build_template(bundle: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Товары"
    _sheet_intro(
        ws, "Товары (SKU)",
        "Одна строка — один товар. Обязательные поля: offer_id, name, current_price_rub. "
        "Колонки с ставками (комиссия, логистика и т. д.) необязательны: если пусто — берутся из листа «Тарифы». "
        "Проценты указывайте числами: 15 или 0,15.",
        "Y",
    )
    _header_row(ws, 4, SKU_HEADERS, [12, 14, 34, 10, 18, 8, 9, 9, 9, 9, 9, 13, 11, 14, 11, 11, 13, 11, 10, 11, 12, 11, 9, 9, 9])

    skus = bundle.get("skus", [])[:5]
    for r, sku in enumerate(skus, start=5):
        values = [
            sku.get("offer_id"), sku.get("market_sku"), sku.get("name"),
            sku.get("category_id"), sku.get("category_name"), sku.get("scheme"),
            sku.get("weight_g"), sku.get("length_mm"), sku.get("width_mm"),
            sku.get("height_mm"), sku.get("volume_l"), sku.get("cogs_rub"),
            sku.get("packaging_rub"), sku.get("delivery_to_warehouse_rub"),
            sku.get("storage_days"), sku.get("return_rate"),
            sku.get("current_price_rub"), sku.get("min_price_rub"),
            sku.get("commission_rate"), sku.get("logistics_rub"),
            sku.get("last_mile_rate"), sku.get("acquiring_rate"),
            sku.get("drr_rate"), sku.get("tax_rate"), sku.get("other_rate"),
        ]
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.font = BODY
            cell.border = BORDER
            if c in (16, 19, 21, 22, 23, 24, 25):
                cell.number_format = PCT
            elif c in (12, 13, 14, 17, 18, 20):
                cell.number_format = MONEY
        ws.row_dimensions[r].height = 20
    ws.freeze_panes = "C5"

    # --- комиссии
    ws = wb.create_sheet("Комиссии")
    _sheet_intro(ws, "Комиссии по категориям",
                 "Ставка комиссии площадки за продажу. Ключ — категория и схема работы.", "E")
    _header_row(ws, 4, ["scheme", "category_id", "category_name", "commission_rate", "comment"],
                [10, 14, 28, 16, 44])
    row = 5
    for scheme, cfg in bundle.get("tariffs", {}).get("schemes", {}).items():
        for cat_id, rate in cfg.get("commission_by_category", {}).items():
            name = next((s.get("category_name") for s in bundle.get("skus", [])
                         if str(s.get("category_id")) == str(cat_id)), "")
            values = [scheme, cat_id, name, rate, ""]
            for c, value in enumerate(values, start=1):
                cell = ws.cell(row=row, column=c, value=value)
                cell.font = BODY
                cell.border = BORDER
                if c == 4:
                    cell.number_format = PCT
            row += 1
        values = [scheme, "default", "Прочие категории", cfg.get("default_commission", 0.17),
                  "Ставка, если категории нет в списке"]
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=c, value=value)
            cell.font = BODY
            cell.border = BORDER
            if c == 4:
                cell.number_format = PCT
        row += 1
    ws.freeze_panes = "A5"

    # --- логистика
    ws = wb.create_sheet("Логистика")
    _sheet_intro(ws, "Логистика по объёму",
                 "Ступени объёма упаковки (литры) и стоимость за единицу. Схема — FBY или FBS.", "E")
    _header_row(ws, 4, ["scheme", "volume_from_l", "volume_to_l", "price_rub", "comment"],
                [10, 16, 16, 14, 44])
    row = 5
    for scheme, cfg in bundle.get("tariffs", {}).get("schemes", {}).items():
        for tier in cfg.get("logistics_tiers", []):
            values = [scheme, tier.get("volume_from_l"), tier.get("volume_to_l"),
                      tier.get("price_rub"), ""]
            for c, value in enumerate(values, start=1):
                cell = ws.cell(row=row, column=c, value=value)
                cell.font = BODY
                cell.border = BORDER
                if c == 4:
                    cell.number_format = MONEY
            row += 1
    ws.freeze_panes = "A5"

    # --- прочие тарифы
    ws = wb.create_sheet("Прочие тарифы")
    _sheet_intro(ws, "Прочие тарифы",
                 "Последняя миля, эквайринг, хранение, возвраты, утилизация, штрафы. "
                 "Параметр — путь в настройках тарифов.", "E")
    _header_row(ws, 4, ["scheme", "parameter", "value", "unit", "comment"], [10, 34, 14, 18, 40])
    row = 5
    for scheme, cfg in bundle.get("tariffs", {}).get("schemes", {}).items():
        pairs = [
            ("last_mile.rate", cfg.get("last_mile", {}).get("rate"), "% от цены", "Ставка последней мили"),
            ("last_mile.min_rub", cfg.get("last_mile", {}).get("min_rub"), "₽", "Минимум за заказ"),
            ("last_mile.max_rub", cfg.get("last_mile", {}).get("max_rub"), "₽", "Максимум за заказ"),
            ("acquiring.rate", cfg.get("acquiring", {}).get("rate"), "% от суммы", "0, если включён в комиссию"),
            ("storage.rate_per_liter_day", cfg.get("storage", {}).get("rate_per_liter_day"), "₽/л/сутки", "Тариф хранения"),
            ("storage.free_days", cfg.get("storage", {}).get("free_days"), "дни", "Льготный период"),
            ("returns.fee_per_return_rub", cfg.get("returns", {}).get("fee_per_return_rub"), "₽", "Стоимость возврата"),
            ("utilization.fee_per_unit_rub", cfg.get("utilization", {}).get("fee_per_unit_rub"), "₽", "Утилизация единицы"),
            ("fines.rate", cfg.get("fines", {}).get("rate"), "% от выручки", "Плановые штрафы"),
            ("delivery_to_warehouse_rub", cfg.get("delivery_to_warehouse_rub"), "₽", "Доставка до склада"),
            ("commission_default", cfg.get("default_commission"), "%", "Комиссия по умолчанию"),
        ]
        for path, value, unit, comment in pairs:
            values = [scheme, path, value, unit, comment]
            for c, value2 in enumerate(values, start=1):
                cell = ws.cell(row=row, column=c, value=value2)
                cell.font = BODY
                cell.border = BORDER
                if c == 2:
                    cell.font = Font(name="Consolas", size=9)
            row += 1
    ws.freeze_panes = "A5"

    # --- акции
    ws = wb.create_sheet("Акции")
    _sheet_intro(ws, "Акции",
                 "Акции, в которые планируется вход. Для типа FIXED_PRICE заполните promo_price_rub, "
                 "для DISCOUNT — seller_discount_rate.", "K")
    _header_row(ws, 4, ["promo_id", "promo_name", "promo_type", "offer_id", "start_date",
                        "end_date", "seller_discount_rate", "promo_price_rub", "drr_rate", "comment"],
                [14, 26, 14, 12, 12, 12, 14, 14, 9, 40])
    row = 5
    for promo in bundle.get("promos", []):
        values = [promo.get("promo_id"), promo.get("promo_name"), promo.get("promo_type"),
                  promo.get("offer_id"), promo.get("start_date"), promo.get("end_date"),
                  promo.get("seller_discount_rate"), promo.get("promo_price_rub"),
                  promo.get("drr_rate"), promo.get("comment")]
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=c, value=value)
            cell.font = BODY
            cell.border = BORDER
            if c in (7, 9):
                cell.number_format = PCT
            if c == 8:
                cell.number_format = MONEY
        row += 1
    ws.freeze_panes = "A5"

    # --- настройки
    ws = wb.create_sheet("Настройки")
    _sheet_intro(ws, "Настройки расчёта", "Цели, налоги и округление цены.", "D")
    _header_row(ws, 4, ["Параметр", "Значение", "Ед.", "Комментарий"], [34, 18, 12, 52])
    settings = bundle.get("settings", {})
    rows = [
        ("Налоговый режим", settings.get("tax", {}).get("regime", "USN_INCOME"), "—", "USN_INCOME / USN_PROFIT / OSNO"),
        ("Ставка налога", settings.get("tax", {}).get("rate", 0.06), "%", "6% для УСН «доходы»"),
        ("Ставка НДС", settings.get("tax", {}).get("vat_rate", 0), "%", "0, если не платите НДС"),
        ("НДС включён в цену", str(settings.get("tax", {}).get("vat_in_price", False)).lower(), "—", "true/false"),
        ("Целевая маржинальность", settings.get("targets", {}).get("margin_rate", 0.2), "%", "Цель по марже"),
        ("Целевой ДРР", settings.get("targets", {}).get("drr_rate", 0.1), "%", "Плановая доля рекламы"),
        ("Минимальная маржа", settings.get("targets", {}).get("min_margin_rub", 200), "₽", "Порог проблемной позиции"),
        ("Округление цены", settings.get("rounding", {}).get("mode", "up_10"), "—", "none / up_1 / up_10 / up_50 / psych_99"),
        ("Шаг округления", settings.get("rounding", {}).get("step", 10), "₽", "Для режима up_N"),
        ("Срок хранения по умолчанию", settings.get("defaults", {}).get("storage_days", 45), "дни", "Если у товара не задан"),
        ("Доля возвратов по умолчанию", settings.get("defaults", {}).get("return_rate", 0.05), "%", "Если у товара не задана"),
        ("Прочие расходы", settings.get("defaults", {}).get("other_rate", 0.01), "%", "Брак, ПО, зарплаты"),
    ]
    for r, values in enumerate(rows, start=5):
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.font = BODY
            cell.border = BORDER
            if c == 2 and isinstance(value, float) and abs(value) <= 1:
                cell.number_format = PCT
        ws.row_dimensions[r].height = 20
    ws.freeze_panes = "A5"

    # --- инструкция
    guide = wb.create_sheet("Инструкция")
    guide.column_dimensions["A"].width = 4
    guide.column_dimensions["B"].width = 110
    guide.sheet_view.showGridLines = False
    guide.merge_cells("A1:B1")
    guide["A1"] = "Как загрузить свои данные"
    guide["A1"].font = Font(name="Calibri", size=16, bold=True, color=DARK)
    guide.row_dimensions[1].height = 30
    lines = [
        ("1. Заполните лист «Товары»: артикул, название, категория, схема, цена и себестоимость.", False),
        ("2. Проверьте листы «Комиссии», «Логистика», «Прочие тарифы» — это условия площадки.", False),
        ("3. Добавьте акции на листе «Акции», если нужно считать цены входа.", False),
        ("4. Сохраните файл в формате .xlsx и загрузите его кнопкой «Загрузить данные».", False),
        ("", False),
        ("Правила заполнения", True),
        ("• Проценты можно писать как 15 или 0,15 — калькулятор поймёт оба варианта.", False),
        ("• Пустые ставки в строке товара означают «взять из тарифов по категории».", False),
        ("• Если загрузить только лист «Товары», тарифы останутся прежними, а акции сбросятся.", False),
        ("• Категории сопоставляются по category_id; название нужно только для читаемости.", False),
        ("", False),
        ("Что считает приложение", True),
        ("• Маржу и маржинальность по текущей цене.", False),
        ("• Цену под целевую маржинальность и ДРР.", False),
        ("• Маржу и маржинальность по ценам входа в акции, включая предельную скидку.", False),
    ]
    row = 3
    for text, bold in lines:
        guide.merge_cells(f"A{row}:B{row}")
        cell = guide[f"A{row}"]
        cell.value = text
        cell.font = Font(name="Calibri", size=11, bold=bold, color="7A5B00" if bold else "333333")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        guide.row_dimensions[row].height = 20 if len(text) < 100 else 32
        row += 1

    stream = BytesIO()
    wb.save(stream)
    return stream.getvalue()


# ------------------------------------------------------------------ экспорт
def build_export(payload: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Юнит-экономика"
    ws.sheet_view.showGridLines = False

    widths = {"A": 6, "B": 34, "C": 20, "D": 8, "E": 12, "F": 12, "G": 12,
              "H": 13, "I": 13, "J": 13, "K": 13}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.merge_cells("A1:K1")
    ws["A1"] = "Юнит-экономика Яндекс Маркета"
    ws["A1"].font = Font(name="Calibri", size=18, bold=True, color=DARK)
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:K2")
    meta = payload.get("meta", {})
    ws["A2"] = (f"Источник данных: {meta.get('source', '')}  ·  "
                f"Расчёт от {datetime.now().strftime('%d.%m.%Y %H:%M')}  ·  "
                f"Цель по марже: {payload.get('targetMargin', 0)}%  ·  ДРР: {payload.get('targetDrr', 0)}%")
    ws["A2"].font = Font(name="Calibri", size=9, color="606060")

    headers = ["№", "Товар", "Категория", "Схема", "Цена", "Расходы", "Маржа",
               "Маржинальность", "Цена под цель", "Маржа под цель", "Статус"]
    for c, title in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=title)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[4].height = 30

    row = 5
    for index, item in enumerate(payload.get("rows", []), start=1):
        values = [
            index, item.get("name"), item.get("category"), item.get("scheme"), item.get("price"),
            item.get("total"), item.get("margin"), item.get("marginRate"),
            item.get("targetPrice"), item.get("targetMargin"), item.get("status"),
        ]
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=c, value=value)
            cell.font = BODY
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=(c == 2))
            if c in (5, 6, 7, 9, 10):
                cell.number_format = MONEY
            if c in (8,):
                cell.number_format = PCT
        ws.row_dimensions[row].height = 18
        row += 1

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:K{row - 1}"

    # --- акции
    ws2 = wb.create_sheet("Акции")
    ws2.sheet_view.showGridLines = False
    for col, width in {"A": 6, "B": 34, "C": 24, "D": 14, "E": 13, "F": 13,
                       "G": 13, "H": 13, "I": 30}.items():
        ws2.column_dimensions[col].width = width
    ws2.merge_cells("A1:I1")
    ws2["A1"] = "Цены входа в акции"
    ws2["A1"].font = Font(name="Calibri", size=18, bold=True, color=DARK)
    headers2 = ["№", "Товар", "Акция", "Цена акции", "Расходы", "Маржа",
                "Маржинальность", "Макс. скидка", "Вывод"]
    for c, title in enumerate(headers2, start=1):
        cell = ws2.cell(row=3, column=c, value=title)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    row = 4
    for index, promo in enumerate(payload.get("promos", []), start=1):
        values = [index, promo.get("name"), promo.get("promo"), promo.get("price"),
                  promo.get("total"), promo.get("margin"), promo.get("marginRate"),
                  promo.get("maxDiscount"), promo.get("verdict")]
        for c, value in enumerate(values, start=1):
            cell = ws2.cell(row=row, column=c, value=value)
            cell.font = BODY
            cell.border = BORDER
            if c in (4, 5, 6):
                cell.number_format = MONEY
            if c in (7, 8):
                cell.number_format = PCT
        row += 1
    ws2.freeze_panes = "A4"

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    stream = BytesIO()
    wb.save(stream)
    return stream.getvalue()
