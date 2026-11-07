# -*- coding: utf-8 -*-
"""Генерация Excel-файла: комиссии и расходы Яндекс Маркета + калькулятор юнит-экономики."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

OUT = "Яндекс_Маркет_комиссии_и_расходы.xlsx"

# ---------- стили ----------
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
F_SECTION = Font(name="Calibri", size=12, bold=True, color=C_DARK)

AL_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
AL_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

PCT = '0.0%'
RUB0 = '#,##0" ₽"'
RUB2 = '#,##0.00" ₽"'
NUM1 = '0.0'


def style_range(ws, cell_range, border=BORDER, fill=None, font=None, align=None):
    for row in ws[cell_range]:
        for c in row:
            if border:
                c.border = border
            if fill:
                c.fill = fill
            if font:
                c.font = font
            if align:
                c.alignment = align


def title(ws, text, last_col="K", sub=None):
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = text
    c.font = F_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    if sub:
        ws.merge_cells(f"A2:{last_col}2")
        s = ws["A2"]
        s.value = sub
        s.font = F_SUB
        s.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 28


wb = Workbook()

# =========================================================
# ЛИСТ 1 — КОМИССИИ И РАСХОДЫ
# =========================================================
ws = wb.active
ws.title = "Комиссии и расходы"
ws.sheet_properties.tabColor = C_MID
title(
    ws,
    "Яндекс Маркет: комиссии и расходы продавца",
    sub="Ориентировочные ставки для расчёта юнит-экономики (схема FBY). "
        "Значения в жёлтых ячейках можно менять — пример в последних колонках пересчитается.",
)

# --- входные параметры примера ---
labels = [
    ("A3:B3", "Цена продажи, ₽", "C3", 1500, RUB0),
    ("D3:E3", "Объём товара, л", "F3", 3, NUM1),
    ("G3:H3", "Срок хранения, дней", "I3", 45, '0'),
    ("A4:B4", "Доля возвратов", "C4", 0.06, PCT),
    ("D4:E4", "Логистика возврата, ₽", "F4", 60, RUB0),
    ("G4:H4", "Схема работы", "I4", "FBY", None),
]
for mrng, lab, vcell, val, fmt in labels:
    ws.merge_cells(mrng)
    lc = ws[mrng.split(":")[0]]
    lc.value = lab
    lc.font = F_BODY
    lc.alignment = AL_RIGHT
    lc.fill = PatternFill("solid", fgColor=C_LIGHT)
    vc = ws[vcell]
    vc.value = val
    vc.font = F_BOLD
    vc.alignment = AL_CENTER
    vc.fill = PatternFill("solid", fgColor=C_INPUT)
    vc.border = BORDER
    if fmt:
        vc.number_format = fmt

# --- шапка таблицы ---
HDR = 6
headers = [
    ("A", "№", 5),
    ("B", "Статья расхода", 34),
    ("C", "Кто удерживает", 18),
    ("D", "Как считается", 38),
    ("E", "Ставка", 11),
    ("F", "Мин.", 9),
    ("G", "Макс.", 9),
    ("H", "Единица", 18),
    ("I", "База расчёта", 16),
    ("J", "Сумма в примере, ₽", 15),
    ("K", "Комментарий", 44),
]
for col, text, width in headers:
    ws.column_dimensions[col].width = width
    c = ws[f"{col}{HDR}"]
    c.value = text
    c.font = F_HDR
    c.fill = PatternFill("solid", fgColor=C_DARK)
    c.alignment = AL_CENTER
    c.border = BORDER
ws.row_dimensions[HDR].height = 36

# name, who, how, rate, minv, maxv, unit, base, formula, comment, rate_fmt, mm_fmt
rows = [
    ("Комиссия площадки (за продажу)", "Яндекс Маркет",
     "Процент от цены товара. Зависит от категории и схемы работы (FBY / FBS / DBS)",
     0.15, 0.03, 0.25, "% от цены", "Цена продажи", "price",
     "Основная комиссия. Точная ставка — в кабинете: Тарифы → Комиссия за продажу", PCT, PCT),

    ("Логистика (приёмка, размещение, сборка заказа)", "Яндекс Маркет",
     "Фиксированная сумма за единицу, зависит от объёма, веса и схемы",
     45, 20, 150, "₽ / единица", "—", "rate0",
     "FBY: обработка на складе + сборка заказа. Для крупногабарита выше", RUB0, RUB0),

    ("Последняя миля (доставка покупателю)", "Яндекс Маркет",
     "Процент от цены с нижним и верхним лимитом",
     0.055, 25, 500, "% от цены (мин/макс — ₽)", "Цена продажи", "lastmile",
     "Например: 5,5% от цены, но не менее 25 ₽ и не более 500 ₽", PCT, RUB0),

    ("Эквайринг", "Банк / платёжный сервис",
     "Процент от суммы платежа. На Маркете обычно включён в комиссию",
     0.0, 0.0, 0.025, "% от суммы платежа", "Сумма платежа", "price",
     "Отдельная ставка — только при своей кассе (DBS, свой сайт): 1,5–2,5%", PCT, PCT),

    ("Хранение на складе", "Яндекс Маркет",
     "Тариф за литр в сутки × объём товара × срок хранения",
     0.15, 0.05, 0.50, "₽ / л в сутки", "Объём × дни", "storage",
     "Зависит от оборачиваемости; для неликвида тариф заметно выше", RUB2, RUB2),

    ("Возвраты", "Яндекс Маркет",
     "Доля возвратов × стоимость обратной логистики (+ потеря комиссии)",
     0.06, 0.02, 0.15, "доля возвратов", "—", "returns",
     "В примере: 6% возвратов × 60 ₽ = 3,6 ₽ на проданную единицу", PCT, PCT),

    ("Утилизация / штрафы", "Яндекс Маркет",
     "Фикс за утилизацию единицы + штрафы за нарушения",
     0.007, 0.0, 0.02, "% от выручки", "Выручка", "price",
     "Просрочка поставок, брак упаковки, ошибки маркировки. Планово 0,5–1%", PCT, PCT),

    ("Доставка до склада (транспорт)", "Продавец",
     "Стоимость перевозки партии / количество единиц",
     25, 10, 60, "₽ / единица", "—", "rate0",
     "Своя логистика или перевозчик; при мелких партиях дороже", RUB0, RUB0),

    ("Реклама / ДРР", "Продавец",
     "Расходы на продвижение / выручка × 100%",
     0.10, 0.03, 0.25, "% от выручки", "Выручка", "price",
     "Продвижение, буст, промо и скидки за счёт продавца", PCT, PCT),

    ("Налоги", "Продавец",
     "Процент от выручки (УСН «доходы») или от прибыли",
     0.06, 0.0, 0.20, "% от выручки", "Выручка", "price",
     "УСН 6%, АУСН, НДС при ОСНО. С 2025 НДС для УСН при выручке свыше 60 млн ₽", PCT, PCT),

    ("Себестоимость товара", "Продавец",
     "Закупка или производство за единицу",
     600, None, None, "₽ / единица", "—", "rate0",
     "Без возмещаемого НДС; закладывайте брак и потери", RUB0, RUB0),

    ("Упаковка", "Продавец",
     "Стоимость упаковки за единицу",
     20, 5, 50, "₽ / единица", "—", "rate0",
     "Коробка, пакет, наполнитель, маркировка, этикетки", RUB0, RUB0),

    ("Прочее", "Продавец",
     "Процент от выручки (или фиксированная сумма на единицу)",
     0.01, 0.0, 0.05, "% от выручки", "Выручка", "price",
     "Брак и списания, ПО и сервисы, зарплаты, страховые взносы", PCT, PCT),
]

first_data = HDR + 1
for i, (name, who, how, rate, minv, maxv, unit, base, fkey, comment, rfmt, mfmt) in enumerate(rows):
    r = first_data + i
    fill = PatternFill("solid", fgColor=C_ALT) if i % 2 else None

    ws[f"A{r}"] = i + 1
    ws[f"A{r}"].alignment = AL_CENTER
    ws[f"A{r}"].font = F_BODY

    ws[f"B{r}"] = name
    ws[f"B{r}"].font = F_BOLD
    ws[f"B{r}"].alignment = AL_WRAP

    ws[f"C{r}"] = who
    ws[f"C{r}"].font = F_BODY
    ws[f"C{r}"].alignment = AL_WRAP

    ws[f"D{r}"] = how
    ws[f"D{r}"].font = F_BODY
    ws[f"D{r}"].alignment = AL_WRAP

    ws[f"E{r}"] = rate
    ws[f"E{r}"].number_format = rfmt
    ws[f"E{r}"].font = F_BOLD
    ws[f"E{r}"].alignment = AL_CENTER

    if minv is not None:
        ws[f"F{r}"] = minv
        ws[f"F{r}"].number_format = mfmt
        ws[f"F{r}"].alignment = AL_CENTER
        ws[f"F{r}"].font = F_BODY
    if maxv is not None:
        ws[f"G{r}"] = maxv
        ws[f"G{r}"].number_format = mfmt
        ws[f"G{r}"].alignment = AL_CENTER
        ws[f"G{r}"].font = F_BODY

    ws[f"H{r}"] = unit
    ws[f"H{r}"].font = F_BODY
    ws[f"H{r}"].alignment = AL_CENTER

    ws[f"I{r}"] = base
    ws[f"I{r}"].font = F_BODY
    ws[f"I{r}"].alignment = AL_CENTER

    if fkey == "price":
        formula = f"=$C$3*E{r}"
    elif fkey == "rate0":
        formula = f"=E{r}"
    elif fkey == "lastmile":
        formula = f"=MIN(G{r},MAX(F{r},$C$3*E{r}))"
    elif fkey == "storage":
        formula = f"=E{r}*$F$3*$I$3"
    elif fkey == "returns":
        formula = "=$C$4*$F$4"
    else:
        formula = f"=E{r}"

    ws[f"J{r}"] = formula
    ws[f"J{r}"].number_format = RUB2
    ws[f"J{r}"].font = F_BODY
    ws[f"J{r}"].alignment = AL_RIGHT

    ws[f"K{r}"] = comment
    ws[f"K{r}"].font = Font(name="Calibri", size=9, color="404040")
    ws[f"K{r}"].alignment = AL_WRAP

    style_range(ws, f"A{r}:K{r}", fill=fill)
    ws.row_dimensions[r].height = 40

last_data = first_data + len(rows) - 1

# --- итоговая строка ---
tot = last_data + 1
ws.merge_cells(f"A{tot}:E{tot}")
ws[f"A{tot}"] = "ИТОГО расходов на 1 единицу товара"
ws[f"A{tot}"].font = Font(name="Calibri", size=11, bold=True, color=C_DARK)
ws[f"A{tot}"].alignment = AL_RIGHT
ws[f"J{tot}"] = f"=SUM(J{first_data}:J{last_data})"
ws[f"J{tot}"].number_format = RUB2
ws[f"J{tot}"].font = Font(name="Calibri", size=11, bold=True, color=C_DARK)
ws[f"J{tot}"].alignment = AL_RIGHT
for col in "FGHIK":
    ws[f"{col}{tot}"].border = BORDER
ws[f"J{tot}"].border = BORDER
style_range(ws, f"A{tot}:K{tot}", fill=PatternFill("solid", fgColor=C_TOTAL))
ws.row_dimensions[tot].height = 26

# --- итоги ---
s = tot + 2
ws.merge_cells(f"A{s}:K{s}")
ws[f"A{s}"] = "Итоги по примеру"
ws[f"A{s}"].font = F_SECTION
ws[f"A{s}"].alignment = Alignment(horizontal="left", vertical="center")

summary = [
    ("Цена продажи (выручка с единицы)", "=$C$3", RUB2),
    ("Итого расходов", f"=J{tot}", RUB2),
    ("Прибыль с единицы", f"=C{s+1}-C{s+2}", RUB2),
    ("Маржинальность (прибыль / цена)", f"=IFERROR(C{s+3}/C{s+1},0)", PCT),
    ("ROI (прибыль / себестоимость)", f"=IFERROR(C{s+3}/J{first_data + 10},0)", PCT),
    ("Безубыточная цена", (
        f"=IFERROR((J{first_data + 1}+J{first_data + 4}+J{first_data + 5}+J{first_data + 7}"
        f"+J{first_data + 10}+J{first_data + 11})/"
        f"(1-(E{first_data}+E{first_data + 2}+E{first_data + 3}+E{first_data + 6}"
        f"+E{first_data + 8}+E{first_data + 9}+E{first_data + 12})),0)"
    ), RUB2),
]
for i, (lab, formula, fmt) in enumerate(summary):
    r = s + 1 + i
    ws.merge_cells(f"A{r}:B{r}")
    ws[f"A{r}"] = lab
    ws[f"A{r}"].font = F_BODY
    ws[f"A{r}"].alignment = AL_RIGHT
    ws[f"A{r}"].fill = PatternFill("solid", fgColor=C_LIGHT)
    ws[f"C{r}"] = formula
    ws[f"C{r}"].number_format = fmt
    ws[f"C{r}"].font = F_BOLD
    ws[f"C{r}"].alignment = AL_CENTER
    ws[f"C{r}"].border = BORDER
    ws.row_dimensions[r].height = 20

note_r = s + len(summary) + 2
ws.merge_cells(f"A{note_r}:K{note_r}")
ws[f"A{note_r}"] = (
    "Ставки ориентировочные (ноябрь 2026) и зависят от категории, схемы работы, региона и объёма. "
    "Перед расчётом цены сверьте тарифы в кабинете продавца. Файл не является офертой."
)
ws[f"A{note_r}"].font = Font(name="Calibri", size=9, italic=True, color=C_GRAY)
ws[f"A{note_r}"].alignment = AL_WRAP
ws.row_dimensions[note_r].height = 30

ws.freeze_panes = f"C{first_data}"
ws.auto_filter.ref = f"A{HDR}:K{last_data}"
ws.sheet_view.zoomScale = 90
ws.page_setup.orientation = "landscape"
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

# =========================================================
# ЛИСТ 2 — КАЛЬКУЛЯТОР
# =========================================================
w2 = wb.create_sheet("Калькулятор")
w2.sheet_properties.tabColor = "C00000"
title(
    w2,
    "Калькулятор юнит-экономики",
    last_col="D",
    sub="Меняйте только жёлтые ячейки. Суммы, прибыль и безубыточная цена считаются формулами.",
)
for col, w in (("A", 36), ("B", 18), ("C", 12), ("D", 58)):
    w2.column_dimensions[col].width = w

w2.merge_cells("A4:D4")
w2["A4"] = "1. Вводные параметры"
w2["A4"].font = F_SECTION
w2["A4"].alignment = Alignment(horizontal="left", vertical="center")

for i, text in enumerate(["Параметр", "Значение", "Ед.", "Комментарий"]):
    c = w2.cell(row=5, column=i + 1, value=text)
    c.font = F_HDR
    c.fill = PatternFill("solid", fgColor=C_DARK)
    c.alignment = AL_CENTER
    c.border = BORDER
w2.row_dimensions[5].height = 24

inputs = [
    ("Цена продажи", 1500, "₽", RUB0, "Цена, которую платит покупатель"),
    ("Себестоимость товара", 600, "₽", RUB0, "Закупка или производство за единицу"),
    ("Комиссия площадки", 0.15, "%", PCT, "По вашей категории (кабинет → Тарифы)"),
    ("Логистика", 45, "₽/ед.", RUB0, "Обработка и сборка заказа"),
    ("Последняя миля", 0.055, "%", PCT, "Процент от цены"),
    ("Мин. последняя миля", 25, "₽", RUB0, "Нижний лимит тарифа"),
    ("Макс. последняя миля", 500, "₽", RUB0, "Верхний лимит тарифа"),
    ("Эквайринг", 0.0, "%", PCT, "0% — если включён в комиссию Маркета"),
    ("Объём товара", 3, "л", NUM1, "Для расчёта хранения"),
    ("Срок хранения", 45, "дней", '0', "Средний срок до продажи"),
    ("Тариф хранения", 0.15, "₽/л/сутки", RUB2, "Зависит от оборачиваемости"),
    ("Доля возвратов", 0.06, "%", PCT, "Доля заказов, которые вернулись"),
    ("Стоимость возврата", 60, "₽", RUB0, "Обратная логистика + обработка"),
    ("Утилизация / штрафы", 0.007, "%", PCT, "Планово 0,5–1% от выручки"),
    ("Доставка до склада", 25, "₽/ед.", RUB0, "Транспортные расходы на единицу"),
    ("Реклама / ДРР", 0.10, "%", PCT, "Доля рекламы в выручке"),
    ("Налог", 0.06, "%", PCT, "УСН 6% / АУСН / НДС"),
    ("Упаковка", 20, "₽/ед.", RUB0, "Коробка, пакет, маркировка"),
    ("Прочее", 0.01, "%", PCT, "Брак, ПО, зарплаты и т. п."),
]
IN = {}
start_in = 6
for i, (lab, val, unit, fmt, comment) in enumerate(inputs):
    r = start_in + i
    IN[lab] = r
    w2[f"A{r}"] = lab
    w2[f"A{r}"].font = F_BODY
    w2[f"A{r}"].alignment = AL_WRAP
    w2[f"A{r}"].border = BORDER

    w2[f"B{r}"] = val
    w2[f"B{r}"].number_format = fmt
    w2[f"B{r}"].font = F_BOLD
    w2[f"B{r}"].alignment = AL_CENTER
    w2[f"B{r}"].fill = PatternFill("solid", fgColor=C_INPUT)
    w2[f"B{r}"].border = BORDER

    w2[f"C{r}"] = unit
    w2[f"C{r}"].font = F_BODY
    w2[f"C{r}"].alignment = AL_CENTER
    w2[f"C{r}"].border = BORDER

    w2[f"D{r}"] = comment
    w2[f"D{r}"].font = Font(name="Calibri", size=9, color="404040")
    w2[f"D{r}"].alignment = AL_WRAP
    w2[f"D{r}"].border = BORDER
    w2.row_dimensions[r].height = 20

price = f"$B${IN['Цена продажи']}"
cost = f"$B${IN['Себестоимость товара']}"
comm = f"$B${IN['Комиссия площадки']}"
logi = f"$B${IN['Логистика']}"
lm_p = f"$B${IN['Последняя миля']}"
lm_min = f"$B${IN['Мин. последняя миля']}"
lm_max = f"$B${IN['Макс. последняя миля']}"
acq = f"$B${IN['Эквайринг']}"
vol = f"$B${IN['Объём товара']}"
days = f"$B${IN['Срок хранения']}"
stor_t = f"$B${IN['Тариф хранения']}"
ret_d = f"$B${IN['Доля возвратов']}"
ret_c = f"$B${IN['Стоимость возврата']}"
util = f"$B${IN['Утилизация / штрафы']}"
deliv = f"$B${IN['Доставка до склада']}"
ads = f"$B${IN['Реклама / ДРР']}"
tax = f"$B${IN['Налог']}"
pack = f"$B${IN['Упаковка']}"
other = f"$B${IN['Прочее']}"

# --- блок результатов ---
res_title = start_in + len(inputs) + 1
w2.merge_cells(f"A{res_title}:D{res_title}")
w2[f"A{res_title}"] = "2. Расчёт по статьям"
w2[f"A{res_title}"].font = F_SECTION
w2[f"A{res_title}"].alignment = Alignment(horizontal="left", vertical="center")

rh = res_title + 1
for i, text in enumerate(["Статья", "Сумма, ₽", "% от цены", "Формула"]):
    c = w2.cell(row=rh, column=i + 1, value=text)
    c.font = F_HDR
    c.fill = PatternFill("solid", fgColor=C_DARK)
    c.alignment = AL_CENTER
    c.border = BORDER
w2.row_dimensions[rh].height = 24

results = [
    ("Комиссия площадки", f"={price}*{comm}", "% от цены"),
    ("Логистика", f"={logi}", "фикс за единицу"),
    ("Последняя миля", f"=MIN({lm_max},MAX({lm_min},{price}*{lm_p}))", "% с лимитами"),
    ("Эквайринг", f"={price}*{acq}", "% от суммы"),
    ("Хранение", f"={stor_t}*{vol}*{days}", "₽/л/сутки × объём × дни"),
    ("Возвраты", f"={ret_d}*{ret_c}", "доля × стоимость возврата"),
    ("Утилизация / штрафы", f"={price}*{util}", "% от выручки"),
    ("Доставка до склада", f"={deliv}", "фикс за единицу"),
    ("Реклама / ДРР", f"={price}*{ads}", "% от выручки"),
    ("Налоги", f"={price}*{tax}", "% от выручки"),
    ("Себестоимость", f"={cost}", "за единицу"),
    ("Упаковка", f"={pack}", "за единицу"),
    ("Прочее", f"={price}*{other}", "% от выручки"),
]
first_res = rh + 1
for i, (name, formula, how) in enumerate(results):
    r = first_res + i
    fill = PatternFill("solid", fgColor=C_ALT) if i % 2 else None
    w2[f"A{r}"] = name
    w2[f"A{r}"].font = F_BODY
    w2[f"A{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    w2[f"B{r}"] = formula
    w2[f"B{r}"].number_format = RUB2
    w2[f"B{r}"].font = F_BODY
    w2[f"B{r}"].alignment = AL_RIGHT

    w2[f"C{r}"] = f"=IFERROR(B{r}/{price},0)"
    w2[f"C{r}"].number_format = PCT
    w2[f"C{r}"].font = F_BODY
    w2[f"C{r}"].alignment = AL_CENTER

    w2[f"D{r}"] = how
    w2[f"D{r}"].font = Font(name="Calibri", size=9, color="404040")
    w2[f"D{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    style_range(w2, f"A{r}:D{r}", fill=fill)
    w2.row_dimensions[r].height = 20

last_res = first_res + len(results) - 1
tot2 = last_res + 1
w2[f"A{tot2}"] = "ИТОГО расходов"
w2[f"A{tot2}"].font = Font(name="Calibri", size=11, bold=True, color=C_DARK)
w2[f"A{tot2}"].alignment = AL_RIGHT
w2[f"B{tot2}"] = f"=SUM(B{first_res}:B{last_res})"
w2[f"B{tot2}"].number_format = RUB2
w2[f"B{tot2}"].font = Font(name="Calibri", size=11, bold=True, color=C_DARK)
w2[f"B{tot2}"].alignment = AL_RIGHT
w2[f"C{tot2}"] = f"=IFERROR(B{tot2}/{price},0)"
w2[f"C{tot2}"].number_format = PCT
w2[f"C{tot2}"].font = Font(name="Calibri", size=11, bold=True, color=C_DARK)
w2[f"C{tot2}"].alignment = AL_CENTER
style_range(w2, f"A{tot2}:D{tot2}", fill=PatternFill("solid", fgColor=C_TOTAL))
w2.row_dimensions[tot2].height = 24

# --- блок итогов ---
fin = tot2 + 2
w2.merge_cells(f"A{fin}:D{fin}")
w2[f"A{fin}"] = "3. Результат"
w2[f"A{fin}"].font = F_SECTION
w2[f"A{fin}"].alignment = Alignment(horizontal="left", vertical="center")

finals = [
    ("Выручка с единицы", f"={price}", RUB2),
    ("Итого расходов", f"=B{tot2}", RUB2),
    ("Прибыль с единицы", f"={price}-B{tot2}", RUB2),
    ("Маржинальность", f"=IFERROR(({price}-B{tot2})/{price},0)", PCT),
    ("ROI (прибыль / себестоимость)", f"=IFERROR(({price}-B{tot2})/{cost},0)", PCT),
    ("Безубыточная цена", (
        f"=IFERROR((B{first_res + 1}+B{first_res + 4}+B{first_res + 5}+B{first_res + 7}"
        f"+B{first_res + 10}+B{first_res + 11})/"
        f"(1-({comm}+{lm_p}+{acq}+{util}+{ads}+{tax}+{other})),0)"
    ), RUB2),
]
for i, (lab, formula, fmt) in enumerate(finals):
    r = fin + 1 + i
    w2[f"A{r}"] = lab
    w2[f"A{r}"].font = F_BOLD if i >= 2 else F_BODY
    w2[f"A{r}"].alignment = AL_RIGHT
    w2[f"A{r}"].fill = PatternFill("solid", fgColor=C_LIGHT)
    w2[f"B{r}"] = formula
    w2[f"B{r}"].number_format = fmt
    w2[f"B{r}"].font = Font(name="Calibri", size=11, bold=True,
                            color="C00000" if i in (2, 3, 4) else C_DARK)
    w2[f"B{r}"].alignment = AL_CENTER
    style_range(w2, f"A{r}:B{r}")
    w2.row_dimensions[r].height = 22

w2.merge_cells(f"A{fin + len(finals) + 2}:D{fin + len(finals) + 2}")
w2[f"A{fin + len(finals) + 2}"] = (
    "Возвраты считаются как доля возвратов × стоимость одного возврата. "
    "Если возврат сопровождается потерей комиссии и прямой логистики — увеличьте стоимость возврата."
)
w2[f"A{fin + len(finals) + 2}"].font = Font(name="Calibri", size=9, italic=True, color=C_GRAY)
w2[f"A{fin + len(finals) + 2}"].alignment = AL_WRAP
w2.row_dimensions[fin + len(finals) + 2].height = 32

w2.freeze_panes = "A6"
w2.sheet_view.zoomScale = 100
w2.page_setup.orientation = "portrait"
w2.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
w2.page_setup.fitToWidth = 1
w2.page_setup.fitToHeight = 0

# =========================================================
# ЛИСТ 3 — СПРАВКА
# =========================================================
w3 = wb.create_sheet("Справка")
w3.sheet_properties.tabColor = "7F7F7F"
w3.column_dimensions["A"].width = 4
w3.column_dimensions["B"].width = 46
w3.column_dimensions["C"].width = 34
w3.column_dimensions["D"].width = 42

w3.merge_cells("A1:D1")
w3["A1"] = "Справка: как пользоваться и где проверить тарифы"
w3["A1"].font = F_TITLE
w3["A1"].alignment = Alignment(horizontal="left", vertical="center")
w3.row_dimensions[1].height = 32

w3.merge_cells("A2:D2")
w3["A2"] = "Версия шаблона: ноябрь 2026. Ставки — ориентировочные, перед расчётом цены сверьте их в кабинете продавца."
w3["A2"].font = F_SUB
w3["A2"].alignment = AL_WRAP
w3.row_dimensions[2].height = 30


def block(r, heading, lines, numbered=False):
    w3.merge_cells(f"A{r}:D{r}")
    c = w3[f"A{r}"]
    c.value = heading
    c.font = F_SECTION
    c.fill = PatternFill("solid", fgColor=C_LIGHT)
    c.alignment = Alignment(horizontal="left", vertical="center")
    w3.row_dimensions[r].height = 24
    r += 1
    for i, line in enumerate(lines, start=1):
        w3.merge_cells(f"A{r}:D{r}")
        c = w3[f"A{r}"]
        c.value = (f"{i}. " if numbered else "•  ") + line
        c.font = F_BODY
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        w3.row_dimensions[r].height = 18 if len(line) < 95 else 32
        r += 1
    return r + 1


r = 4
r = block(r, "Что внутри", [
    "«Комиссии и расходы» — 13 статей затрат с ориентировочными ставками и расчётом на примере товара за 1 500 ₽.",
    "«Калькулятор» — подставьте свои цифры, суммы, прибыль и безубыточная цена пересчитаются автоматически.",
    "«Справка» — методика, ориентиры по категориям и чек-лист проверки тарифов.",
])

r = block(r, "Как пользоваться", [
    "Откройте лист «Калькулятор» и замените значения в жёлтых ячейках на свои.",
    "Проверьте комиссию по своей категории: от неё зависит больше всего.",
    "Уточните логистику и последнюю милю под объём и вес именно вашего товара.",
    "Посмотрите блок «Результат»: прибыль с единицы, маржинальность, ROI и безубыточная цена.",
    "Сверьте итог с листом «Комиссии и расходы», чтобы ничего не забыть.",
], numbered=True)

r = block(r, "Где взять точные тарифы", [
    "Кабинет Яндекс Маркета → Настройки → Тарифы: комиссия по вашей категории, логистика, хранение.",
    "Калькулятор тарифов в кабинете продавца — показывает расчёт по конкретному товару.",
    "Справка Маркета: yandex.ru/support/marketplace (разделы «Комиссии», «Логистика», «Хранение»).",
    "Договор-оферта и приложения с тарифами — там зафиксированы ваши условия.",
])

# --- таблица ориентиров по категориям ---
w3.merge_cells(f"A{r}:D{r}")
c = w3[f"A{r}"]
c.value = "Ориентировочные комиссии по категориям (проверяйте в кабинете)"
c.font = F_SECTION
c.fill = PatternFill("solid", fgColor=C_LIGHT)
c.alignment = Alignment(horizontal="left", vertical="center")
w3.row_dimensions[r].height = 24
r += 1

for i, text in enumerate(["Категория", "Комиссия за продажу", "Комментарий"]):
    col = 2 + i
    cell = w3.cell(row=r, column=col, value=text)
    cell.font = F_HDR
    cell.fill = PatternFill("solid", fgColor=C_DARK)
    cell.alignment = AL_CENTER
    cell.border = BORDER
w3.row_dimensions[r].height = 24
r += 1

cats = [
    ("Электроника и техника", "5–10%", "Нижняя граница рынка, высокая конкуренция"),
    ("Бытовая техника", "5–12%", "Крупногабарит — дороже логистика"),
    ("Одежда", "20–25%", "Высокая доля возвратов и невыкупов"),
    ("Обувь", "20–25%", "Обязательна примерка, возвраты выше среднего"),
    ("Красота и косметика", "15–20%", "Сроки годности, требования к маркировке"),
    ("Детские товары", "12–18%", "Сертификация и упаковка"),
    ("Товары для дома", "12–18%", "Средняя маржинальность"),
    ("Продукты питания", "8–14%", "Короткие сроки, высокая оборачиваемость"),
    ("Зоотовары", "12–18%", "Тяжёлые корма — следите за логистикой"),
    ("Спорт и отдых", "10–16%", "Сезонность спроса"),
    ("Автотовары", "8–14%", "Крупные позиции — дорогая логистика"),
    ("Аксессуары и ювелирка", "15–22%", "Мелкий объём, невысокая логистика"),
]
for i, (name, rate, comment) in enumerate(cats):
    fill = PatternFill("solid", fgColor=C_ALT) if i % 2 else None
    for col, val in ((2, name), (3, rate), (4, comment)):
        cell = w3.cell(row=r, column=col, value=val)
        cell.font = F_BOLD if col == 3 else F_BODY
        cell.alignment = AL_CENTER if col == 3 else AL_WRAP
        cell.border = BORDER
        if fill:
            cell.fill = fill
    w3.row_dimensions[r].height = 20
    r += 1

r += 1
r = block(r, "Важно", [
    "Все ставки в файле — ориентировочные, для планирования. Это не оферта и не официальные тарифы Маркета.",
    "Итоговая комиссия зависит от категории, схемы (FBY / FBS / DBS), региона, объёма и оборачиваемости товара.",
    "С 2025 года упрощенцы с выручкой свыше 60 млн ₽ платят НДС — учитывайте это в статье «Налоги».",
    "Реклама и скидки за счёт продавца сильно меняют экономику: закладывайте ДРР на вывод товара в топ.",
    "Считайте несколько сценариев: пессимистичный (высокие возвраты и ДРР) и базовый.",
])

r = block(r, "Чек-лист перед запуском", [
    "Комиссия по категории актуальна на дату расчёта.",
    "Логистика и последняя миля проверены под объём и вес товара.",
    "Учтены возвраты и невыкупы, а не только успешные продажи.",
    "Заложен бюджет на рекламу и промо.",
    "Учтены налоги, НДС и страховые взносы.",
    "Посчитана безубыточная цена и минимальная допустимая наценка.",
    "Проверены штрафы, утилизация и требования к упаковке.",
])

w3.sheet_view.showGridLines = False
w3.page_setup.orientation = "portrait"
w3.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
w3.page_setup.fitToWidth = 1
w3.page_setup.fitToHeight = 0

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved:", OUT)
