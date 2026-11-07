# -*- coding: utf-8 -*-
"""Разбор загруженных таблиц (xlsx / csv) в набор данных калькулятора."""

from __future__ import annotations

import csv
import io
import re
from datetime import datetime

from openpyxl import load_workbook


class DataError(Exception):
    """Понятная пользователю ошибка разбора файла."""


def norm(value) -> str:
    text = str(value or "").lower().replace("ё", "е").replace("\xa0", " ")
    text = re.sub(r"[^0-9a-zа-я%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def num(value, default=None):
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("\xa0", " ").strip()
    if not text:
        return default
    negative = text.startswith("-")
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return default
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    if text.count(".") > 1:
        head, _, tail = text.rpartition(".")
        text = head.replace(".", "") + "." + tail
    try:
        result = float(text)
    except ValueError:
        return default
    return -result if negative else result


def rate(value, default=None):
    """Ставка: 15 -> 0.15, '15%' -> 0.15, 0.15 -> 0.15."""
    if value is None:
        return default
    text = str(value).strip()
    has_percent = "%" in text
    value_num = num(text, None)
    if value_num is None:
        return default
    if has_percent or value_num > 1:
        return value_num / 100.0
    return value_num


def flag(value, default=False):
    text = norm(value)
    if text in ("true", "да", "1", "yes", "включено", "вкл"):
        return True
    if text in ("false", "нет", "0", "no", "выключено", "выкл"):
        return False
    return default


# ------------------------------------------------------------------ алиасы
ALIASES_SKU = {
    "offer_id": ["offer_id", "offer id", "артикул продавца", "артикул", "ваш sku", "скю продавца", "id товара", "sku продавца", "sku"],
    "market_sku": ["market_sku", "sku маркета", "маркет sku", "id на маркете"],
    "name": ["name", "наименование", "название товара", "название", "товар"],
    "category_id": ["category_id", "id категории", "код категории", "категория id"],
    "category_name": ["category_name", "название категории", "категория"],
    "scheme": ["scheme", "схема работы", "схема"],
    "weight_g": ["weight_g", "вес г", "вес, г", "вес"],
    "length_mm": ["length_mm", "длина мм", "длина"],
    "width_mm": ["width_mm", "ширина мм", "ширина"],
    "height_mm": ["height_mm", "высота мм", "высота"],
    "volume_l": ["volume_l", "объем л", "объём л", "объем", "объём", "литраж"],
    "cogs_rub": ["cogs_rub", "себестоимость", "закупка", "закупочная цена", "cost"],
    "packaging_rub": ["packaging_rub", "упаковка"],
    "delivery_to_warehouse_rub": ["delivery_to_warehouse_rub", "доставка до склада", "доставка на склад"],
    "storage_days": ["storage_days", "срок хранения", "хранение дней", "дней хранения"],
    "return_rate": ["return_rate", "доля возвратов", "возвраты", "процент возвратов"],
    "current_price_rub": ["current_price_rub", "price_rub", "текущая цена", "цена продажи", "цена, руб", "цена"],
    "min_price_rub": ["min_price_rub", "минимальная цена", "мин цена"],
    "commission_rate": ["commission_rate", "комиссия площадки", "комиссия"],
    "logistics_rub": ["logistics_rub", "логистика руб", "стоимость логистики", "логистика"],
    "last_mile_rate": ["last_mile_rate", "последняя миля"],
    "acquiring_rate": ["acquiring_rate", "эквайринг"],
    "drr_rate": ["drr_rate", "дрр", "реклама"],
    "tax_rate": ["tax_rate", "налог"],
    "other_rate": ["other_rate", "прочие расходы", "прочее"],
}

ALIASES_PROMO = {
    "promo_id": ["promo_id", "id акции", "акция id"],
    "promo_name": ["promo_name", "название акции", "акция"],
    "promo_type": ["promo_type", "тип акции", "тип"],
    "offer_id": ["offer_id", "артикул продавца", "артикул", "sku"],
    "start_date": ["start_date", "начало", "дата начала"],
    "end_date": ["end_date", "окончание", "дата окончания"],
    "seller_discount_rate": ["seller_discount_rate", "скидка продавца", "скидка"],
    "promo_price_rub": ["promo_price_rub", "цена в акции", "цена акции"],
    "drr_rate": ["drr_rate", "дрр в акции", "дрр"],
    "comment": ["comment", "комментарий"],
}

ALIASES_COMM = {
    "scheme": ["scheme", "схема"],
    "category_id": ["category_id", "id категории", "категория id"],
    "category_name": ["category_name", "категория"],
    "commission_rate": ["commission_rate", "комиссия", "ставка"],
}

ALIASES_LOG = {
    "scheme": ["scheme", "схема"],
    "volume_from_l": ["volume_from_l", "объем от", "объём от", "от л"],
    "volume_to_l": ["volume_to_l", "объем до", "объём до", "до л"],
    "price_rub": ["price_rub", "стоимость", "цена", "тариф"],
}

ALIASES_TAR = {
    "scheme": ["scheme", "схема"],
    "parameter": ["parameter", "параметр"],
    "value": ["value", "значение"],
}

ALIASES_SET = {
    "label": ["параметр", "настройка", "показатель"],
    "value": ["значение"],
}


def match_field(header, aliases):
    text = norm(header)
    if not text:
        return None
    best, best_len = None, 0
    for field, variants in aliases.items():
        for variant in variants:
            variant_n = norm(variant)
            if variant_n and variant_n in text and len(variant_n) > best_len:
                best, best_len = field, len(variant_n)
    return best


def map_row(row, aliases):
    mapping = {}
    for index, cell in enumerate(row):
        field = match_field(cell, aliases)
        if field and field not in mapping:
            mapping[field] = index
    return mapping


def find_header(rows, aliases, required, limit=25):
    best = (None, None, 0)
    for index, row in enumerate(rows[:limit]):
        mapping = map_row(row, aliases)
        if all(key in mapping for key in required) and len(mapping) > best[2]:
            best = (index, mapping, len(mapping))
    return best


def cell(row, mapping, key):
    index = mapping.get(key)
    if index is None or index >= len(row):
        return None
    return row[index]


# ------------------------------------------------------------------ чтение
def read_sheets(data: bytes, filename: str):
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if ext in ("xlsx", "xlsm"):
        try:
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        except Exception as exc:  # noqa: BLE001
            raise DataError("Не удалось открыть файл Excel. Нужен формат .xlsx.") from exc
        sheets = []
        for sheet in wb.worksheets:
            rows = [list(r) for r in sheet.iter_rows(values_only=True)]
            if any(any(v not in (None, "") for v in row) for row in rows):
                sheets.append((sheet.title, rows))
        wb.close()
        return sheets

    if ext in ("csv", "txt"):
        text = None
        for encoding in ("utf-8-sig", "cp1251", "utf-16"):
            try:
                text = data.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise DataError("Не удалось прочитать кодировку CSV. Сохраните файл в UTF-8.")
        sample = text[:4000]
        try:
            delimiter = csv.Sniffer().sniff(sample, delimiters=";,\t").delimiter
        except csv.Error:
            delimiter = ";" if sample.count(";") >= sample.count(",") else ","
        return [("CSV", [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)])]

    if ext == "xls":
        raise DataError("Старый формат .xls не поддерживается. Сохраните файл как .xlsx или .csv.")
    raise DataError("Поддерживаются файлы .xlsx и .csv.")


def sheet_kind(name: str):
    n = norm(name)
    if "товар" in n or "sku" in n or "прайс" in n:
        return "skus"
    if "комисс" in n:
        return "commissions"
    if "логист" in n:
        return "logistics"
    if "тариф" in n or "прочие" in n:
        return "tariffs"
    if "акци" in n:
        return "promos"
    if "настрой" in n:
        return "settings"
    return None


# ------------------------------------------------------------------ парсеры
def parse_skus(rows):
    header_index, mapping, _ = find_header(rows, ALIASES_SKU, ["name", "current_price_rub"])
    if header_index is None:
        raise DataError(
            "Не нашёл таблицу товаров. Нужны колонки «Наименование» и «Цена». "
            "Скачайте шаблон и заполните лист «Товары»."
        )

    skus, warnings = [], []
    skipped = 0
    for row in rows[header_index + 1:]:
        if not row or all(v in (None, "") for v in row):
            continue
        name = str(cell(row, mapping, "name") or "").strip()
        if not name:
            skipped += 1
            continue
        if norm(name).startswith(("итого", "всего", "итог")):
            skipped += 1
            continue
        price = num(cell(row, mapping, "current_price_rub"))
        if price is None or price <= 0:
            skipped += 1
            continue

        offer_id = cell(row, mapping, "offer_id")
        offer_id = str(offer_id).strip() if offer_id not in (None, "") else f"SKU-{len(skus) + 1}"
        category_id = cell(row, mapping, "category_id")
        category_id = str(category_id).strip() if category_id not in (None, "") else ""
        scheme = str(cell(row, mapping, "scheme") or "FBY").strip().upper()
        if scheme not in ("FBY", "FBS", "DBS"):
            scheme = "FBY"

        sku = {
            "offer_id": offer_id,
            "market_sku": str(cell(row, mapping, "market_sku") or ""),
            "name": name,
            "category_id": category_id,
            "category_name": str(cell(row, mapping, "category_name") or category_id or "Прочее"),
            "brand": "",
            "scheme": scheme,
            "weight_g": num(cell(row, mapping, "weight_g"), 0) or 0,
            "length_mm": num(cell(row, mapping, "length_mm"), 0) or 0,
            "width_mm": num(cell(row, mapping, "width_mm"), 0) or 0,
            "height_mm": num(cell(row, mapping, "height_mm"), 0) or 0,
            "volume_l": num(cell(row, mapping, "volume_l"), 0) or 0,
            "cogs_rub": num(cell(row, mapping, "cogs_rub"), 0) or 0,
            "packaging_rub": num(cell(row, mapping, "packaging_rub"), 0) or 0,
            "delivery_to_warehouse_rub": num(cell(row, mapping, "delivery_to_warehouse_rub"), 0) or 0,
            "storage_days": num(cell(row, mapping, "storage_days"), 45) or 45,
            "return_rate": rate(cell(row, mapping, "return_rate"), 0) or 0,
            "current_price_rub": price,
            "min_price_rub": num(cell(row, mapping, "min_price_rub"), 0) or 0,
            "vat_included": False,
        }

        for key in ("commission_rate", "last_mile_rate", "acquiring_rate", "drr_rate",
                    "tax_rate", "other_rate"):
            value = rate(cell(row, mapping, key)) if key in mapping else None
            if value is not None:
                sku[key] = value
        logistics = num(cell(row, mapping, "logistics_rub")) if "logistics_rub" in mapping else None
        if logistics is not None:
            sku["logistics_rub"] = logistics

        if sku["volume_l"] <= 0 and all(sku[k] > 0 for k in ("length_mm", "width_mm", "height_mm")):
            sku["volume_l"] = round(
                sku["length_mm"] * sku["width_mm"] * sku["height_mm"] / 1_000_000, 3
            )

        skus.append(sku)

    if skipped:
        warnings.append(f"Пропущено строк без названия или цены: {skipped}.")
    if not skus:
        raise DataError("В файле не нашлось ни одного товара с заполненной ценой.")
    return skus, warnings


def parse_commissions(rows, tariffs):
    header_index, mapping, _ = find_header(rows, ALIASES_COMM, ["scheme", "commission_rate"])
    if header_index is None:
        return 0
    count = 0
    for row in rows[header_index + 1:]:
        scheme = str(cell(row, mapping, "scheme") or "").strip().upper()
        if scheme not in tariffs.get("schemes", {}):
            continue
        cat = cell(row, mapping, "category_id")
        value = rate(cell(row, mapping, "commission_rate"))
        if value is None:
            continue
        cat = str(cat).strip() if cat not in (None, "") else "default"
        if cat.lower() in ("default", "прочие", "другое", "по умолчанию"):
            tariffs["schemes"][scheme]["default_commission"] = value
        else:
            tariffs["schemes"][scheme].setdefault("commission_by_category", {})[cat] = value
        count += 1
    return count


def parse_logistics(rows, tariffs):
    header_index, mapping, _ = find_header(rows, ALIASES_LOG, ["scheme", "price_rub"])
    if header_index is None:
        return 0
    tiers: dict[str, list] = {}
    for row in rows[header_index + 1:]:
        scheme = str(cell(row, mapping, "scheme") or "").strip().upper()
        if scheme not in tariffs.get("schemes", {}):
            continue
        price = num(cell(row, mapping, "price_rub"))
        if price is None:
            continue
        tiers.setdefault(scheme, []).append({
            "volume_from_l": num(cell(row, mapping, "volume_from_l"), 0) or 0,
            "volume_to_l": num(cell(row, mapping, "volume_to_l"), None),
            "price_rub": price,
        })
    for scheme, items in tiers.items():
        items.sort(key=lambda item: item["volume_from_l"])
        if items:
            tariffs["schemes"][scheme]["logistics_tiers"] = items
    return sum(len(v) for v in tiers.values())


def set_path(target: dict, path: str, value):
    parts = path.split(".")
    node = target
    for part in parts[:-1]:
        node = node.setdefault(part, {})
    node[parts[-1]] = value


def parse_tariffs(rows, tariffs):
    header_index, mapping, _ = find_header(rows, ALIASES_TAR, ["scheme", "parameter", "value"])
    if header_index is None:
        return 0
    count = 0
    for row in rows[header_index + 1:]:
        scheme = str(cell(row, mapping, "scheme") or "").strip().upper()
        if scheme not in tariffs.get("schemes", {}):
            continue
        path = str(cell(row, mapping, "parameter") or "").strip()
        if not path:
            continue
        raw = cell(row, mapping, "value")
        cfg = tariffs["schemes"][scheme]
        if path == "commission_default":
            value = rate(raw)
            if value is not None:
                cfg["default_commission"] = value
                count += 1
            continue
        if any(word in path for word in ("rate", "commission")):
            value = rate(raw)
        elif "included" in path or "refunded" in path:
            value = flag(raw)
        else:
            value = num(raw)
        if value is None:
            continue
        set_path(cfg, path, value)
        count += 1
    return count


def parse_promos(rows, sku_ids):
    header_index, mapping, _ = find_header(rows, ALIASES_PROMO, ["offer_id"])
    if header_index is None:
        return [], 0
    promos = []
    for row in rows[header_index + 1:]:
        offer_id = str(cell(row, mapping, "offer_id") or "").strip()
        if not offer_id or offer_id not in sku_ids:
            continue
        promo_id = str(cell(row, mapping, "promo_id") or "").strip() or f"PROMO-{len(promos) + 1}"
        promo_price = num(cell(row, mapping, "promo_price_rub"))
        discount = rate(cell(row, mapping, "seller_discount_rate"))
        if promo_price is None and discount is None:
            continue
        promos.append({
            "promo_id": promo_id,
            "promo_name": str(cell(row, mapping, "promo_name") or promo_id),
            "promo_type": str(cell(row, mapping, "promo_type") or ("FIXED_PRICE" if promo_price else "DISCOUNT")).upper(),
            "offer_id": offer_id,
            "start_date": str(cell(row, mapping, "start_date") or ""),
            "end_date": str(cell(row, mapping, "end_date") or ""),
            "seller_discount_rate": None if promo_price is not None else discount,
            "promo_price_rub": promo_price,
            "platform_subsidy_rate": 0,
            "participation_fee_rub": 0,
            "drr_rate": rate(cell(row, mapping, "drr_rate"), 0) or 0,
            "comment": str(cell(row, mapping, "comment") or ""),
        })
    return promos, len(promos)


SETTINGS_MAP = {
    "налоговый режим": ("tax", "regime"),
    "ставка налога": ("tax", "rate"),
    "ставка ндс": ("tax", "vat_rate"),
    "ндс включен в цену": ("tax", "vat_in_price"),
    "целевая маржинальность": ("targets", "margin_rate"),
    "целевой дрр": ("targets", "drr_rate"),
    "минимальная маржа": ("targets", "min_margin_rub"),
    "округление цены": ("rounding", "mode"),
    "шаг округления": ("rounding", "step"),
    "срок хранения по умолчанию": ("defaults", "storage_days"),
    "доля возвратов по умолчанию": ("defaults", "return_rate"),
    "прочие расходы": ("defaults", "other_rate"),
}


def parse_settings(rows, settings):
    header_index, mapping, _ = find_header(rows, ALIASES_SET, ["label", "value"], limit=10)
    if header_index is None:
        header_index = 0
        mapping = {"label": 0, "value": 1}
    count = 0
    for row in rows[header_index + 1:]:
        label = norm(cell(row, mapping, "label"))
        raw = cell(row, mapping, "value")
        if not label or raw in (None, ""):
            continue
        target = SETTINGS_MAP.get(label)
        if not target:
            continue
        section, key = target
        if key == "regime":
            settings.setdefault("tax", {})["regime"] = str(raw).strip().upper()
        elif key in ("margin_rate", "drr_rate", "vat_rate", "return_rate", "other_rate", "rate"):
            value = rate(raw)
            if value is not None:
                settings.setdefault(section, {})[key] = value
                count += 1
        elif key == "vat_in_price":
            settings.setdefault("tax", {})["vat_in_price"] = flag(raw)
            count += 1
        elif key == "mode":
            settings.setdefault("rounding", {})["mode"] = str(raw).strip()
            count += 1
        else:
            value = num(raw)
            if value is not None:
                settings.setdefault(section, {})[key] = value
                count += 1
    return count


# ------------------------------------------------------------------ сборка
def parse_upload(data: bytes, filename: str, current: dict) -> dict:
    if len(data) > 25 * 1024 * 1024:
        raise DataError("Файл слишком большой (больше 25 МБ).")

    sheets = read_sheets(data, filename)
    warnings: list[str] = []
    parsed: dict[str, object] = {}
    used_sheets = []

    for name, rows in sheets:
        kind = sheet_kind(name)
        if not kind or kind in parsed:
            continue
        parsed[kind] = rows
        used_sheets.append(name)

    if "skus" not in parsed:
        # возможно, файл без понятных названий листов — ищем таблицу товаров
        best = None
        for name, rows in sheets:
            try:
                skus, warns = parse_skus(rows)
            except DataError:
                continue
            if best is None or len(skus) > len(best[1]):
                best = (name, skus, warns)
        if best is None:
            raise DataError(
                "В файле не нашлось таблицы товаров. Нужны колонки «Наименование» и «Цена». "
                "Скачайте шаблон — в нём уже есть лист «Товары»."
            )
        parsed["skus"] = sheets[[n for n, _ in sheets].index(best[0])][1]
        used_sheets.append(best[0])

    bundle = {
        "meta": dict(current.get("meta", {})),
        "settings": dict(current.get("settings", {})),
        "tariffs": dict(current.get("tariffs", {})),
        "skus": [],
        "promos": [],
    }
    # глубокая копия через json, чтобы не менять текущий набор
    import json as _json
    bundle = _json.loads(_json.dumps(bundle))

    skus, sku_warnings = parse_skus(parsed["skus"])
    warnings.extend(sku_warnings)
    bundle["skus"] = skus
    sku_ids = {sku["offer_id"] for sku in skus}

    if "commissions" in parsed:
        count = parse_commissions(parsed["commissions"], bundle["tariffs"])
        if count:
            warnings.append(f"Обновлены комиссии: {count} строк.")
    if "logistics" in parsed:
        count = parse_logistics(parsed["logistics"], bundle["tariffs"])
        if count:
            warnings.append(f"Обновлена логистика: {count} ступеней.")
    if "tariffs" in parsed:
        count = parse_tariffs(parsed["tariffs"], bundle["tariffs"])
        if count:
            warnings.append(f"Обновлены тарифы: {count} строк.")
    if "settings" in parsed:
        parse_settings(parsed["settings"], bundle["settings"])

    if "promos" in parsed:
        promos, count = parse_promos(parsed["promos"], sku_ids)
        bundle["promos"] = promos
        if count:
            warnings.append(f"Загружены акции: {count}.")
    else:
        bundle["promos"] = []

    if not any(k in parsed for k in ("commissions", "logistics", "tariffs", "settings")):
        warnings.append("Тарифы взяты из предыдущего набора — в файле не было листов с тарифами.")

    bundle["meta"] = {
        "source": f"Файл: {filename}",
        "version": datetime.now().strftime("%Y.%m.%d %H:%M"),
        "loaded_at": datetime.now().isoformat(timespec="seconds"),
        "file_name": filename,
        "tariffs_source": (
            f"Из файла: {filename}"
            if any(k in parsed for k in ("commissions", "logistics", "tariffs"))
            else bundle.get("meta", {}).get("tariffs_source", "Загруженные тарифы")
        ),
        "warnings": warnings,
        "sheets": used_sheets,
    }
    return bundle
