# -*- coding: utf-8 -*-
"""Референсная реализация калькулятора юнит-экономики Яндекс Маркета.

Читает data/settings.json, data/tariffs.json, data/skus.csv, data/promos.csv
и формирует data/expected_results.csv — эталонные расчёты для проверки
боевой реализации (golden cases).

Сценарии по каждому SKU:
  current        — маржа по текущей цене
  target         — цена под целевую маржинальность и ДРР
  promo:<id>     — маржа по цене входа в акцию
"""

from __future__ import annotations

import csv
import json
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"


# ---------------------------------------------------------------- утилиты
def load_json(name: str) -> dict:
    with open(DATA / name, encoding="utf-8") as fh:
        return json.load(fh)


def load_csv(name: str) -> list[dict]:
    with open(DATA / name, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def num(value, default=0.0) -> float:
    if value is None:
        return default
    text = str(value).strip().replace(",", ".")
    if text == "":
        return default
    return float(text)


def money(value: float) -> float:
    return round(value + 0.0, 2)


def tier_price(tiers: list[dict], volume_l: float) -> float:
    for tier in tiers:
        upper = tier.get("volume_to_l")
        if upper is None or volume_l <= float(upper):
            return float(tier["price_rub"])
    return float(tiers[-1]["price_rub"])


def commission_rate(sku: dict, tariff: dict) -> float:
    return float(
        tariff["commission_by_category"].get(
            str(sku["category_id"]), tariff["default_commission"]
        )
    )


def return_rate(sku: dict, settings: dict) -> float:
    return num(sku.get("return_rate"), settings["defaults"]["return_rate"])


# ---------------------------------------------------------------- расчёт
def unit_costs(sku: dict, tariff: dict, settings: dict,
               price: float, drr_rate: float) -> dict:
    """Полная себестоимость продажи одной единицы по цене price."""
    c_rate = commission_rate(sku, tariff)
    commission = price * c_rate

    logistics = tier_price(tariff["logistics_tiers"], num(sku["volume_l"]))

    lm = tariff["last_mile"]
    last_mile = min(float(lm["max_rub"]), max(float(lm["min_rub"]), price * float(lm["rate"])))

    acquiring = price * float(tariff["acquiring"]["rate"])

    days = num(sku.get("storage_days"), settings["defaults"]["storage_days"])
    storage = float(tariff["storage"]["rate_per_liter_day"]) * num(sku["volume_l"]) * days

    rr = return_rate(sku, settings)
    returns = rr * float(tariff["returns"]["fee_per_return_rub"])
    if not tariff["returns"]["commission_refunded"]:
        returns += rr * commission

    utilization = float(tariff["utilization"]["fee_per_unit_rub"]) * rr
    fines = price * float(tariff["fines"]["rate"])

    delivery = num(sku.get("delivery_to_warehouse_rub"),
                   tariff.get("delivery_to_warehouse_rub", 0))

    ads = price * drr_rate
    other = price * settings["defaults"]["other_rate"]

    cogs = num(sku["cogs_rub"])
    packaging = num(sku["packaging_rub"])

    tax_cfg = settings["tax"]
    revenue = price / (1 + tax_cfg["vat_rate"]) if tax_cfg["vat_in_price"] else price

    costs_without_tax = (
        commission + logistics + last_mile + acquiring + storage + returns
        + utilization + fines + delivery + ads + other + cogs + packaging
    )
    if tax_cfg["regime"] == "USN_PROFIT":
        tax = max(0.0, revenue - costs_without_tax) * tax_cfg["rate"]
    else:  # USN_INCOME, OSNO-упрощение, АУСН «доходы»
        tax = revenue * tax_cfg["rate"]

    total = costs_without_tax + tax
    margin = revenue - total

    return {
        "revenue": revenue,
        "commission": commission,
        "logistics": logistics,
        "last_mile": last_mile,
        "acquiring": acquiring,
        "storage": storage,
        "returns": returns,
        "utilization": utilization,
        "fines": fines,
        "delivery": delivery,
        "drr": ads,
        "tax": tax,
        "other": other,
        "cogs": cogs,
        "packaging": packaging,
        "total": total,
        "margin": margin,
        "margin_rate": margin / revenue if revenue else 0.0,
        "roi": margin / cogs if cogs else 0.0,
    }


def margin_rate_at(sku, tariff, settings, price, drr_rate) -> float:
    return unit_costs(sku, tariff, settings, price, drr_rate)["margin_rate"]


def round_price(value: float, mode: str, step: float = 10) -> float:
    if mode == "none":
        return round(value, 2)
    if mode == "up_1":
        return float(math.ceil(value))
    if mode == "up_10":
        return float(math.ceil(value / 10) * 10)
    if mode == "up_50":
        return float(math.ceil(value / 50) * 50)
    if mode == "psych_99":
        return float(math.floor(value / 100) * 100 + 99)
    return float(math.ceil(value / step) * step)


def solve_target_price(sku, tariff, settings, margin_target, drr_rate):
    """Бисекция по цене: маржинальность монотонно растёт с ценой.

    Возвращает (цена, признак_ограничения) или (None, причина).
    """
    low, high = 0.0, 1_000_000.0
    top_rate = margin_rate_at(sku, tariff, settings, high, drr_rate)
    if top_rate < margin_target:
        return None, "целевая маржинальность недостижима"

    for _ in range(200):
        mid = (low + high) / 2
        if margin_rate_at(sku, tariff, settings, mid, drr_rate) < margin_target:
            low = mid
        else:
            high = mid
        if high - low < 0.01:
            break

    price = round_price(high, settings["rounding"]["mode"], settings["rounding"]["step"])
    limit = num(sku.get("min_price_rub"), 0.0)
    note = ""
    if limit and price < limit:
        price = limit
        note = "цена поднята до минимальной"
    return price, note


def flags_for(costs: dict, sku: dict, settings: dict, target_rate=None) -> str:
    flags = []
    if costs["margin"] < 0:
        flags.append("отрицательная маржа")
    if num(sku.get("min_price_rub"), 0) and costs["revenue"] < num(sku["min_price_rub"]):
        flags.append("ниже минимальной цены")
    if target_rate is not None and costs["margin_rate"] + 1e-9 < target_rate:
        flags.append("ниже целевой маржинальности")
    if costs["margin"] < settings["targets"]["min_margin_rub"]:
        flags.append("маржа ниже минимума")
    return "; ".join(flags) if flags else "ok"


# ---------------------------------------------------------------- сборка
def main() -> None:
    settings = load_json("settings.json")
    tariffs = load_json("tariffs.json")
    skus = load_csv("skus.csv")
    promos = load_csv("promos.csv")

    members = {}
    for promo in promos:
        members.setdefault(promo["offer_id"], []).append(promo)

    columns = [
        "offer_id", "name", "scenario", "price_rub", "revenue_rub",
        "commission_rub", "logistics_rub", "last_mile_rub", "acquiring_rub",
        "storage_rub", "returns_rub", "utilization_rub", "fines_rub",
        "delivery_rub", "drr_rub", "tax_rub", "other_rub", "cogs_rub",
        "packaging_rub", "total_costs_rub", "margin_rub", "margin_rate",
        "roi", "flags",
    ]
    rows = []

    def emit(sku, scenario, price, drr, target_rate=None):
        scheme = sku["scheme"] or settings["default_scheme"]
        tariff = tariffs["schemes"][scheme]
        costs = unit_costs(sku, tariff, settings, price, drr)
        rows.append({
            "offer_id": sku["offer_id"],
            "name": sku["name"],
            "scenario": scenario,
            "price_rub": money(price),
            "revenue_rub": money(costs["revenue"]),
            "commission_rub": money(costs["commission"]),
            "logistics_rub": money(costs["logistics"]),
            "last_mile_rub": money(costs["last_mile"]),
            "acquiring_rub": money(costs["acquiring"]),
            "storage_rub": money(costs["storage"]),
            "returns_rub": money(costs["returns"]),
            "utilization_rub": money(costs["utilization"]),
            "fines_rub": money(costs["fines"]),
            "delivery_rub": money(costs["delivery"]),
            "drr_rub": money(costs["drr"]),
            "tax_rub": money(costs["tax"]),
            "other_rub": money(costs["other"]),
            "cogs_rub": money(costs["cogs"]),
            "packaging_rub": money(costs["packaging"]),
            "total_costs_rub": money(costs["total"]),
            "margin_rub": money(costs["margin"]),
            "margin_rate": round(costs["margin_rate"], 4),
            "roi": round(costs["roi"], 4),
            "flags": flags_for(costs, sku, settings, target_rate),
        })

    for sku in skus:
        scheme = sku["scheme"] or settings["default_scheme"]
        tariff = tariffs["schemes"][scheme]
        drr = settings["targets"]["drr_rate"]
        price = num(sku["current_price_rub"])

        emit(sku, "current", price, drr)

        target_margin = settings["targets"]["margin_rate"]
        new_price, note = solve_target_price(sku, tariff, settings, target_margin, drr)
        if new_price is None:
            rows.append({
                "offer_id": sku["offer_id"], "name": sku["name"],
                "scenario": "target", "flags": note,
            })
        else:
            emit(sku, "target", new_price, drr, target_rate=target_margin)
            if note:
                rows[-1]["flags"] = (rows[-1]["flags"] + "; " + note).replace("ok; ", "")

        for promo in members.get(sku["offer_id"], []):
            promo_price = num(promo.get("promo_price_rub"), 0.0)
            if not promo_price:
                base = num(sku["current_price_rub"])
                promo_price = base * (1 - num(promo.get("seller_discount_rate"), 0.0))
            promo_drr = num(promo.get("drr_rate"), drr)
            emit(sku, f"promo:{promo['promo_id']}", promo_price, promo_drr)

    out = DATA / "expected_results.csv"
    with open(out, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key, "") for key in columns})

    print(f"OK: {len(rows)} rows -> {out.relative_to(ROOT)}")

    print("\nSVODKA (price | costs | margin | rate)")
    for row in rows:
        if "price_rub" in row:
            print(f"{row['offer_id']:<10} {row['scenario']:<22} "
                  f"{row['price_rub']:>9.2f} {row['total_costs_rub']:>9.2f} "
                  f"{row['margin_rub']:>9.2f} {row['margin_rate']*100:>6.1f}%  {row['flags']}")


if __name__ == "__main__":
    main()
