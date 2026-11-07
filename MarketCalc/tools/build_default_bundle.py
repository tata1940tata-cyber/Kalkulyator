# -*- coding: utf-8 -*-
"""Сборка демо-набора данных приложения из файлов проекта."""

import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent


def load_json(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def load_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def num(value, default=0.0):
    if value is None or str(value).strip() == "":
        return default
    return float(str(value).replace(",", "."))


def optional_num(value):
    if value is None or str(value).strip() == "":
        return None
    return num(value)


def main():
    skus = []
    for row in load_csv(ROOT / "data/skus.csv"):
        skus.append({
            "offer_id": row["offer_id"],
            "market_sku": row["market_sku"],
            "name": row["name"],
            "category_id": row["category_id"],
            "category_name": row["category_name"],
            "brand": row["brand"],
            "scheme": row["scheme"],
            "weight_g": num(row["weight_g"]),
            "length_mm": num(row["length_mm"]),
            "width_mm": num(row["width_mm"]),
            "height_mm": num(row["height_mm"]),
            "volume_l": num(row["volume_l"]),
            "cogs_rub": num(row["cogs_rub"]),
            "packaging_rub": num(row["packaging_rub"]),
            "delivery_to_warehouse_rub": num(row["delivery_to_warehouse_rub"]),
            "storage_days": num(row["storage_days"]),
            "return_rate": num(row["return_rate"]),
            "current_price_rub": num(row["current_price_rub"]),
            "min_price_rub": num(row["min_price_rub"]),
            "vat_included": str(row.get("vat_included", "false")).lower() == "true",
        })

    promos = []
    for row in load_csv(ROOT / "data/promos.csv"):
        promos.append({
            "promo_id": row["promo_id"],
            "promo_name": row["promo_name"],
            "promo_type": row["promo_type"],
            "offer_id": row["offer_id"],
            "start_date": row["start_date"],
            "end_date": row["end_date"],
            "seller_discount_rate": optional_num(row["seller_discount_rate"]),
            "promo_price_rub": optional_num(row["promo_price_rub"]),
            "platform_subsidy_rate": num(row["platform_subsidy_rate"]),
            "participation_fee_rub": num(row["participation_fee_rub"]),
            "drr_rate": num(row["drr_rate"]),
            "comment": row.get("comment", ""),
        })

    bundle = {
        "meta": {
            "source": "Демо-данные — ноябрь 2026",
            "version": "2026.11",
            "loaded_at": None,
            "file_name": "demo",
            "tariffs_source": "Кабинет → Тарифы (демо)",
            "warnings": [],
        },
        "settings": load_json(ROOT / "data/settings.json"),
        "tariffs": load_json(ROOT / "data/tariffs.json"),
        "skus": skus,
        "promos": promos,
    }

    out = ROOT / "MarketCalc/app/data/default_bundle.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(bundle, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: {len(skus)} SKU, {len(promos)} promos -> {out}")


if __name__ == "__main__":
    main()
