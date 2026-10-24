"""Generate a deliberately messy Excel workbook of warehouse sales orders.

The data is synthetic. It imitates problems typical of hand-typed monthly
reports: a title row above the header, column names that differ between
sheets, mixed date formats, numbers stored as text with units or currency
signs, duplicates, blank rows and hand-typed "TOTAL" rows.

Usage:
    python src/make_sample.py sample_data/orders_messy.xlsx
"""
from __future__ import annotations

import random
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

PRODUCTS = {
    "FLT-001": ("Oil filter", 4200),
    "FLT-014": ("Air filter", 6800),
    "HYD-220": ("Hydraulic oil 20 L", 24300),
    "BRG-305": ("Bearing 6305", 3100),
    "BLT-080": ("V-belt A-80", 2700),
    "ELC-012": ("Battery 12V 90Ah", 48500),
}
WAREHOUSES = ["Kokshetau", "Astana", "Petropavl"]

# Each sheet uses its own header spelling and column order.
HEADERS = [
    ["Order No", "Date", "SKU", "Product", "Qty", "Price", "Amount", "Warehouse"],
    ["№ заказа", "Дата", "Артикул", "Товар", "Кол-во", "Цена", "Сумма", "Склад"],
    ["order", "shipment date", "sku code", "item", "quantity", "unit price", "total", "wh"],
]


def fmt_date(d: date, style: int) -> object:
    if style == 0:
        return d  # real Excel date
    if style == 1:
        return d.strftime("%d.%m.%Y")  # day-first with dots (KZ/RU style)
    if style == 2:
        return d.strftime("%m/%d/%Y")  # month-first with slashes (US style)
    return d.strftime("%Y-%m-%d")  # ISO


def fmt_money(x: float, style: int) -> object:
    if style == 0:
        return x
    if style == 1:
        return f"{x:,.0f} ₸".replace(",", " ")
    if style == 2:
        return f"{x:,.2f}".replace(",", " ").replace(".", ",") + " тг"
    return str(int(x))


def build(path: Path, seed: int = 7) -> None:
    rnd = random.Random(seed)
    wb = Workbook()
    wb.remove(wb.active)
    order_no = 1000

    for m, month in enumerate([1, 2, 3]):
        ws = wb.create_sheet(["Jan 2025", "февраль", "Mar-25"][m])
        ws.append([f"Sales report — month {month}"])  # title row above header
        ws.append([])
        ws.append(HEADERS[m])
        date_style = [0, 1, 2][m]
        rows = []
        for _ in range(rnd.randint(14, 18)):
            order_no += 1
            sku = rnd.choice(list(PRODUCTS))
            name, price = PRODUCTS[sku]
            qty = rnd.randint(1, 12)
            d = date(2025, month, 1) + timedelta(days=rnd.randint(0, 27))
            sku_txt = rnd.choice([sku, sku.lower(), f" {sku} "])
            qty_txt = rnd.choice([qty, qty, f"{qty} шт"])
            row = [
                order_no,
                fmt_date(d, rnd.choice([date_style, date_style, 3])),
                sku_txt,
                name if rnd.random() > 0.1 else name.upper(),
                qty_txt,
                fmt_money(price, rnd.choice([0, 0, 1, 2])),
                fmt_money(price * qty, rnd.choice([0, 1, 3])),
                rnd.choice(WAREHOUSES + [w.lower() for w in WAREHOUSES]),
            ]
            rows.append(row)
        # Problems a real file would contain.
        rows.insert(4, list(rows[3]))  # exact duplicate
        rows.insert(8, [None] * 8)  # blank row
        bad = list(rows[10])
        bad[6] = "уточнить"  # text instead of an amount
        rows[10] = bad
        wrong = list(rows[12])
        if isinstance(wrong[6], (int, float)):
            wrong[6] = wrong[6] + 5000  # amount does not match qty x price
        else:
            wrong[6] = 999999
        rows[12] = wrong
        for r in rows:
            ws.append(r)
        ws.append(["ИТОГО", None, None, None, None, None, 123456, None])

    ws = wb.create_sheet("notes")
    ws.append(["Remember to send the report to accounting by the 5th"])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "sample_data/orders_messy.xlsx")
    build(out)
    print(f"Saved {out}")
