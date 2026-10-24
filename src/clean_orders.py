"""Clean a messy multi-sheet Excel workbook of sales orders into one table.

What it does
------------
1. Reads every sheet, finds the real header row (skips title rows) and maps
   differently spelled column names (EN/RU) to one standard set.
2. Parses dates written in different formats. Dates with dots are treated
   as day-first (KZ/RU convention). For dates with slashes the script looks at
   the whole sheet: if any value has a first part > 12 the sheet is day-first,
   if any has a second part > 12 it is month-first. If the sheet gives no
   evidence, the row is kept but flagged instead of guessed silently.
3. Converts numbers stored as text ("1 250,50 тг", "12 шт") to real numbers.
4. Normalises SKU codes, product names and warehouse names.
5. Removes blank rows, hand-typed TOTAL rows and exact duplicates.
6. Checks the business rule amount = qty x price and flags mismatches.
7. Writes a workbook with three sheets: Clean, Issues, Summary.

Usage:
    python src/clean_orders.py sample_data/orders_messy.xlsx output/orders_clean.xlsx
"""
from __future__ import annotations

import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

STANDARD = ["order_id", "date", "sku", "product", "qty", "unit_price", "amount", "warehouse"]

ALIASES = {
    "order_id": ["order no", "order", "№ заказа", "номер заказа", "order id"],
    "date": ["date", "дата", "shipment date", "дата отгрузки"],
    "sku": ["sku", "артикул", "sku code", "код"],
    "product": ["product", "товар", "item", "наименование"],
    "qty": ["qty", "кол-во", "quantity", "количество"],
    "unit_price": ["price", "цена", "unit price"],
    "amount": ["amount", "сумма", "total", "итого по строке"],
    "warehouse": ["warehouse", "склад", "wh"],
}
_LOOKUP = {alias: std for std, names in ALIASES.items() for alias in names}
TOTAL_WORDS = re.compile(r"^\s*(итого|total|всего)\b", re.IGNORECASE)


def _norm_header(value: object) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    return _LOOKUP.get(str(value).strip().lower())


def find_header_row(raw: pd.DataFrame, min_matches: int = 4) -> int | None:
    """Return the index of the first row that looks like a header."""
    for i, row in raw.iterrows():
        if sum(_norm_header(v) is not None for v in row) >= min_matches:
            return int(i)
    return None


def to_number(value: object) -> float | None:
    """'1 250,50 тг' -> 1250.5, '12 шт' -> 12.0, 'уточнить' -> None."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(" ", " ").strip()
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or not re.search(r"\d", text):
        return None
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


_SLASH = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$")
_DOT = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\s*$")
_ISO = re.compile(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$")


def slash_convention(values: list[object]) -> str | None:
    """Decide 'dmy' or 'mdy' for slash dates on one sheet, or None if unknown."""
    dmy = mdy = False
    for v in values:
        m = _SLASH.match(str(v)) if isinstance(v, str) else None
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2))
        dmy |= a > 12
        mdy |= b > 12
    if dmy and not mdy:
        return "dmy"
    if mdy and not dmy:
        return "mdy"
    return None


def parse_date(value: object, convention: str | None) -> tuple[date | None, str | None]:
    """Return (date, problem). problem is None when the date is certain."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, "missing date"
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date(), None
    if isinstance(value, date):
        return value, None
    if isinstance(value, (int, float)):  # Excel serial number
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))).date(), None
    text = str(value)
    if m := _ISO.match(text):
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3))), None
    if m := _DOT.match(text):
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1))), None
    if m := _SLASH.match(text):
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 12:
            return date(y, b, a), None
        if b > 12:
            return date(y, a, b), None
        if convention == "dmy":
            return date(y, b, a), None
        if convention == "mdy":
            return date(y, a, b), None
        return date(y, a, b), "ambiguous date (day/month), assumed month-first"
    return None, f"unreadable date: {text!r}"


def read_sheet(path: Path, sheet: str) -> pd.DataFrame | None:
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    header_row = find_header_row(raw)
    if header_row is None:
        return None  # not a data sheet (e.g. notes)
    cols = [_norm_header(v) for v in raw.iloc[header_row]]
    body = raw.iloc[header_row + 1 :].copy()
    body.columns = [c or f"_extra_{i}" for i, c in enumerate(cols)]
    body = body[[c for c in STANDARD if c in body.columns]]
    body["source_sheet"] = sheet
    body["source_row"] = body.index + 1  # 1-based Excel row number
    return body


def clean(path: Path) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    sheets = pd.ExcelFile(path).sheet_names
    frames, stats = [], {"sheets_read": [], "sheets_skipped": []}
    for s in sheets:
        df = read_sheet(path, s)
        (stats["sheets_read"] if df is not None else stats["sheets_skipped"]).append(s)
        if df is not None:
            frames.append(df)
    data = pd.concat(frames, ignore_index=True)
    stats["rows_in"] = len(data)

    # Drop blank rows and hand-typed totals.
    value_cols = [c for c in STANDARD if c in data.columns]
    blank = data[value_cols].isna().all(axis=1)
    total = data["order_id"].astype(str).str.match(TOTAL_WORDS)
    stats["blank_rows_removed"] = int(blank.sum())
    stats["total_rows_removed"] = int(total.sum())
    data = data[~blank & ~total].copy()

    issues: list[dict] = []

    def flag(row: pd.Series, problem: str) -> None:
        issues.append(
            {
                "source_sheet": row["source_sheet"],
                "source_row": row["source_row"],
                "order_id": row.get("order_id"),
                "problem": problem,
            }
        )

    # Dates, using a per-sheet convention for slash dates.
    conventions = {
        s: slash_convention(data.loc[data["source_sheet"] == s, "date"].tolist())
        for s in data["source_sheet"].unique()
    }
    parsed = []
    for _, row in data.iterrows():
        d, problem = parse_date(row["date"], conventions[row["source_sheet"]])
        if problem:
            flag(row, problem)
        parsed.append(d)
    data["date"] = pd.to_datetime(parsed)

    # Numbers.
    for col in ["order_id", "qty", "unit_price", "amount"]:
        data[col] = data[col].map(to_number)
    for _, row in data[data["amount"].isna()].iterrows():
        flag(row, "amount is not a number")

    # Text fields.
    data["sku"] = data["sku"].astype(str).str.strip().str.upper()
    # Product name: use the most common spelling for each SKU, preferring
    # spellings that are not typed in ALL CAPS.
    data["product"] = data["product"].astype(str).str.strip()
    canonical = {}
    for sku, names in data.groupby("sku")["product"]:
        normal = names[~names.str.isupper()]
        canonical[sku] = (normal if len(normal) else names).mode().iloc[0]
    data["product"] = data["sku"].map(canonical)
    data["warehouse"] = data["warehouse"].astype(str).str.strip().str.title()
    data["order_id"] = data["order_id"].astype("Int64")

    # Exact duplicates (same values in every business column).
    dup = data.duplicated(subset=STANDARD, keep="first")
    stats["duplicates_removed"] = int(dup.sum())
    data = data[~dup].copy()

    # Business rule: amount = qty x unit_price.
    expected = (data["qty"] * data["unit_price"]).round(2)
    mismatch = data["amount"].notna() & (expected - data["amount"]).abs().gt(0.01)
    for _, row in data[mismatch].iterrows():
        flag(row, f"amount {row['amount']:,.0f} ≠ qty × price {row['qty'] * row['unit_price']:,.0f}")
    data["amount_check"] = "ok"
    data.loc[mismatch, "amount_check"] = "mismatch"
    data.loc[data["amount"].isna(), "amount_check"] = "missing"

    data = data.sort_values(["date", "order_id"]).reset_index(drop=True)
    issues_df = pd.DataFrame(issues, columns=["source_sheet", "source_row", "order_id", "problem"])
    stats["rows_out"] = len(data)
    stats["issues"] = len(issues_df)
    stats["slash_date_convention"] = conventions
    return data, issues_df, stats


def summary(data: pd.DataFrame) -> pd.DataFrame:
    ok = data[data["amount_check"] == "ok"]
    out = (
        ok.assign(month=ok["date"].dt.to_period("M").astype(str))
        .pivot_table(index="month", columns="warehouse", values="amount", aggfunc="sum", fill_value=0)
    )
    out["Total"] = out.sum(axis=1)
    return out.reset_index()


def write(data: pd.DataFrame, issues: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    cols = STANDARD + ["amount_check", "source_sheet", "source_row"]
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="DD.MM.YYYY") as xw:
        data[cols].to_excel(xw, sheet_name="Clean", index=False)
        issues.to_excel(xw, sheet_name="Issues", index=False)
        summary(data).to_excel(xw, sheet_name="Summary", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print(__doc__)
        return 1
    src, dst = Path(argv[1]), Path(argv[2])
    data, issues, stats = clean(src)
    write(data, issues, dst)
    print(f"Read sheets: {', '.join(stats['sheets_read'])} | skipped: {', '.join(stats['sheets_skipped']) or '-'}")
    print(
        f"Rows in: {stats['rows_in']} -> out: {stats['rows_out']} "
        f"(blank {stats['blank_rows_removed']}, totals {stats['total_rows_removed']}, "
        f"duplicates {stats['duplicates_removed']})"
    )
    print(f"Rows flagged for review: {stats['issues']} -> see sheet 'Issues'")
    print(f"Saved {dst}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
