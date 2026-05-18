"""Import existing 売上エクセル into the SQLite database.

Run this once after first launch to migrate historical data.

Usage:
    python import_excel.py "path/to/売上エクセル修正済20260514.xlsx"
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import openpyxl

from config import DB_PATH
from db import connect, init_db


SHEET_NAME = "納品先全リスト"

COL = {
    "goto_memo": 1,
    "company_order": 2,
    "sale_method": 3,
    "year_or_date": 4,
    "delivery_date": 5,
    "visit_date": 6,
    "time": 7,
    "status": 8,
    "postal": 9,
    "name": 10,
    "address": 11,
    "phone": 12,
    "model": 13,
    "mfg_date": 14,
    "serial": 15,
    "remark1": 16,
    "invoice_no": 17,
    "sale_month": 18,
    "freight_flag": 19,
    "freight": 20,
    "total": 21,
    "yahoo_id": 22,
    "email": 23,
    "payment": 24,
    "payment_date": 25,
    "remark2": 26,
    "remark3": 27,
    "remark4": 28,
    "remark5": 29,
    "remark6": 30,
}


def _s(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    return s if s else None


def _i(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip().replace(",", "").replace("円", "")
    try:
        return int(float(s))
    except ValueError:
        return None


def _join_memos(*vals) -> str | None:
    parts = [str(v).strip() for v in vals if v not in (None, "")]
    return "\n".join(parts) if parts else None


def import_excel(xlsx_path: Path) -> tuple[int, int]:
    """Returns (imported_count, skipped_count)."""
    init_db()
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SHEET_NAME]

    imported = 0
    skipped = 0

    with connect(DB_PATH) as conn:
        for r in range(2, ws.max_row + 1):
            def get(key):
                return ws.cell(row=r, column=COL[key]).value

            model = _s(get("model"))
            name = _s(get("name"))
            total = _i(get("total"))

            # Skip empty rows
            if not model and not name and not total:
                skipped += 1
                continue

            # Required: model (we synthesize from name if missing)
            if not model:
                model = "(機種不明)"

            unit_memo = _join_memos(get("remark1"))
            sale_memo = _join_memos(get("goto_memo"), get("remark2"), get("remark3"),
                                    get("remark4"), get("remark5"), get("remark6"))

            # Insert unit
            cur = conn.execute(
                """INSERT INTO units(serial_no, model, mfg_date, status, memo)
                   VALUES (?, ?, ?, '出荷済', ?)""",
                (_s(get("serial")), model, _s(get("mfg_date")), unit_memo),
            )
            unit_id = cur.lastrowid

            # Insert sale
            conn.execute(
                """INSERT INTO sales(
                    unit_id, sale_date, delivery_date,
                    customer_name, customer_company,
                    postal, address, phone, email, yahoo_id,
                    sale_method, invoice_no, sale_month,
                    freight, total_amount,
                    payment_status, payment_date, memo
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    unit_id,
                    _s(get("year_or_date")),
                    _s(get("delivery_date")),
                    name,
                    _s(get("company_order")),
                    _s(get("postal")),
                    _s(get("address")),
                    _s(get("phone")),
                    _s(get("email")),
                    _s(get("yahoo_id")),
                    _s(get("sale_method")),
                    _s(get("invoice_no")),
                    _s(get("sale_month")),
                    _i(get("freight")),
                    total,
                    _s(get("payment")),
                    _s(get("payment_date")),
                    sale_memo,
                ),
            )
            imported += 1

    return imported, skipped


def main():
    if len(sys.argv) < 2:
        print("Usage: python import_excel.py <path-to-xlsx>")
        sys.exit(1)
    xlsx = Path(sys.argv[1])
    if not xlsx.exists():
        print(f"File not found: {xlsx}")
        sys.exit(1)
    imp, skp = import_excel(xlsx)
    print(f"Imported: {imp} rows / Skipped: {skp} rows")
    print(f"Database: {DB_PATH}")


if __name__ == "__main__":
    main()
