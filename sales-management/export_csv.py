"""CSV export helpers."""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

from config import EXPORT_DIR
from db import connect


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def export_sales(rows, dest: Path | None = None) -> Path:
    """rows: iterable of sqlite3.Row from a sales+units join."""
    dest = dest or EXPORT_DIR / f"sales_{_timestamp()}.csv"
    rows = list(rows)
    if not rows:
        with open(dest, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(["(no rows)"])
        return dest
    fields = rows[0].keys()
    with open(dest, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r[k] for k in fields})
    return dest


def export_stock(dest: Path | None = None) -> Path:
    dest = dest or EXPORT_DIR / f"stock_{_timestamp()}.csv"
    with connect() as conn:
        rows = list(conn.execute(
            """
            SELECT u.id, u.serial_no, u.model, u.mfg_date, u.status, u.memo,
                   p.purchase_date, p.vendor_name, p.amount AS purchase_amount
            FROM units u
            LEFT JOIN sales s ON s.unit_id = u.id
            LEFT JOIN purchases p ON p.unit_id = u.id
            WHERE s.id IS NULL AND u.status != '出荷済'
            ORDER BY u.model, u.id
            """
        ))
    return export_sales(rows, dest)


def export_all_units(dest: Path | None = None) -> Path:
    """Full join view — all units with purchase and sale info."""
    dest = dest or EXPORT_DIR / f"all_units_{_timestamp()}.csv"
    with connect() as conn:
        rows = list(conn.execute(
            """
            SELECT
                u.id AS unit_id, u.serial_no, u.model, u.mfg_date, u.status,
                u.memo AS unit_memo,
                p.purchase_date, p.vendor_name, p.vendor_company,
                p.amount AS purchase_amount,
                s.sale_date, s.delivery_date, s.customer_name, s.customer_company,
                s.postal, s.address, s.phone, s.email,
                s.sale_method, s.total_amount, s.payment_status, s.payment_date,
                s.memo AS sale_memo
            FROM units u
            LEFT JOIN purchases p ON p.unit_id = u.id
            LEFT JOIN sales s ON s.unit_id = u.id
            ORDER BY u.id
            """
        ))
    return export_sales(rows, dest)
