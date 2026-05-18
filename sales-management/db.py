"""SQLite schema and connection helpers for the sales management app."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path

from config import DB_PATH


SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS units (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    serial_no    TEXT,
    model        TEXT NOT NULL,
    mfg_date     TEXT,
    status       TEXT NOT NULL DEFAULT '在庫',
    memo         TEXT,
    created_at   TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at   TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_units_serial ON units(serial_no);
CREATE INDEX IF NOT EXISTS idx_units_model  ON units(model);
CREATE INDEX IF NOT EXISTS idx_units_status ON units(status);

CREATE TABLE IF NOT EXISTS purchases (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    unit_id       INTEGER NOT NULL REFERENCES units(id) ON DELETE CASCADE,
    purchase_date TEXT,
    vendor_name   TEXT,
    vendor_company TEXT,
    amount        INTEGER,
    invoice_no    TEXT,
    memo          TEXT,
    created_at    TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_purchases_unit ON purchases(unit_id);

CREATE TABLE IF NOT EXISTS sales (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    unit_id        INTEGER NOT NULL UNIQUE REFERENCES units(id) ON DELETE CASCADE,
    sale_date      TEXT,
    delivery_date  TEXT,
    customer_name  TEXT,
    customer_company TEXT,
    postal         TEXT,
    address        TEXT,
    phone          TEXT,
    email          TEXT,
    yahoo_id       TEXT,
    sale_method    TEXT,
    invoice_no     TEXT,
    sale_month     TEXT,
    freight        INTEGER,
    total_amount   INTEGER,
    payment_status TEXT,
    payment_date   TEXT,
    memo           TEXT,
    created_at     TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_sales_unit ON sales(unit_id);
CREATE INDEX IF NOT EXISTS idx_sales_date ON sales(sale_date);
CREATE INDEX IF NOT EXISTS idx_sales_customer ON sales(customer_name);

CREATE TABLE IF NOT EXISTS attachments (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    unit_id      INTEGER REFERENCES units(id) ON DELETE CASCADE,
    sale_id      INTEGER REFERENCES sales(id) ON DELETE CASCADE,
    purchase_id  INTEGER REFERENCES purchases(id) ON DELETE CASCADE,
    file_path    TEXT NOT NULL,
    kind         TEXT,
    caption      TEXT,
    created_at   TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_attach_unit ON attachments(unit_id);
"""


def init_db(db_path: Path | str = DB_PATH) -> None:
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.executescript(SCHEMA)


@contextmanager
def connect(db_path: Path | str = DB_PATH):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def find_units_by_serial(serial: str, db_path: Path | str = DB_PATH) -> list[sqlite3.Row]:
    """Return all units (active or sold) matching the serial — for duplicate warning."""
    if not serial or not serial.strip():
        return []
    with connect(db_path) as conn:
        return list(conn.execute(
            """
            SELECT u.*, s.sale_date, s.customer_name
            FROM units u
            LEFT JOIN sales s ON s.unit_id = u.id
            WHERE u.serial_no = ?
            ORDER BY u.id
            """,
            (serial.strip(),),
        ))


def stock_summary(db_path: Path | str = DB_PATH) -> list[sqlite3.Row]:
    """Per-model count of in-stock units (not yet sold)."""
    with connect(db_path) as conn:
        return list(conn.execute(
            """
            SELECT u.model, COUNT(*) AS in_stock
            FROM units u
            LEFT JOIN sales s ON s.unit_id = u.id
            WHERE s.id IS NULL AND u.status != '出荷済'
            GROUP BY u.model
            ORDER BY in_stock DESC, u.model
            """
        ))
