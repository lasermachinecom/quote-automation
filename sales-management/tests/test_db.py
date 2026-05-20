"""Tests for db.py — schema initialisation and read helpers."""
from __future__ import annotations

import sqlite3

import db as db_mod


def test_init_db_creates_expected_tables(tmp_db):
    with sqlite3.connect(tmp_db) as conn:
        rows = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        ).fetchall()
    names = {r[0] for r in rows}
    assert {"units", "purchases", "sales", "orders", "attachments"} <= names


def test_init_db_is_idempotent(tmp_db):
    # Running init_db twice must not raise (CREATE IF NOT EXISTS).
    db_mod.init_db(tmp_db)
    db_mod.init_db(tmp_db)


def test_find_units_by_serial_empty_for_blank(tmp_db):
    assert db_mod.find_units_by_serial("", tmp_db) == []
    assert db_mod.find_units_by_serial("   ", tmp_db) == []


def test_find_units_by_serial_returns_match(tmp_db):
    with db_mod.connect(tmp_db) as conn:
        conn.execute(
            "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
            ("SN-001", "LT6040", "在庫"),
        )

    hits = db_mod.find_units_by_serial("SN-001", tmp_db)
    assert len(hits) == 1
    assert hits[0]["serial_no"] == "SN-001"
    assert hits[0]["model"] == "LT6040"
    assert hits[0]["status"] == "在庫"


def test_find_units_by_serial_includes_sold_units(tmp_db):
    """Duplicate-warning needs to see both in-stock and shipped units."""
    with db_mod.connect(tmp_db) as conn:
        cur = conn.execute(
            "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
            ("SN-DUP", "LT6040", "出荷済"),
        )
        unit_id = cur.lastrowid
        conn.execute(
            "INSERT INTO sales(unit_id, sale_date, customer_name) VALUES (?, ?, ?)",
            (unit_id, "2025-01-15", "客先A"),
        )
        conn.execute(
            "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
            ("SN-DUP", "LT6040", "在庫"),
        )

    hits = db_mod.find_units_by_serial("SN-DUP", tmp_db)
    assert len(hits) == 2
    statuses = {h["status"] for h in hits}
    assert statuses == {"出荷済", "在庫"}


def test_find_units_by_serial_strips_whitespace(tmp_db):
    with db_mod.connect(tmp_db) as conn:
        conn.execute(
            "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
            ("SN-SPACE", "LT6040", "在庫"),
        )
    assert len(db_mod.find_units_by_serial("  SN-SPACE  ", tmp_db)) == 1


def test_stock_summary_counts_only_unsold(tmp_db):
    with db_mod.connect(tmp_db) as conn:
        # 2x LT6040 in stock
        for sn in ("A1", "A2"):
            conn.execute(
                "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
                (sn, "LT6040", "在庫"),
            )
        # 1x LT6040 already shipped (should be excluded)
        cur = conn.execute(
            "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
            ("A3", "LT6040", "出荷済"),
        )
        conn.execute(
            "INSERT INTO sales(unit_id, sale_date) VALUES (?, ?)",
            (cur.lastrowid, "2025-02-01"),
        )
        # 1x FL5500 in stock
        conn.execute(
            "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
            ("B1", "FL5500", "在庫"),
        )

    summary = {row["model"]: row["in_stock"] for row in db_mod.stock_summary(tmp_db)}
    assert summary == {"LT6040": 2, "FL5500": 1}


def test_connect_rolls_back_on_exception(tmp_db):
    """Failed transactions must not leave partial data behind."""
    class Boom(Exception):
        pass

    try:
        with db_mod.connect(tmp_db) as conn:
            conn.execute(
                "INSERT INTO units(serial_no, model, status) VALUES (?, ?, ?)",
                ("ROLLBACK", "X", "在庫"),
            )
            raise Boom()
    except Boom:
        pass

    assert db_mod.find_units_by_serial("ROLLBACK", tmp_db) == []
