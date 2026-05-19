"""Tests for the pure helper functions in import_excel.py.

These don't require openpyxl at import time — but the module does, so we
skip if it's missing on the runner.
"""
from __future__ import annotations

from datetime import datetime

import pytest

pytest.importorskip("openpyxl")

import import_excel as ie  # noqa: E402


# ---------------- _s ----------------

@pytest.mark.parametrize(
    "value, expected",
    [
        (None, None),
        ("", None),
        ("   ", None),
        ("abc", "abc"),
        ("  trim me  ", "trim me"),
        (123, "123"),
        (datetime(2025, 1, 15), "2025-01-15"),
    ],
)
def test_s_normalises_value(value, expected):
    assert ie._s(value) == expected


# ---------------- _i ----------------

@pytest.mark.parametrize(
    "value, expected",
    [
        (None, None),
        ("", None),
        ("abc", None),
        ("1000", 1000),
        ("1,500,000", 1500000),
        ("¥2,000円", 2000),
        (1500.7, 1500),
        (1500, 1500),
    ],
)
def test_i_parses_currency(value, expected):
    assert ie._i(value) == expected


# ---------------- _date_str ----------------

@pytest.mark.parametrize(
    "value, expected",
    [
        (None, None),
        ("", None),
        ("not a date", None),
        ("2024.12.20", "2024-12-20"),
        ("2024/12/20", "2024-12-20"),
        ("2024-12-20", "2024-12-20"),
        ("2024年12月20日", "2024-12-20"),
        ("2024年12月20日 入荷", "2024-12-20"),
        (datetime(2024, 12, 20), "2024-12-20"),
    ],
)
def test_date_str_parses_various_forms(value, expected):
    assert ie._date_str(value) == expected


def test_date_str_rejects_impossible_dates():
    assert ie._date_str("2024/13/40") is None


# ---------------- _is_shipped ----------------

@pytest.mark.parametrize(
    "value, expected",
    [
        (None, False),
        ("", False),
        ("   ", False),
        ("-", False),
        ("－", False),
        ("—", False),
        ("2024-12-20", True),
        (datetime(2024, 12, 20), True),
    ],
)
def test_is_shipped(value, expected):
    assert ie._is_shipped(value) is expected


# ---------------- _join_memo ----------------

def test_join_memo_skips_empty_values():
    out = ie._join_memo(
        ("通番", "1"),
        ("出力", None),
        ("コントローラ", ""),
        ("備考", "中古"),
    )
    assert out == "通番: 1\n備考: 中古"


def test_join_memo_returns_none_when_all_empty():
    assert ie._join_memo(("a", None), ("b", "")) is None
