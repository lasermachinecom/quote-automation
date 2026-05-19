"""Shared pytest fixtures.

Tests run against a temporary SQLite DB so they never touch real data.
We make `sales-management/` importable as a flat module set.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Make modules in sales-management/ importable as top-level (db, config, ...)
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture
def tmp_db(tmp_path):
    """Initialise a fresh DB in tmp_path and return its path."""
    import db as db_mod

    db_path = tmp_path / "sales.db"
    db_mod.init_db(db_path)
    return db_path
