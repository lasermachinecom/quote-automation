"""Path configuration for the sales management app."""
from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "sales.db"
ATTACHMENTS_DIR = DATA_DIR / "attachments"
EXPORT_DIR = DATA_DIR / "exports"

for d in (DATA_DIR, ATTACHMENTS_DIR, EXPORT_DIR):
    d.mkdir(parents=True, exist_ok=True)

APP_TITLE = "売上・在庫管理"
APP_VERSION = "0.1.0"
