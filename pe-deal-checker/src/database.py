"""SQLite persistence for PE/VC deal records."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any


DEAL_COLUMNS = (
    "record_key",
    "project_short_name",
    "normalized_project_short_name",
    "founded_time",
    "city",
    "expected_application",
    "previous_valuation",
    "invested_institutions",
    "pre_money_valuation",
    "current_round_amount",
    "financing_deadline",
    "main_business",
    "value_description",
    "revenue",
    "profit",
    "deal_source",
    "pass_status",
    "notes_or_rejection_reason",
    "source_file",
    "source_sheet",
    "source_row",
)


def connect(db_path: str | Path) -> sqlite3.Connection:
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS deals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_key TEXT NOT NULL UNIQUE,
            project_short_name TEXT NOT NULL,
            normalized_project_short_name TEXT NOT NULL,
            founded_time TEXT,
            city TEXT,
            expected_application TEXT,
            previous_valuation TEXT,
            invested_institutions TEXT,
            pre_money_valuation TEXT,
            current_round_amount TEXT,
            financing_deadline TEXT,
            main_business TEXT,
            value_description TEXT,
            revenue TEXT,
            profit TEXT,
            deal_source TEXT,
            pass_status TEXT,
            notes_or_rejection_reason TEXT,
            source_file TEXT,
            source_sheet TEXT,
            source_row INTEGER,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_deals_normalized_project_short_name "
        "ON deals(normalized_project_short_name)"
    )
    conn.commit()


def record_exists(conn: sqlite3.Connection, record_key: str) -> bool:
    row = conn.execute("SELECT 1 FROM deals WHERE record_key = ? LIMIT 1", (record_key,)).fetchone()
    return row is not None


def insert_deal(conn: sqlite3.Connection, record: dict[str, Any]) -> bool:
    if record_exists(conn, record["record_key"]):
        return False

    now = datetime.now().isoformat(timespec="seconds")
    values = [record.get(column, "") for column in DEAL_COLUMNS]
    columns_sql = ", ".join(DEAL_COLUMNS + ("created_at", "updated_at"))
    placeholders = ", ".join("?" for _ in DEAL_COLUMNS + ("created_at", "updated_at"))
    conn.execute(
        f"INSERT INTO deals ({columns_sql}) VALUES ({placeholders})",
        values + [now, now],
    )
    return True


def get_deal_by_id(conn: sqlite3.Connection, deal_id: int) -> dict[str, Any] | None:
    row = conn.execute("SELECT * FROM deals WHERE id = ?", (deal_id,)).fetchone()
    return dict(row) if row else None


def get_deals_by_record_key(conn: sqlite3.Connection, record_key: str) -> list[dict[str, Any]]:
    rows = conn.execute(
        "SELECT * FROM deals WHERE record_key = ? ORDER BY id",
        (record_key,),
    ).fetchall()
    return [dict(row) for row in rows]


def get_all_deals(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    rows = conn.execute("SELECT * FROM deals ORDER BY id").fetchall()
    return [dict(row) for row in rows]

