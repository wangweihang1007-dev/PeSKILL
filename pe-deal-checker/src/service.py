"""Agent-callable service functions for the PE/VC deal checker."""

from __future__ import annotations

from contextlib import closing
from pathlib import Path
from typing import Any

from .checker import check_deal
from .database import connect, get_deal_by_id, init_db, insert_deal
from .excel_loader import load_deal_records
from .normalizer import normalize_project_short_name


DEFAULT_DB_PATH = "data/database/deals.db"


def update_deal_database(excel_path: str, db_path: str = DEFAULT_DB_PATH) -> dict[str, Any]:
    """Append new Excel deal records into SQLite using normalized short-name keys."""
    records, load_stats = load_deal_records(excel_path)
    inserted_rows = 0
    skipped_duplicate_rows = 0

    with closing(connect(db_path)) as conn:
        init_db(conn)
        for record in records:
            normalized_name = normalize_project_short_name(record["project_short_name"])
            if not normalized_name:
                continue
            record["normalized_project_short_name"] = normalized_name
            record["record_key"] = normalized_name
            if insert_deal(conn, record):
                inserted_rows += 1
            else:
                skipped_duplicate_rows += 1
        conn.commit()

    return {
        "status": "success",
        "mode": "append_only",
        "total_rows": load_stats["total_rows"],
        "inserted_rows": inserted_rows,
        "skipped_empty_name_rows": load_stats["skipped_empty_name_rows"],
        "skipped_duplicate_rows": skipped_duplicate_rows,
        "processed_sheets": load_stats["processed_sheets"],
        "database_path": str(Path(db_path)),
    }


def check_deal_exists(project_short_name: str, db_path: str = DEFAULT_DB_PATH) -> dict[str, Any]:
    """Check whether a project short name already exists in the deal database."""
    with closing(connect(db_path)) as conn:
        init_db(conn)
        return check_deal(conn, project_short_name)


def get_deal_detail(deal_id: int, db_path: str = DEFAULT_DB_PATH) -> dict[str, Any] | None:
    """Return the complete database row for one deal id."""
    with closing(connect(db_path)) as conn:
        init_db(conn)
        return get_deal_by_id(conn, deal_id)
