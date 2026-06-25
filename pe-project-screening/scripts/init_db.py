"""Initialize the local PE project-screening SQLite database."""

from __future__ import annotations

import argparse

from utils import PROCESS_NOTES, SCHEMA_SQL, connect_db, ensure_database, print_result, resolve_path, DEFAULT_DB_PATH, write_operation_log


def initialize_database(db_path: str | None = None) -> dict:
    """Create missing tables and record the initialization operation."""
    path = ensure_database(db_path)
    with connect_db(path) as connection:
        tables = [
            row["name"]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
            )
        ]
        write_operation_log(
            connection,
            operation_type="init_database",
            input_params={"db_path": str(path)},
            affected_count=0,
            modified_database=True,
            result_status="success",
            process_note=PROCESS_NOTES["init"],
        )
    return {"status": "success", "database_path": str(path), "tables": tables, "schema_sql": SCHEMA_SQL}


def main() -> None:
    parser = argparse.ArgumentParser(description="初始化 PE 项目初筛数据库")
    parser.add_argument("--db-path", help="SQLite 路径；默认 data/project_screening.db")
    args = parser.parse_args()
    result = initialize_database(args.db_path)
    print_result(result, PROCESS_NOTES["init"])


if __name__ == "__main__":
    main()
