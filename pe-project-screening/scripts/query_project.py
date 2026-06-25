"""Query projects, screening history, or operation logs."""

from __future__ import annotations

import argparse
from typing import Any

from utils import (
    PROCESS_NOTES,
    connect_db,
    ensure_database,
    normalize_project_name,
    print_result,
    rows_to_dicts,
    write_operation_log,
)


def query_projects(
    *,
    db_path: str | None = None,
    project_name: str = "",
    city: str = "",
    source: str = "",
    category: str = "",
    pass_status: str = "",
    business: str = "",
    invested_institution: str = "",
    intake_time: str = "",
    remark: str = "",
    special_tag: str = "",
    limit: int = 50,
) -> dict[str, Any]:
    """Query project_main and attach complete screening history."""
    database_path = ensure_database(db_path)
    conditions: list[str] = []
    values: list[Any] = []
    params = {
        "project_name": project_name,
        "city": city,
        "source": source,
        "category": category,
        "pass_status": pass_status,
        "business": business,
        "invested_institution": invested_institution,
        "intake_time": intake_time,
        "remark": remark,
        "special_tag": special_tag,
        "limit": limit,
    }

    if project_name:
        conditions.append("(normalized_project_name LIKE ? OR project_short_name LIKE ?)")
        values.extend([f"%{normalize_project_name(project_name)}%", f"%{project_name}%"])
    field_filters = {
        "city": city,
        "project_source": source,
        "original_category": category,
        "current_pass_status": pass_status,
        "main_business": business,
        "invested_institutions": invested_institution,
        "intake_time": intake_time,
        "remark_or_reject_reason": remark,
        "special_tag": special_tag,
    }
    for field, value in field_filters.items():
        if value:
            conditions.append(f"{field} LIKE ?")
            values.append(f"%{value}%")

    where_sql = " WHERE " + " AND ".join(conditions) if conditions else ""
    sql = f"SELECT * FROM project_main{where_sql} ORDER BY updated_at DESC, project_short_name LIMIT ?"
    values.append(max(1, min(limit, 500)))

    with connect_db(database_path) as connection:
        projects = rows_to_dicts(connection.execute(sql, values).fetchall())
        for project in projects:
            history = connection.execute(
                """
                SELECT * FROM screening_records
                WHERE project_id = ?
                ORDER BY screening_time ASC, created_at ASC, rowid ASC
                """,
                (project["project_id"],),
            ).fetchall()
            project["screening_history"] = rows_to_dicts(history)

        write_operation_log(
            connection,
            operation_type="query_project",
            user_input=project_name,
            input_params=params,
            affected_count=len(projects),
            modified_database=False,
            result_status="success",
            process_note=PROCESS_NOTES["query"],
        )

    return {
        "status": "success",
        "database_path": str(database_path),
        "query": params,
        "match_count": len(projects),
        "message": "" if projects else "未查询到匹配项目。",
        "projects": projects,
    }


def query_logs(db_path: str | None = None, limit: int = 50) -> dict[str, Any]:
    """Return recent operation logs and append a view-log event."""
    database_path = ensure_database(db_path)
    with connect_db(database_path) as connection:
        logs = rows_to_dicts(
            connection.execute(
                "SELECT * FROM operation_logs ORDER BY operation_time DESC, rowid DESC LIMIT ?",
                (max(1, min(limit, 500)),),
            ).fetchall()
        )
        write_operation_log(
            connection,
            operation_type="view_logs",
            input_params={"limit": limit},
            affected_count=len(logs),
            modified_database=False,
            result_status="success",
            process_note=PROCESS_NOTES["logs"],
        )
    return {"status": "success", "database_path": str(database_path), "log_count": len(logs), "logs": logs}


def main() -> None:
    parser = argparse.ArgumentParser(description="查询项目、初筛历史或操作日志")
    parser.add_argument("--db-path")
    parser.add_argument("--project-name")
    parser.add_argument("--city")
    parser.add_argument("--source")
    parser.add_argument("--category")
    parser.add_argument("--pass-status")
    parser.add_argument("--business")
    parser.add_argument("--invested-institution")
    parser.add_argument("--intake-time")
    parser.add_argument("--remark")
    parser.add_argument("--special-tag")
    parser.add_argument("--show-logs", action="store_true")
    parser.add_argument("--limit", type=int, default=50)
    args = parser.parse_args()

    if args.show_logs:
        result = query_logs(args.db_path, args.limit)
        print_result(result, PROCESS_NOTES["logs"])
        return
    result = query_projects(
        db_path=args.db_path,
        project_name=args.project_name or "",
        city=args.city or "",
        source=args.source or "",
        category=args.category or "",
        pass_status=args.pass_status or "",
        business=args.business or "",
        invested_institution=args.invested_institution or "",
        intake_time=args.intake_time or "",
        remark=args.remark or "",
        special_tag=args.special_tag or "",
        limit=args.limit,
    )
    print_result(result, PROCESS_NOTES["query"])


if __name__ == "__main__":
    main()
