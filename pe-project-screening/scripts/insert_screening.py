"""Append a screening record and update the project's current state."""

from __future__ import annotations

import argparse
from typing import Any

from utils import (
    PROCESS_NOTES,
    connect_db,
    ensure_database,
    new_id,
    normalize_project_name,
    now_iso,
    print_result,
    validate_screening_values,
    write_operation_log,
)


def insert_screening_record(
    *,
    project_name: str = "",
    project_id: str = "",
    project_source: str,
    screening_category: str,
    pass_status: str,
    db_path: str | None = None,
    screening_time: str = "",
    screening_round: str = "",
    description: str = "",
    remark: str = "",
    source_type: str = "人工录入",
    source_detail: str = "",
    operator: str = "小龙虾",
) -> dict[str, Any]:
    """Append one immutable screening decision for an existing project."""
    if not project_id.strip() and not project_name.strip():
        raise ValueError("project_id 和项目简称至少提供一个")
    if not project_source.strip():
        raise ValueError("项目来源不能为空")
    validate_screening_values(screening_category, pass_status)

    database_path = ensure_database(db_path)
    normalized_name = normalize_project_name(project_name)
    timestamp = screening_time.strip() or now_iso()
    created_at = now_iso()

    with connect_db(database_path) as connection:
        if project_id:
            matches = connection.execute(
                "SELECT * FROM project_main WHERE project_id = ?",
                (project_id,),
            ).fetchall()
        else:
            exact = connection.execute(
                "SELECT * FROM project_main WHERE normalized_project_name = ?",
                (normalized_name,),
            ).fetchall()
            matches = exact or connection.execute(
                "SELECT * FROM project_main WHERE project_short_name LIKE ? ORDER BY project_short_name",
                (f"%{project_name}%",),
            ).fetchall()
        if not matches:
            raise LookupError(f"未查询到匹配项目：{project_id or project_name}")
        if len(matches) > 1:
            candidates = [
                {
                    "project_id": row["project_id"],
                    "project_short_name": row["project_short_name"],
                    "source_sheet": row["source_sheet"],
                    "source_row": row["source_row"],
                }
                for row in matches
            ]
            raise LookupError(f"匹配到多个同名项目，请使用 project_id 精确录入：{candidates}")

        project = matches[0]
        project_id = project["project_id"]
        record_id = new_id("scr")
        connection.execute("BEGIN")
        connection.execute("UPDATE screening_records SET is_current = 0 WHERE project_id = ?", (project_id,))
        connection.execute(
            """
            INSERT INTO screening_records (
                record_id, project_id, screening_time, screening_round,
                project_source, screening_category, pass_status,
                screening_description, remark_or_reject_reason,
                source_type, source_detail, operator, is_current, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
            """,
            (
                record_id,
                project_id,
                timestamp,
                screening_round,
                project_source,
                screening_category,
                pass_status,
                description,
                remark,
                source_type,
                source_detail,
                operator,
                created_at,
            ),
        )
        connection.execute(
            """
            UPDATE project_main SET
                current_screening_category = ?,
                current_pass_status = ?,
                remark_or_reject_reason = CASE WHEN ? <> '' THEN ? ELSE remark_or_reject_reason END,
                project_source = CASE WHEN ? <> '' THEN ? ELSE project_source END,
                updated_at = ?
            WHERE project_id = ?
            """,
            (
                screening_category,
                pass_status,
                remark,
                remark,
                project_source,
                project_source,
                created_at,
                project_id,
            ),
        )
        write_operation_log(
            connection,
            operation_type="insert_screening_record",
            user_input=project_id or project_name,
            input_params={
                "screening_time": timestamp,
                "screening_round": screening_round,
                "project_source": project_source,
                "screening_category": screening_category,
                "pass_status": pass_status,
                "description": description,
                "remark": remark,
                "source_type": source_type,
                "source_detail": source_detail,
                "operator": operator,
            },
            affected_project_id=project_id,
            affected_project_name=project["project_short_name"],
            affected_count=1,
            modified_database=True,
            added_screening_record=True,
            result_status="success",
            process_note=PROCESS_NOTES["insert"],
        )
        connection.commit()
        history_count = connection.execute(
            "SELECT COUNT(*) AS count FROM screening_records WHERE project_id = ?",
            (project_id,),
        ).fetchone()["count"]

    return {
        "status": "success",
        "database_path": str(database_path),
        "project_id": project_id,
        "project_short_name": project["project_short_name"],
        "record_id": record_id,
        "screening_time": timestamp,
        "screening_round": screening_round,
        "project_source": project_source,
        "screening_category": screening_category,
        "pass_status": pass_status,
        "history_count": history_count,
        "history_appended": True,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="新增一条项目初筛历史记录")
    parser.add_argument("--db-path")
    selector = parser.add_mutually_exclusive_group(required=True)
    selector.add_argument("--project-id", help="查询结果中的项目 ID，适合同名项目精确录入")
    selector.add_argument("--project-name", help="项目简称；同名项目超过一个时需改用 --project-id")
    parser.add_argument("--screening-time")
    parser.add_argument("--screening-round")
    parser.add_argument("--project-source", required=True)
    parser.add_argument("--screening-category", required=True)
    parser.add_argument("--pass-status", required=True)
    parser.add_argument("--description")
    parser.add_argument("--remark")
    parser.add_argument("--source-type", default="人工录入")
    parser.add_argument("--source-detail")
    parser.add_argument("--operator", default="小龙虾")
    args = parser.parse_args()

    result = insert_screening_record(
        db_path=args.db_path,
        project_name=args.project_name or "",
        project_id=args.project_id or "",
        screening_time=args.screening_time or "",
        screening_round=args.screening_round or "",
        project_source=args.project_source,
        screening_category=args.screening_category,
        pass_status=args.pass_status,
        description=args.description or "",
        remark=args.remark or "",
        source_type=args.source_type,
        source_detail=args.source_detail or "",
        operator=args.operator,
    )
    print_result(result, PROCESS_NOTES["insert"])


if __name__ == "__main__":
    main()
