"""String-based lookup for deal records."""

from __future__ import annotations

from typing import Any

from .database import get_all_deals, get_deals_by_record_key
from .normalizer import normalize_project_short_name


RESULT_EXACT = "已存在"
RESULT_CONTAINS = "疑似存在"
RESULT_NONE = "未查到"


OUTPUT_FIELD_MAP = {
    "project_short_name": "项目简称",
    "founded_time": "成立时间",
    "city": "城市",
    "expected_application": "申报预期",
    "previous_valuation": "前轮估值",
    "invested_institutions": "已投机构",
    "pre_money_valuation": "投前估值",
    "current_round_amount": "本轮投资额",
    "financing_deadline": "融资截止时间",
    "main_business": "主营业务",
    "value_description": "价值说明",
    "revenue": "收入",
    "profit": "利润",
    "deal_source": "项目来源",
    "pass_status": "是否通过",
    "notes_or_rejection_reason": "备注/否决原因",
}


def format_match(row: dict[str, Any]) -> dict[str, Any]:
    item: dict[str, Any] = {"id": row["id"]}
    for source_field, output_field in OUTPUT_FIELD_MAP.items():
        item[output_field] = row.get(source_field) or ""
    item["source_file"] = row.get("source_file") or ""
    item["source_sheet"] = row.get("source_sheet") or ""
    item["source_row"] = row.get("source_row") or 0
    return item


def check_deal(conn: Any, project_short_name: str) -> dict[str, Any]:
    normalized_query = normalize_project_short_name(project_short_name)
    result = {
        "query": project_short_name,
        "normalized_query": normalized_query,
        "result": RESULT_NONE,
        "match_type": "none",
        "matches": [],
    }
    if not normalized_query:
        return result

    exact_matches = get_deals_by_record_key(conn, normalized_query)
    if exact_matches:
        result["result"] = RESULT_EXACT
        result["match_type"] = "exact"
        result["matches"] = [format_match(row) for row in exact_matches]
        return result

    contains_matches = []
    for row in get_all_deals(conn):
        candidate = row.get("normalized_project_short_name") or ""
        if candidate and (normalized_query in candidate or candidate in normalized_query):
            contains_matches.append(row)

    if contains_matches:
        result["result"] = RESULT_CONTAINS
        result["match_type"] = "contains"
        result["matches"] = [format_match(row) for row in contains_matches]

    return result

