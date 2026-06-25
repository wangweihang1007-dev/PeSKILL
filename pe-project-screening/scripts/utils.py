"""Shared database, configuration, normalization, and output helpers."""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import unicodedata
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import yaml


SKILL_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DB_PATH = SKILL_ROOT / "data" / "project_screening.db"
DEFAULT_EXPORT_DIR = SKILL_ROOT / "output" / "exports"
FIELD_MAPPING_PATH = SKILL_ROOT / "config" / "field_mapping.yaml"
STATUS_MAPPING_PATH = SKILL_ROOT / "config" / "status_mapping.yaml"

PROCESS_NOTES = {
    "init": "本次操作为项目初筛数据库初始化，已检查并创建所需数据表，未新增投资判断。",
    "query": "本次操作为项目初筛信息查询，仅返回数据库已有记录，未新增初筛判断，未改变项目当前状态。",
    "insert": "本次操作为新增一条项目初筛记录，已记录初筛时间、项目来源、初筛分类及是否通过字段，未删除或覆盖历史初筛记录。",
    "import": "本次操作为根据最新项目表执行数据库全量更新，旧项目数据已重建，后续查询结果以本次导入数据为准。",
    "export": "本次操作为导出当前数据库内容，仅输出数据库已有记录，未重新判断或改写项目结论。",
    "logs": "本次操作为查看操作日志，仅返回数据库已有日志记录，未改变项目当前状态。",
}


SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS project_main (
    project_id TEXT PRIMARY KEY,
    normalized_project_name TEXT NOT NULL,
    project_short_name TEXT NOT NULL,
    company_name TEXT,
    founded_date TEXT,
    city TEXT,
    listing_expectation TEXT,
    previous_valuation TEXT,
    invested_institutions TEXT,
    pre_money_valuation TEXT,
    post_money_or_investment_amount TEXT,
    financing_deadline TEXT,
    main_business TEXT,
    value_description TEXT,
    last_year_revenue TEXT,
    last_year_profit TEXT,
    project_source TEXT,
    intake_time TEXT,
    source_raw TEXT,
    original_category TEXT,
    current_screening_category TEXT,
    current_pass_status TEXT,
    remark_or_reject_reason TEXT,
    special_tag TEXT,
    source_file TEXT,
    source_sheet TEXT,
    source_row INTEGER,
    raw_data_json TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS screening_records (
    record_id TEXT PRIMARY KEY,
    project_id TEXT NOT NULL,
    screening_time TEXT NOT NULL,
    screening_round TEXT,
    project_source TEXT,
    screening_category TEXT NOT NULL,
    pass_status TEXT NOT NULL,
    screening_description TEXT,
    remark_or_reject_reason TEXT,
    source_type TEXT,
    source_detail TEXT,
    operator TEXT,
    is_current INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL,
    FOREIGN KEY (project_id) REFERENCES project_main(project_id)
);

CREATE TABLE IF NOT EXISTS operation_logs (
    log_id TEXT PRIMARY KEY,
    operation_time TEXT NOT NULL,
    operation_type TEXT NOT NULL,
    user_input TEXT,
    input_params TEXT,
    affected_project_id TEXT,
    affected_project_name TEXT,
    affected_count INTEGER NOT NULL DEFAULT 0,
    modified_database INTEGER NOT NULL DEFAULT 0,
    added_screening_record INTEGER NOT NULL DEFAULT 0,
    exported_file_path TEXT,
    result_status TEXT NOT NULL,
    process_note TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS import_errors (
    error_id TEXT PRIMARY KEY,
    import_time TEXT NOT NULL,
    source_file TEXT,
    sheet_name TEXT,
    row_number INTEGER,
    project_short_name TEXT,
    field_name TEXT,
    raw_value TEXT,
    error_type TEXT NOT NULL,
    error_message TEXT NOT NULL,
    handling_status TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_project_short_name ON project_main(project_short_name);
CREATE INDEX IF NOT EXISTS idx_project_normalized_name ON project_main(normalized_project_name);
CREATE INDEX IF NOT EXISTS idx_project_city ON project_main(city);
CREATE INDEX IF NOT EXISTS idx_project_source ON project_main(project_source);
CREATE INDEX IF NOT EXISTS idx_project_category ON project_main(original_category);
CREATE INDEX IF NOT EXISTS idx_project_pass_status ON project_main(current_pass_status);
CREATE INDEX IF NOT EXISTS idx_screening_project_id ON screening_records(project_id);
CREATE INDEX IF NOT EXISTS idx_screening_time ON screening_records(screening_time);
CREATE INDEX IF NOT EXISTS idx_screening_current ON screening_records(is_current);
CREATE INDEX IF NOT EXISTS idx_logs_operation_time ON operation_logs(operation_time);
CREATE INDEX IF NOT EXISTS idx_logs_operation_type ON operation_logs(operation_type);
"""


def now_iso() -> str:
    """Return a second-precision local timestamp."""
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def resolve_path(path: str | Path | None, default: Path) -> Path:
    """Resolve a user path; Skill-relative paths stay anchored to the Skill root."""
    if path is None:
        return default.resolve()
    candidate = Path(path).expanduser()
    if not candidate.is_absolute():
        candidate = SKILL_ROOT / candidate
    return candidate.resolve()


def connect_db(db_path: str | Path | None = None) -> sqlite3.Connection:
    """Open SQLite with foreign keys and row dictionaries enabled."""
    path = resolve_path(db_path, DEFAULT_DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def ensure_database(db_path: str | Path | None = None) -> Path:
    """Create the database and all tables if absent."""
    path = resolve_path(db_path, DEFAULT_DB_PATH)
    with connect_db(path) as connection:
        connection.executescript(SCHEMA_SQL)
    return path


def load_yaml(path: Path) -> dict[str, Any]:
    """Load a UTF-8 YAML configuration file."""
    with path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def clean_header(value: Any) -> str:
    """Normalize an Excel header for alias matching."""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"[\s\n\r\t:：/、，,（）()]+", "", text).lower()


def safe_text(value: Any) -> str:
    """Convert Excel values to stable text without inventing missing data."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.replace(microsecond=0).isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_project_name(value: Any) -> str:
    """Build a deterministic comparison key from a project short name."""
    text = unicodedata.normalize("NFKC", safe_text(value)).lower()
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", text)


def project_id_for(project_name: str, identity_seed: str = "") -> str:
    """按项目名与来源位置生成稳定 ID；同名项目不合并。"""
    normalized = normalize_project_name(project_name)
    identity = f"{normalized}|{identity_seed}"
    return "prj_" + hashlib.sha256(identity.encode("utf-8")).hexdigest()[:24]


def new_id(prefix: str) -> str:
    """Generate a compact unique identifier."""
    return f"{prefix}_{uuid.uuid4().hex}"


def json_text(value: Any) -> str:
    """Serialize values for audit columns."""
    return json.dumps(value, ensure_ascii=False, default=str, sort_keys=True)


def split_source_and_time(raw_value: Any) -> tuple[str, str]:
    """Split a mixed project-source/intake-time cell when a date is explicit."""
    raw = safe_text(raw_value)
    if not raw:
        return "", ""
    patterns = [
        r"(?P<date>20\d{2}[-/.年]\d{1,2}(?:[-/.月]\d{1,2}日?)?)",
        r"(?P<date>\d{1,2}月\d{1,2}日)",
        r"(?P<date>20\d{2}年\d{1,2}月)",
    ]
    for pattern in patterns:
        match = re.search(pattern, raw)
        if match:
            intake_time = match.group("date")
            source = (raw[: match.start()] + " " + raw[match.end() :]).strip(" -_/，,；;")
            return source.strip(), intake_time
    return raw, ""


def detect_special_tag(cell: Any) -> str:
    """Map workbook row fill colors documented in the source workbook."""
    fill = getattr(cell, "fill", None)
    if fill is None or fill.fill_type != "solid":
        return ""
    color = fill.fgColor
    if color.type == "rgb":
        rgb = str(color.rgb).upper()
        if rgb.endswith("FFFF00"):
            return "老项目更新后再初筛"
    if color.type == "theme":
        theme = color.theme
        tint = color.tint or 0
        if theme == 9 and tint > 0.5:
            return "通过后投研放弃"
        if theme == 0 and tint < 0:
            return "通过但未正式启动"
    return ""


def derive_status(raw_status: Any, sheet_name: str, special_tag: str = "") -> tuple[str, str]:
    """Map mixed source statuses to separate category and pass fields."""
    config = load_yaml(STATUS_MAPPING_PATH)
    raw = safe_text(raw_status)
    if raw:
        normalized = re.sub(r"\s+", "", raw)
        for item in config.get("status_patterns", []):
            if safe_text(item.get("contains")) in normalized:
                return safe_text(item.get("screening_category")), safe_text(item.get("pass_status"))
    override = config.get("special_tag_overrides", {}).get(special_tag)
    if override:
        return safe_text(override.get("screening_category")), safe_text(override.get("pass_status"))
    default = config.get("sheet_defaults", {}).get(sheet_name, {})
    return (
        safe_text(default.get("screening_category")) or "待补充",
        safe_text(default.get("pass_status")) or "待定",
    )


def validate_screening_values(screening_category: str, pass_status: str) -> None:
    """Reject unsupported controlled values before writing a screening record."""
    config = load_yaml(STATUS_MAPPING_PATH)
    categories = set(config.get("allowed_screening_categories", []))
    statuses = set(config.get("allowed_pass_statuses", []))
    if screening_category not in categories:
        raise ValueError(f"初筛分类不合法：{screening_category}；允许值：{sorted(categories)}")
    if pass_status not in statuses:
        raise ValueError(f"是否通过不合法：{pass_status}；允许值：{sorted(statuses)}")


def write_operation_log(
    connection: sqlite3.Connection,
    *,
    operation_type: str,
    user_input: str = "",
    input_params: dict[str, Any] | None = None,
    affected_project_id: str = "",
    affected_project_name: str = "",
    affected_count: int = 0,
    modified_database: bool = False,
    added_screening_record: bool = False,
    exported_file_path: str = "",
    result_status: str = "success",
    process_note: str,
) -> str:
    """Append an operation audit record."""
    log_id = new_id("log")
    connection.execute(
        """
        INSERT INTO operation_logs (
            log_id, operation_time, operation_type, user_input, input_params,
            affected_project_id, affected_project_name, affected_count,
            modified_database, added_screening_record, exported_file_path,
            result_status, process_note
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            log_id,
            now_iso(),
            operation_type,
            user_input,
            json_text(input_params or {}),
            affected_project_id,
            affected_project_name,
            affected_count,
            int(modified_database),
            int(added_screening_record),
            exported_file_path,
            result_status,
            process_note,
        ),
    )
    return log_id


def write_import_error(
    connection: sqlite3.Connection,
    *,
    import_time: str,
    source_file: str,
    sheet_name: str,
    row_number: int | None,
    project_short_name: str,
    field_name: str,
    raw_value: Any,
    error_type: str,
    error_message: str,
    handling_status: str,
) -> str:
    """Append an import exception while retaining the source value."""
    error_id = new_id("err")
    connection.execute(
        """
        INSERT INTO import_errors (
            error_id, import_time, source_file, sheet_name, row_number,
            project_short_name, field_name, raw_value, error_type,
            error_message, handling_status, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            error_id,
            import_time,
            source_file,
            sheet_name,
            row_number,
            project_short_name,
            field_name,
            safe_text(raw_value),
            error_type,
            error_message,
            handling_status,
            now_iso(),
        ),
    )
    return error_id


def rows_to_dicts(rows: Iterable[sqlite3.Row]) -> list[dict[str, Any]]:
    """Convert SQLite rows to JSON-ready dictionaries."""
    return [dict(row) for row in rows]


def print_result(payload: dict[str, Any], process_note: str) -> None:
    """Print machine-readable data followed by the mandatory process note."""
    print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))
    print("\n【过程说明】")
    print(process_note)
