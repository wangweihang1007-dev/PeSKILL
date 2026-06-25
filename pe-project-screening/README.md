# pe-project-screening

本项目是一个本地 PE 项目初筛管理 Skill，适用于 OpenClaw / 小龙虾及其他可执行 Python 脚本的 Agent 环境。第一版只使用 SQLite、Excel 和命令行，不接外部数据库，不主动联网。

## 项目结构

```text
pe-project-screening/
├── SKILL.md
├── README.md
├── requirements.txt
├── agents/
│   └── openai.yaml
├── config/
│   ├── field_mapping.yaml
│   └── status_mapping.yaml
├── data/
│   └── .gitkeep
├── input/
│   └── .gitkeep
├── output/
│   └── exports/
│       └── .gitkeep
├── examples/
│   ├── query_examples.md
│   ├── import_examples.md
│   └── output_examples.md
└── scripts/
    ├── init_db.py
    ├── import_excel.py
    ├── query_project.py
    ├── insert_screening.py
    ├── export_db.py
    └── utils.py
```

## 环境

```powershell
python -m pip install -r requirements.txt
```

依赖：

- pandas
- openpyxl
- PyYAML

## 默认路径

- SQLite：`data/project_screening.db`
- 待导入文件：`input/`
- 导出目录：`output/exports/`

所有默认相对路径均以 Skill 根目录为基准，而不是以当前命令行目录为基准。

## 快速使用

```powershell
# 1. 初始化
python scripts/init_db.py

# 2. 以真实项目表全量重建
python scripts/import_excel.py "D:\wechatfile\xwechat_files\wxid_212yrh4z0oft22_bea2\msg\file\2026-06\初筛项目信息表20260621(1).xlsx"

# 3. 查询
python scripts/query_project.py --project-name "思澈科技"
python scripts/query_project.py --city "上海" --category "通 过"
python scripts/query_project.py --source "中芯聚源" --pass-status "是"

# 4. 新增初筛记录
python scripts/insert_screening.py `
  --project-name "思澈科技" `
  --project-source "人工录入" `
  --screening-round "第二次初筛" `
  --screening-category "暂缓" `
  --pass-status "待定" `
  --description "等待补充材料"

# 同名项目超过一个时，先查询，再使用 project_id 精确录入
python scripts/insert_screening.py `
  --project-id "prj_xxxxxxxxxxxxxxxxxxxxxxxx" `
  --project-source "人工录入" `
  --screening-category "待补充" `
  --pass-status "待定"

# 5. 导出当前数据库
python scripts/export_db.py
```

每个脚本均支持：

```powershell
python scripts/query_project.py --help
```

## 针对参考 Excel 的识别方式

参考表包含 `通 过`、`放弃`、`孵化及观察跟踪类` 三个 sheet。每个 sheet 的第 1 行是说明，第 2、3 行是两层表头，数据从第 4 行开始。

导入器会：

1. 在前 10 行自动定位“项目简称”表头；
2. 将“估值（万元）/投前”和“估值（万元）/投后或本轮投资额”拆开；
3. 将“上一年度/收入”和“上一年度/利润”拆开；
4. 将“备注”“备注：”“否决原因”统一映射；
5. 把 sheet 原名保存为 `original_category`；
6. 识别灰色、绿色、黄色行的特殊标记；
7. 跳过只有年份的分隔行；
8. 对空项目名、缺失字段和异常格式写入 `import_errors`。

## 全量更新规则

全量更新会清空：

- `project_main`
- `screening_records`
- `import_errors`

全量更新不会清空：

- `operation_logs`

导入阶段不去重。同一项目简称出现多次时，每个 Excel 有效行都独立进入项目主表，并拥有独立 `project_id`。查询同名项目时返回全部候选。

## 数据库建表 SQL

```sql
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
```

实际脚本还会创建查询索引。完整 SQL 以 `scripts/utils.py` 中的 `SCHEMA_SQL` 为准。

## 当前状态与历史

- Excel 每个有效项目行会生成一条初始 `screening_records`。
- 人工录入初筛时只能 `INSERT` 新记录。
- 旧记录的 `is_current` 会改为 0，但记录本身不会删除或覆盖。
- `project_main.current_screening_category` 和 `current_pass_status` 始终同步到最新记录。

## 边界

- 不主动联网。
- 不补写缺失事实。
- 不自行判断项目是否值得投资。
- 查询不修改项目状态。
- 导出不重新判断项目。
- 每次操作必须输出【过程说明】。
