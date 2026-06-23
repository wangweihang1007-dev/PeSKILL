# PE/私募项目初筛查询工具（不去重版）

本项目是一个本地版 PE/VC/私募项目初筛查询工具。它读取 Excel 历史项目表，将所有非空项目追加导入 SQLite，并按标准化后的“项目简称”做精确或包含匹配查询。

这个版本保留重复项目记录：同一个“项目简称”在同一个 Excel 或重复导入同一个 Excel 时，都会继续写入 SQLite。不使用向量数据库、embedding、RapidFuzz、前端界面或大模型判断。

## 安装

```bash
pip install -r requirements.txt
```

## 数据路径

默认 Excel 路径：

```text
data/raw/初筛项目信息表.xlsx
```

默认 SQLite 路径：

```text
data/database/deals.db
```

会读取以下存在的 Sheet：

- 通过
- 放弃
- 孵化及观察跟踪类

缺失 Sheet 会自动跳过。

## 命令行导入

```bash
python scripts/update_database.py --excel data/raw/初筛项目信息表.xlsx
```

也可以指定数据库：

```bash
python scripts/update_database.py --excel data/raw/初筛项目信息表.xlsx --db data/database/deals.db
```

## Python 函数调用

```python
from src.service import update_deal_database, check_deal_exists, get_deal_detail

update_result = update_deal_database(
    "data/raw/初筛项目信息表.xlsx",
    "data/database/deals.db",
)

query_result = check_deal_exists("某项目简称", "data/database/deals.db")

detail = get_deal_detail(1, "data/database/deals.db")
```

`check_deal_exists` 返回结构：

```python
{
    "query": "某项目简称",
    "normalized_query": "某项目简称",
    "result": "已存在 / 疑似存在 / 未查到",
    "match_type": "exact / contains / none",
    "matches": []
}
```

## 字段

导入字段包括：项目简称、成立时间、城市、申报预期、前轮估值、已投机构、投前估值、本轮投资额、融资截止时间、主营业务、价值说明、收入、利润、项目来源、是否通过、备注/否决原因。

系统同时保留来源追溯字段：`source_file`、`source_sheet`、`source_row`。

## 导入与查询规则

项目记录键：

```text
record_key = normalized_project_short_name
```

`record_key` 只用于查询匹配，不设置唯一约束；导入时不会跳过重复记录。

标准化规则：全角转半角、去首尾空格、去多余空格、英文转小写、去常见标点和括号引号，仅保留中文、英文和数字。

查询先做精确匹配，命中返回“已存在”和 `exact`；否则做字符串包含匹配，命中返回“疑似存在”和 `contains`；都未命中则返回“未查到”和 `none`。

