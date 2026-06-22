---
name: pe-deal-checker
description: Build, inspect, or operate a local PE/VC/私募项目初筛查重工具 that imports Excel sheets into SQLite and exposes Python functions for project short-name duplicate checks. Use when the user asks for PE/私募 deal screening deduplication, Excel-to-SQLite append-only imports, project short-name normalization, or agent-callable Python functions such as update_deal_database, check_deal_exists, and get_deal_detail.
---

# PE Deal Checker

## Overview

Create or maintain a minimal local PE/VC/私募 project duplicate checker. Keep the implementation deterministic: Excel input, SQLite storage, normalized project-short-name keys, exact/contains string matching, and structured Python return values.

## Core Contract

Implement the first version only:

- Read Excel workbooks with `pandas` and `openpyxl`.
- Read only existing sheets from: `通过`, `放弃`, `孵化及观察跟踪类`.
- Store records in SQLite at `data/database/deals.db` by default.
- Use `record_key = normalized_project_short_name` and set `record_key TEXT UNIQUE`.
- Import in append-only mode: insert missing `record_key` records and skip existing records.
- Return structured dictionaries from all public functions.
- Do not use vector databases, embeddings, semantic retrieval, RapidFuzz, similarity scores, frontends, MCP, OpenAI APIs, or LLM-based duplicate judgment.

## Recommended Project Shape

Use this structure unless the user requests otherwise:

```text
pe-deal-checker/
├── data/
│   ├── raw/
│   └── database/
├── src/
│   ├── excel_loader.py
│   ├── normalizer.py
│   ├── database.py
│   ├── checker.py
│   └── service.py
├── scripts/
│   └── update_database.py
├── tests/
├── requirements.txt
└── README.md
```

## Data Fields

Extract these Excel columns when present and store missing columns as empty strings:

`项目简称`, `成立时间`, `城市`, `申报预期`, `前轮估值`, `已投机构`, `投前估值`, `本轮投资额`, `融资截止时间`, `主营业务`, `价值说明`, `收入`, `利润`, `项目来源`, `是否通过`, `备注/否决原因`.

Always keep source traceability:

- `source_file`: source workbook path
- `source_sheet`: source sheet name
- `source_row`: original Excel row number, normally DataFrame index + 2 when row 1 is the header

Skip rows whose `项目简称` is empty.

## Normalization

Normalize project short names before import and query:

- Strip leading and trailing whitespace.
- Convert full-width characters to half-width with Unicode NFKC normalization.
- Lowercase English letters.
- Remove whitespace, common punctuation, book-title marks, quotes, and brackets.
- Preserve only Chinese characters, English letters, and numbers.

Preserve the original `项目简称` separately from `normalized_project_short_name`.

## Public Functions

Expose these functions from `src/service.py`:

```python
update_deal_database(excel_path: str, db_path: str = "data/database/deals.db") -> dict
check_deal_exists(project_short_name: str, db_path: str = "data/database/deals.db") -> dict
get_deal_detail(deal_id: int, db_path: str = "data/database/deals.db") -> dict | None
```

`update_deal_database` must return:

```python
{
    "status": "success",
    "mode": "append_only",
    "total_rows": 0,
    "inserted_rows": 0,
    "skipped_empty_name_rows": 0,
    "skipped_duplicate_rows": 0,
    "processed_sheets": [],
    "database_path": ""
}
```

`check_deal_exists` must return:

```python
{
    "query": "",
    "normalized_query": "",
    "result": "已存在 | 疑似存在 | 未查到",
    "match_type": "exact | contains | none",
    "matches": []
}
```

Each match should include Chinese-facing project fields plus `id`, `source_file`, `source_sheet`, and `source_row`.

## Matching Rules

Run exact matching first: normalized query equals `record_key` or `normalized_project_short_name`. Return `已存在` and `exact`.

If no exact match exists, run contains matching: normalized query contains a historical normalized name, or a historical normalized name contains the normalized query. Return `疑似存在` and `contains`.

If no match exists, return `未查到` and `none`.

## Validation

Validate with a small temporary Excel workbook containing:

- one valid project row,
- one empty-name row,
- one duplicate project row,
- at least one exact query,
- at least one contains query,
- one missing query.

Confirm repeat imports do not insert duplicates and that `source_sheet` and `source_row` appear in query results.
