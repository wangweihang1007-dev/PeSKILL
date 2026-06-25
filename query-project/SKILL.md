---
name: query-project
description: Local PE/VC project screening lookup and deduplication skill. Use when the user asks whether a project short name has appeared before, whether a PE/VC project passed initial screening, or asks to 查重 / 初筛查询 / 项目查询 from an 初筛项目信息表 Excel workbook.
---

# Query Project

Use this skill to build and query a local JSON index for PE/VC initial-screening project records.

## Core Workflow

1. Locate the user's latest `初筛项目信息表*.xlsx`.
2. Build or refresh the local index with `scripts/parse_and_index.py`.
3. Query project short names with `scripts/query_project.py` or `scripts/query_project.cmd`.
4. Return the tool output in Chinese, preserving the tool's project fields.

Do not upload or publish the Excel workbook, `project_index.json`, MiniMax keys, or other private project data unless the user explicitly asks.

## Scripts

The reusable code lives in `scripts/`:

- `scripts/parse_and_index.py`: read Excel sheets, clean project rows, and write `project_index.json`.
- `scripts/query_project.py`: query the generated index by project short name.
- `scripts/query_project.cmd`: Windows wrapper that sets UTF-8 output and runs `query_project.py`.

Prefer the bundled Python runtime when available:

```powershell
$py = 'C:\Users\27851\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
```

Fallback to `python` only if the bundled runtime is unavailable and the local environment has `pandas` and `openpyxl`.

## Build Index

From the skill directory:

```powershell
$env:PYTHONIOENCODING = 'utf-8'
& $py .\scripts\parse_and_index.py <初筛项目信息表.xlsx> .\scripts\project_index.json
```

The parser reads these workbook sheets when present:

- `通 过`
- `放弃`
- `孵化及观察跟踪类`

It writes a JSON object with `meta` and `projects`. Keep this generated JSON local unless the user confirms it is safe to publish.

## Query

From the skill directory:

```powershell
.\scripts\query_project.cmd <项目简称>
```

or:

```powershell
$env:PYTHONIOENCODING = 'utf-8'
& $py .\scripts\query_project.py <项目简称>
```

If the index is stored elsewhere, set:

```powershell
$env:PROJECT_INDEX_PATH = '<path-to-project_index.json>'
```

## Matching Behavior

- Exact project-short-name matches return historical records.
- A missed query longer than two characters asks the user to retry with a two-character abbreviation.
- A two-character missed query runs prefix matching against `name_2char`.
- Prefix matches with up to three candidates are listed; larger result sets ask the user to narrow the query.
- Do not invent semantic matches. If the user wants qualitative similarity, label it separately from the deterministic lookup result.

## Response Fields

When returning a match, preserve the tool's fields:

- 项目简称
- 初筛结果
- 初筛日期
- 城市
- 成立时间
- 估值
- 主营业务
- 来源
- 投资机构

## Validation

After rebuilding an index, test at least five known project names from the generated `projects` array. Treat a case as passed when the command returns exit code `0`, includes the queried project name, and includes the expected `初筛结果`.
