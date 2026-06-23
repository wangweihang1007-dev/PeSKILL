---
name: query-project
description: Query the local PE project screening index built from the latest 初筛项目信息表 workbook. Use when the user asks whether a project short name has appeared before, whether a PE/VC project passed initial screening, or asks to查重/初筛查询/项目查询.
---

# Query Project

Use the local query tool at:

```powershell
D:\OpenClaw\tools\query_project\query_project.cmd <项目简称>
```

The index file is:

```text
D:\OpenClaw\tools\query_project\project_index.json
```

The current index was rebuilt from:

```text
D:\wechatfile\xwechat_files\wxid_212yrh4z0oft22_bea2\msg\file\2026-06\初筛项目信息表20260621(1).xlsx
```

## Behavior

- Query by project short name.
- Exact matches return the historical screening record.
- For a non-exact two-character query, the tool returns prefix candidates.
- For long misses, ask the user to provide a shorter two-character project abbreviation.

## Output

Return the tool output to the user in Chinese. Preserve the fields shown by the tool:

- 项目简称
- 初筛结果
- 初筛日期
- 城市
- 成立时间
- 估值
- 主营业务
- 来源
- 投资机构

Do not use semantic similarity or LLM-only matching for duplicate judgment unless the user explicitly asks for a separate qualitative opinion.
