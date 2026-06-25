---
name: pe-project-screening
description: 基于 SQLite 和 Excel 管理本地 PE 机构项目初筛数据库。适用于 Codex、OpenClaw 或小龙虾需要从多 sheet 项目表初始化或全量重建数据库，按项目简称、城市、来源、分类、状态查询项目，新增且不覆盖历史的初筛记录，导出当前数据库，查看操作日志，或保留导入异常与来源追溯信息等场景。
---

# PE 项目初筛管理

通过本 Skill 的脚本操作本地数据库。所有项目事实和初筛结论必须有来源依据，不得新增投资判断。

## 固定路径

- 数据库：`data/project_screening.db`
- 输入暂存目录：`input/`
- 导出目录：`output/exports/`
- 字段映射：`config/field_mapping.yaml`
- 状态映射：`config/status_mapping.yaml`

所有相对路径均以本 Skill 根目录为基准。

## 执行流程

1. 识别用户要求：初始化、全量导入、查询、新增初筛、导出或查看日志。
2. 执行对应脚本。
3. 原样返回数据库已有结果，不编造缺失项目信息。
4. 每次回复末尾必须保留脚本输出的 `【过程说明】`。

## 常用命令

```powershell
python scripts/init_db.py
python scripts/import_excel.py "input/初筛项目信息表.xlsx"
python scripts/query_project.py --project-name "思澈科技"
python scripts/query_project.py --city "上海" --pass-status "是"
python scripts/insert_screening.py --project-name "思澈科技" --project-source "人工录入" --screening-category "暂缓" --pass-status "待定"
python scripts/insert_screening.py --project-id "prj_xxx" --project-source "人工录入" --screening-category "待补充" --pass-status "待定"
python scripts/export_db.py
```

仅当用户明确要求使用非默认数据库时，才传入 `--db-path`。

## 强制规则

- 全量导入时重建 `project_main`、`screening_records` 和 `import_errors`。
- 全量导入不得删除 `operation_logs`。
- 将每个 Excel sheet 原名保存为 `original_category`。
- 导入阶段不按项目简称去重；每个有效 Excel 行独立入库并生成独立 `project_id`。
- 初筛记录只能新增，不得更新或删除旧记录。
- 项目当前状态取最新一条初筛记录。
- 查询不得改变项目或初筛状态；允许写入查询操作日志。
- 导出只能读取数据库已有内容，不得在导出时重新判断项目。
- 除非用户明确要求，不主动联网。
- 不自行判断项目是否值得投资。
- 对未知、异常或冲突值保留原文，并写入 `import_errors`。

## 参考项目表的专用识别

参考工作簿具有以下结构：

- sheet 为 `通 过`、`放弃`、`孵化及观察跟踪类`；
- 第一行是说明，后续两行是合并表头；
- 存在 `估值（万元）` 和 `上一年度` 两组二级表头；
- 灰色、绿色、黄色行表示特殊项目标记；
- 项目数据中夹有年份分隔行。

使用 `import_excel.py` 自动定位表头、拆分组合表头、识别填充颜色，并跳过非项目分隔行。

## 输出要求

每次返回结构化结果，并至少包含：

- 执行状态；
- 数据库或导出路径；
- 影响数量或匹配数量；
- 相关项目数据或异常；
- `【过程说明】`。

不得将过程说明改写为投资结论。
