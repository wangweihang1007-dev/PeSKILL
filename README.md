# PeSKILL - 冯源资本 PE/VC 数字化投资助手工具集

PeSKILL 是一套专门为 **冯源资本** 投资业务定制的数字化工具集，适用于 AI Agent（如 Codex、OpenClaw 或小龙虾）及命令行环境。该项目旨在通过自动化脚本和结构化规则，提升投资项目初筛、会议纪要整理以及项目表录入的效率与准确性。

## 目录

- [项目结构](#项目结构)
- [环境依赖与配置](#环境依赖与配置)
- [核心工具使用指南](#核心工具使用指南)
  - [1. PE 项目初筛管理 (pe-project-screening)](#1-pe-项目初筛管理-pe-project-screening)
  - [2. 投资会议纪要生成 (investment-meeting-minutes)](#2-投资会议纪要生成-investment-meeting-minutes)
  - [3. 投资项目表录入 (investment-project-intake)](#3-投资项目表录入-investment-project-intake)
- [旧版说明 (OLD)](#旧版说明-old)
- [开发与校验规范](#开发与校验规范)

---

## 项目结构

```text
PeSKILL/
├── README.md                          # 本使用文档
├── .gitignore                         # Git 忽略文件
├── pe-project-screening/              # 【核心1】PE 项目初筛管理 (SQLite 数据库管理)
│   ├── SKILL.md                       # Agent 技能指令
│   ├── README.md                      # 模块说明及快速开始
│   ├── requirements.txt               # 模块依赖
│   ├── config/                        # 字段与状态映射配置
│   ├── data/                          # 本地 SQLite 数据库目录 (project_screening.db)
│   ├── input/                         # 待导入 Excel 放置目录
│   ├── output/                        # 数据库 Excel 导出目录
│   ├── examples/                      # 查询与导入示例
│   └── scripts/                       # 核心操作脚本 (导入、查询、插入初筛、导出)
├── investment-meeting-minutes/        # 【核心2】投资会议纪要生成 (Word文档自动构建)
│   ├── SKILL.md                       # Agent 会议纪要指令
│   ├── assets/                        # 会议纪要 DOCX 模板
│   ├── references/                    # 业务规则、提示词及 JSON Schema
│   └── scripts/                       # 准备、质量检查及文档生成脚本
├── investment-project-intake/         # 【核心3】投资项目表录入 (Excel行追加)
│   ├── SKILL.md                       # Agent 录入行生成指令
│   ├── assets/                        # 四象半导体格式参考及指引 Excel
│   ├── references/                    # 字段定义、口径规范与优先级规则
│   └── scripts/                       # 文本提取与录入校验脚本
└── OLD/                               # 【旧版归档】历史工具 (pe-deal-checker, query-project 等)
```

---

## 环境依赖与配置

### 1. 安装 Python
项目各模块均基于 Python 3.8+ 编写，请确保您的系统中已安装 Python。

### 2. 安装依赖包
在命令行中进入各工具根目录（或全局环境），安装所需的依赖：
```powershell
# 安装项目初筛管理依赖
pip install pandas openpyxl PyYAML

# 若需运行会议纪要生成工具，还需安装 python-docx
pip install python-docx
```

---

## 核心工具使用指南

### 1. PE 项目初筛管理 (`pe-project-screening`)

该模块基于本地 SQLite 数据库与 Excel 联动，实现项目排重、多维度过滤查询、追加初筛记录并导出审计表格。

#### 💡 核心脚本与常用命令：

*   **数据库初始化**
    创建数据表结构和查询索引：
    ```powershell
    python scripts/init_db.py
    ```
*   **全量导入 Excel 项目表**
    将包含 `通 过`、`放弃`、`孵化及观察跟踪类` 工作表的真实 Excel 数据一键解析并导入本地数据库：
    ```powershell
    python scripts/import_excel.py "D:\path\to\初筛项目信息表.xlsx"
    ```
*   **多条件项目查询**
    支持名称、城市、业务、来源、通过状态、特殊标记等复合查询：
    ```powershell
    # 按简称查询
    python scripts/query_project.py --project-name "思澈科技"
    # 按城市和初筛分类组合查询
    python scripts/query_project.py --city "上海" --category "通过"
    # 模糊搜索业务领域
    python scripts/query_project.py --business "MCU"
    ```
*   **新增初筛记录**
    在历史记录中追加新轮次记录，并自动同步更新主表中项目的当前状态（旧历史不会被覆盖）：
    ```powershell
    python scripts/insert_screening.py `
      --project-name "思澈科技" `
      --project-source "人工录入" `
      --screening-round "第二次初筛" `
      --screening-category "暂缓" `
      --pass-status "待定" `
      --description "等待补充财务数据"
    ```
*   **导出当前数据库为 Excel**
    导出一份格式精美、自动排版的审查用 Excel 报表：
    ```powershell
    python scripts/export_db.py
    ```

---

### 2. 投资会议纪要生成 (`investment-meeting-minutes`)

将投资访谈转录草稿及公司 BP/背景材料转化为语法修正后的转录稿、专业的 Q&A 问答以及最终的会议纪要（格式化 Word 文档），并能够自动将 BP 提取的图片匹配插入到纪要的对应章节。

#### 💡 核心执行流：

1.  **准备项目临时目录**
    将原始材料整理并拷贝至临时工作目录：
    ```powershell
    python scripts/prepare_project.py --project "<项目文件夹绝对路径>" --work-dir "<临时工作目录绝对路径>"
    ```
2.  **文字修正与质量校验**
    校准术语与专有名词（例如将“思车”校正为“思澈”），生成 `corrected_transcript.txt`。运行门禁校验：
    ```powershell
    python scripts/quality_check.py corrected --source "<work>/transcript.txt" --corrected "<work>/corrected_transcript.txt"
    ```
3.  **提取 Q&A**
    模型从修正转录稿中自动梳理出逻辑清晰的问答，保存为 `qa.json`。运行校验：
    ```powershell
    python scripts/quality_check.py qa --qa "<work>/qa.json"
    ```
4.  **生成并打包 DOCX 文档**
    读取模板，在输出目录生成并输出 3 个标准的归档 Word 文件：
    ```powershell
    # 验证最终纪要 JSON 结构
    python scripts/quality_check.py minutes --minutes "<work>/minutes.json"
    # 生成最终 Word 文件
    python scripts/build_documents.py --work-dir "<work>" --project "<项目文件夹>" --template "assets/meeting-minutes-template.docx"
    ```
    *最终生成的文件结构如下：*
    ```text
    <项目文件夹>/AI会议纪要输出/<时间戳>/
    ├─ 01_修正转录.docx
    ├─ 02_QA整理.docx
    └─ 03_<项目名称>会议纪要.docx
    ```

---

### 3. 投资项目表录入 (`investment-project-intake`)

自动从零散的项目 BP、纪要或补充材料中抽取结构化事实，遵循“七段式价值法则”，直接追加到项目表 Excel 末尾，确保不臆测、不捏造数字。

#### 💡 核心步骤：

1.  **文本提取**
    解析 PDF/PPTX 等多格式材料：
    ```powershell
    python scripts/extract_reference_text.py <材料路径> --output-dir <临时目录>
    ```
2.  **建立事实台账**：在临时目录产生 `source_facts.json` 记录各项证据来源与时点。
3.  **价值栏起草规范**
    严格按照以下七个部分整理“价值”描述：
    *   `1.团队` / `2.股权结构` / `3.产品` / `4.技术` / `5.生产、客户` / `6.市场` / `7.收入`。
4.  **校验输入草稿**：
    ```powershell
    python scripts/validate_entry.py draft.json
    ```
5.  **写入 Excel 项目表**：调用对应工具或脚本将草稿行格式化追加到您的投资项目总表中。

---

## 旧版说明 (OLD)

在 `OLD/` 文件夹下存放了历史版本的脚本，目前已被最新模块替代或融合：
*   `query-project`：使用本地 `project_index.json` 实现的单文件检索脚本。新版本已被 `pe-project-screening` 的 SQLite 数据库管理程序替代。
*   `pe-deal-checker` 系列：早期的重名/去重校验脚本，目前初筛管理模块在导入与查询阶段已内置更完备的同名项目冲突与匹配逻辑。

---

## 开发与校验规范

1.  **不制造假数据**：项目资料中不存在的信息应保持留空或标注为 `待确认`/`待核实`，坚决不能凭空想象。
2.  **操作日志记录**：所有的查询、插入、导出操作都会向 `operation_logs` 插入审计明细，不可被删除。
3.  **路径基准**：脚本中所有默认的相对路径均以**工具所在的根目录**为基准进行解析，方便跨环境运行。
