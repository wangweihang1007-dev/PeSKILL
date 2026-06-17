# 工作流与数据接口

## 输入发现

递归扫描项目文件夹，但排除 `AI会议纪要输出`、隐藏目录和临时Office文件。

转录候选：`.docx`、`.txt`，文件名优先级为 `初稿` > `原文` > `转录` > `录音修正`。排除含 `会议纪要`、`QA`、`总结`、`投资可行性` 的文件。

背景材料：`.pdf`、`.pptx`。排除文件名含 `会议笔记`、`手写`、`批注`。扫描版PDF若无可提取文字，状态为阻塞并要求用户提供可读版本；第一版不做OCR。

如果最高优先级存在多个候选且无法唯一确定，状态为阻塞。

## 临时工作区

`prepare_project.py`生成：

```text
<work>/
├─ manifest.json
├─ transcript.txt
├─ background.txt
├─ background_index.json
├─ source_facts.json
└─ correction_chunks.json
```

`correction_chunks.json`中的每个对象包含 `index`、`text`、`characters`、`source_start_line`、`source_end_line`。按发言轮次切分，默认每块不超过4500字；单个超长发言按句子边界拆分。

`source_facts.json`由执行 Skill 的 Codex 在首次阅读背景材料后创建，至少包含 `terminology`、`meeting_metadata`、`numeric_facts`、`dated_facts`、`conflicts` 和 `sources`。每个事实记录来源文件和可定位的页码/幻灯片/文本标记；无法定位时不得伪造。

## qa.json

```json
{
  "items": [
    {
      "question": "请介绍公司的核心产品。",
      "answer": "我们的核心产品是……",
      "source_markers": ["发言人1 12:30", "发言人2 12:48"],
      "needs_verification": false
    }
  ]
}
```

`source_markers`可为空，但不得编造时间戳。相邻追问可合并，前提是没有损失问题意图和回答细节。

## minutes.json

```json
{
  "project_name": "公司简称或全称",
  "company_name": "公司全称，未知则待确认",
  "meeting_topic": "会议主题，未知则待确认",
  "meeting_purpose": "了解公司主营业务、融资需求及后续发展规划",
  "meeting_date": "YYYY/MM/DD或待确认",
  "participants": ["冯源资本 张三", "公司 李四"],
  "sections": {
    "1 公司定位": ["段落"],
    "2.1 产品": ["段落"],
    "2.2 市场情况": ["段落"],
    "2.3 核心客户": ["段落"],
    "3.1 核心技术体系": ["段落"],
    "3.2 技术差异化优势": ["段落"],
    "4 财务情况": ["段落"],
    "5.1 历史融资": ["段落"],
    "5.2 本轮融资安排": ["段落"],
    "6 发展计划": ["段落"]
  }
}
```

章节没有信息时写“现有材料未提供相关信息，待确认。”，不要删除章节。

## 输出

`build_documents.py`读取 `corrected_transcript.txt`、`qa.json`、`minutes.json`，生成三个DOCX。最终会议纪要将 `qa.json` 追加为“7 访谈记录”。
