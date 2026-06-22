from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


REQUIRED_SECTIONS = [
    "1 公司定位", "2.1 产品", "2.2 市场情况", "2.3 核心客户",
    "3.1 核心技术体系", "3.2 技术差异化优势", "4 财务情况",
    "5.1 历史融资", "5.2 本轮融资安排", "6 发展计划",
]
MARKER_RE = re.compile(r"^(?:\[?\d{1,2}:\d{2}(?::\d{2})?\]?\s*)?(?:发言人\s*\d+|说话人\s*\d+|speaker\s*\d+|主持人|投资方|公司方)\s*[：:]", re.I | re.M)
NUMBER_RE = re.compile(r"\d+(?:\.\d+)?(?:%|％|万元|亿元|元|年|月|日|颗|片|mm|cm|w|v|pa)?", re.I)


def corrected(source: Path, result: Path) -> tuple[list[str], list[str]]:
    src = source.read_text(encoding="utf-8")
    dst = result.read_text(encoding="utf-8")
    errors, warnings = [], []
    src_markers, dst_markers = MARKER_RE.findall(src), MARKER_RE.findall(dst)
    if len(dst_markers) < len(src_markers):
        errors.append(f"发言标记疑似丢失：源 {len(src_markers)}，修正后 {len(dst_markers)}")
    src_nums, dst_nums = set(NUMBER_RE.findall(src)), set(NUMBER_RE.findall(dst))
    missing = sorted(src_nums - dst_nums)
    if missing:
        warnings.append("以下数字未在修正稿中检出，请逐项复核：" + "、".join(missing[:30]))
    ratio = len(dst) / max(len(src), 1)
    if ratio < 0.70:
        errors.append(f"修正稿长度仅为源稿的 {ratio:.1%}，疑似发生总结或删减")
    elif ratio < 0.85:
        warnings.append(f"修正稿长度为源稿的 {ratio:.1%}，请检查是否误删实质发言")
    return errors, warnings


def qa(path: Path) -> tuple[list[str], list[str]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    errors, warnings, seen = [], [], set()
    items = data.get("items")
    if not isinstance(items, list) or not items:
        return ["qa.json 的 items 必须是非空数组"], []
    for i, item in enumerate(items, 1):
        q, a = str(item.get("question", "")).strip(), str(item.get("answer", "")).strip()
        if not q or not a:
            errors.append(f"第 {i} 项缺少问题或回答")
        key = re.sub(r"\W", "", q).lower()
        if key in seen:
            errors.append(f"重复问题：{q}")
        seen.add(key)
        if re.match(r"^\s*\d+[.、)]", q):
            errors.append(f"问题不得编号：{q}")
        if "\n-" in a or "\n1." in a:
            warnings.append(f"第 {i} 项回答疑似被拆成列表，应改为完整段落")
    return errors, warnings


def minutes(path: Path) -> tuple[list[str], list[str]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    errors, warnings = [], []
    sections = data.get("sections", {})
    if list(sections.keys()) != REQUIRED_SECTIONS:
        errors.append("十个固定子章节缺失或顺序不正确")
    for field in ("project_name", "company_name", "meeting_topic", "meeting_purpose", "meeting_date", "participants"):
        if field not in data:
            errors.append(f"缺少字段：{field}")
    dump = json.dumps(data, ensure_ascii=False)
    if "http://" in dump or "https://" in dump:
        warnings.append("检测到网址，请确认未引入外部搜索内容")
    return errors, warnings


def main() -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="mode", required=True)
    p = sub.add_parser("corrected"); p.add_argument("--source", required=True); p.add_argument("--corrected", required=True)
    p = sub.add_parser("qa"); p.add_argument("--qa", required=True)
    p = sub.add_parser("minutes"); p.add_argument("--minutes", required=True)
    args = parser.parse_args()
    if args.mode == "corrected":
        errors, warnings = corrected(Path(args.source), Path(args.corrected))
    elif args.mode == "qa":
        errors, warnings = qa(Path(args.qa))
    else:
        errors, warnings = minutes(Path(args.minutes))
    print(json.dumps({"status": "failed" if errors else "passed", "errors": errors, "warnings": warnings}, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
