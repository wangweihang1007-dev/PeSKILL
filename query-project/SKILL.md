---
name: query-project
description: Self-contained PE/VC project screening lookup skill. Use when the user asks whether a project short name has appeared before, whether a PE/VC project passed initial screening, or asks to 查重 / 初筛查询 / 项目查询.
---

# Query Project

This is a self-contained OpenClaw skill plus the source code for the local PE/VC project screening lookup tool.

The skill is designed for a local workflow:

1. Put the source files under `D:\OpenClaw\tools\query_project`.
2. Build `project_index.json` from a local Excel workbook with `parse_and_index.py`.
3. Query project short names with `query_project.cmd <项目简称>`.

No project data, Excel workbook, MiniMax key, or generated index is included in this file.

## Runtime Contract

Use the local query command:

```powershell
D:\OpenClaw\tools\query_project\query_project.cmd <项目简称>
```

Expected local index path:

```text
D:\OpenClaw\tools\query_project\project_index.json
```

To rebuild the index from an Excel workbook:

```powershell
$env:PYTHONIOENCODING = 'utf-8'
C:\Users\27851\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe `
  D:\OpenClaw\tools\query_project\parse_and_index.py `
  <初筛项目信息表.xlsx> `
  D:\OpenClaw\tools\query_project\project_index.json
```

## Behavior

- Query by project short name.
- Exact matches return the historical screening record.
- For a non-exact two-character query, the tool returns prefix candidates.
- For long misses, ask the user to provide a shorter two-character project abbreviation.
- Do not use semantic similarity or LLM-only matching for duplicate judgment unless the user explicitly asks for a separate qualitative opinion.

## Output

Return the tool output to the user in Chinese. Preserve these fields when shown by the tool:

- 项目简称
- 初筛结果
- 初筛日期
- 城市
- 成立时间
- 估值
- 主营业务
- 来源
- 投资机构

## Install From This Single File

Create this directory:

```powershell
New-Item -ItemType Directory -Force -Path D:\OpenClaw\tools\query_project
```

Then copy the embedded source blocks below into their corresponding files:

```text
D:\OpenClaw\tools\query_project\query_project.py
D:\OpenClaw\tools\query_project\parse_and_index.py
D:\OpenClaw\tools\query_project\query_project.cmd
```

After copying, rebuild the index from the user's latest Excel workbook. Do not commit or upload `project_index.json` unless the user explicitly approves publishing the underlying project data.

## Embedded Source: query_project.py

```python
"""
项目初筛查询工具 — 供对话中使用

查询逻辑:
1. 精确匹配 → 命中则返回
2. 未命中 + 名称 > 2字 → 提示精简到2字
3. 未命中 + 名称 = 2字 → 前缀匹配 name_2char
   - ≤3条 → 列出候选
   - >3条 → 提示命中数，让用户缩小范围
   - 0条 → 未查到
"""

import json
import os
import sys

INDEX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'project_index.json')


def normalize_name(name):
    s = str(name).strip().replace('\n', '').replace('\r', '')
    s = s.replace(' ', '')
    s = s.replace('（', '(').replace('）', ')')
    return s


def load_index():
    with open(INDEX_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def search_exact(name, projects):
    """精确匹配"""
    name = normalize_name(name)
    return [p for p in projects if p['name'] == name]


def search_prefix_2char(two_char, projects):
    """2字前缀匹配（name_2char字段）"""
    two_char = normalize_name(two_char)
    if len(two_char) != 2:
        return []
    return [p for p in projects if p['name_2char'] == two_char]


def format_result(matches, projects):
    """格式化输出匹配结果"""
    if not matches:
        return None

    lines = []
    for i, m in enumerate(matches, 1):
        lines.append(f"--- 匹配 #{i} ---")
        lines.append(f"  项目简称: {m['name']}")
        lines.append(f"  初筛结果: {m['result']}")
        lines.append(f"  初筛日期: {m['entry_date'] or '未记录'}")
        lines.append(f"  城市:     {m['city'] or '未记录'}")
        lines.append(f"  成立时间: {m['established'] or '未记录'}")
        lines.append(f"  估值:     {m['valuation'] or '未记录'}")
        lines.append(f"  主营业务: {m['business'][:120] if m['business'] else '未记录'}")
        lines.append(f"  来源:     {m['entry_source'] or '未记录'}")
        lines.append(f"  投资机构: {m['investors'][:100] if m['investors'] else '未记录'}")
    return '\n'.join(lines)


def query(name):
    """主查询入口"""
    projects = load_index()['projects']
    name_norm = normalize_name(name)

    # Step 1: 精确匹配
    matches = search_exact(name, projects)
    if matches:
        result = format_result(matches, projects)
        return {
            'status': 'found',
            'count': len(matches),
            'output': f"## 查重结果：共 {len(matches)} 条记录\n\n{result}"
        }

    LQ = '\u201c'  # 左双引号 "
    RQ = '\u201d'  # 右双引号 "

    # Step 2: 名称长度检查
    if len(name_norm) > 2:
        return {
            'status': 'not_found_long',
            'output': f'未找到{LQ}{name}{RQ}的匹配记录。\n\n项目名称长度 > 2字，请将项目名精简到 **2字简称** 后重新查询。'
        }

    # Step 3: 2字前缀匹配
    prefix_matches = search_prefix_2char(name_norm, projects)
    if len(prefix_matches) == 0:
        return {
            'status': 'not_found',
            'output': f'未找到{LQ}{name}{RQ}的匹配记录。\n\n已使用2字简称进行精确匹配和前缀匹配，均无结果。'
        }
    elif len(prefix_matches) <= 3:
        result = format_result(prefix_matches, projects)
        return {
            'status': 'prefix_match',
            'count': len(prefix_matches),
            'output': f'精确匹配未命中，但前缀{LQ}{name_norm}{RQ}匹配到 {len(prefix_matches)} 个候选：\n\n{result}\n\n请确认是否为以上项目。'
        }
    else:
        names = [p['name'] for p in prefix_matches]
        return {
            'status': 'prefix_too_many',
            'count': len(prefix_matches),
            'output': f'前缀{LQ}{name_norm}{RQ}命中 {len(prefix_matches)} 条记录，过多无法逐一列出。\n\n匹配项目: {", ".join(names[:15])}...\n\n请提供更精确的名称。'
        }


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python query.py <项目名称>")
        sys.exit(1)

    name = sys.argv[1]
    result = query(name)
    print(result['output'])
    sys.exit(0 if result['status'] in ('found', 'prefix_match') else 1)

```

## Embedded Source: parse_and_index.py

```python
"""
项目初筛信息表解析与索引构建工具
读取 Excel → 清洗数据 → 输出 JSON 索引文件

用法:
    python parse_and_index.py                          # 使用默认路径
    python parse_and_index.py <excel_path>              # 指定Excel路径
    python parse_and_index.py <excel_path> <output_path> # 指定输出路径
"""

import sys
import json
import re
import os
from datetime import datetime, timedelta
from collections import Counter

import pandas as pd
import numpy as np

# ============================================================
# 日期解析器 — 6级优先级 fallback
# ============================================================

def parse_date(text):
    """从混合文本中提取录入时间，返回 YYYY-MM-DD 字符串或空字符串"""
    if not text or pd.isna(text):
        return ''
    text = str(text).strip()

    # P1: 标准中文日期 2022年11月16日 / 2022年11月 / 2022年
    m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})[日号]', text)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r'(\d{4})年(\d{1,2})月', text)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"

    # P2: 斜杠格式 2022/11/16 / 2022/10
    m = re.search(r'(\d{4})/(\d{1,2})(?:/(\d{1,2}))?', text)
    if m:
        y, mo = m.group(1), m.group(2)
        d = m.group(3) if m.group(3) else '01'
        return f"{y}-{int(mo):02d}-{int(d):02d}"

    # P3: 点分隔格式 2022.11.16 / 2022.9.20
    m = re.search(r'(\d{4})\.(\d{1,2})(?:\.(\d{1,2}))?', text)
    if m:
        y, mo = m.group(1), m.group(2)
        d = m.group(3) if m.group(3) else '01'
        return f"{y}-{int(mo):02d}-{int(d):02d}"

    # P4: 紧凑数字 20221116 / 20220920
    m = re.search(r'(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])', text)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # P5: 短年份+中文 22年11月 / 22年11月16日
    m = re.search(r'(\d{2})年(\d{1,2})月(\d{1,2})[日号]', text)
    if m:
        year = 2000 + int(m.group(1)) if int(m.group(1)) < 50 else 1900 + int(m.group(1))
        return f"{year}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r'(\d{2})年(\d{1,2})月', text)
    if m:
        year = 2000 + int(m.group(1)) if int(m.group(1)) < 50 else 1900 + int(m.group(1))
        return f"{year}-{int(m.group(2)):02d}"

    return ''


def parse_established_date(val):
    """解析成立时间列，处理Excel序列号日期和文本日期"""
    if pd.isna(val):
        return ''
    
    # Excel序列号日期 (如 43525 → 2019-03-11)
    if isinstance(val, (int, float)):
        try:
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=int(val))
            return dt.strftime('%Y-%m-%d')
        except Exception:
            return str(int(val))
    
    val_str = str(val).strip().replace('\n', '')
    if not val_str:
        return ''

    # 中文日期
    m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})[日号]?', val_str)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r'(\d{4})年(\d{1,2})月', val_str)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"

    # 斜杠/点格式
    m = re.search(r'(\d{4})[/.-](\d{1,2})(?:[/.-](\d{1,2}))?', val_str)
    if m:
        y, mo = m.group(1), m.group(2)
        d = m.group(3) if m.group(3) else '01'
        return f"{y}-{int(mo):02d}-{int(d):02d}"

    return val_str[:50]


def extract_entry_source(col14_text):
    """从Col14文本中提取项目来源（去除录入时间后的剩余文本）"""
    if not col14_text:
        return ''
    text = str(col14_text).strip()
    # 移除录入时间部分
    text = re.sub(r'[【\[]?录入时间[：:]*\s*\d{4}[年./-]\d{1,2}[月日./-]*\d{0,2}[日]*(?:[】\]])?', '', text)
    text = re.sub(r'[【\[]?录入时间[：:]*\s*\d{2}年\d{1,2}月\d{0,2}[日号]?(?:[】\]])?', '', text)
    # 清理残留标记
    text = text.replace('【', '').replace('】', '').replace('[', '').replace(']', '').strip()
    # 合并连续空白
    text = re.sub(r'\s+', ' ', text)
    return text[:200]


def normalize_name(name):
    """规范化项目名称：去空格/换行、全角括号统一为半角"""
    if not name:
        return ''
    s = str(name).strip()
    s = s.replace('\n', '').replace('\r', '')
    s = re.sub(r'\s+', '', s)  # 去所有空白
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace('【', '[').replace('】', ']')
    return s


def is_valid_project_name(name):
    """判断是否有效的项目名称（排除数字、空值、过长文本等非项目行）"""
    if not name or pd.isna(name):
        return False
    name_str = str(name).strip()
    if not name_str:
        return False
    # 纯数字 → 非项目
    if name_str.replace('.', '').replace('-', '').isdigit():
        return False
    # 过长文本（>30字）-> 非项目简称
    if len(name_str) > 30:
        return False
    # 排除纯标注行
    skip_keywords = ['补充更新', '标灰色', '序号', 'NaN', 'nan']
    if name_str in skip_keywords:
        return False
    return True


def truncate(text, max_len=300):
    """截断超长文本"""
    if not text or pd.isna(text):
        return ''
    s = str(text).strip()
    return s[:max_len] if len(s) > max_len else s


# ============================================================
# 主解析逻辑
# ============================================================

def parse_excel(excel_path):
    """解析Excel文件，返回项目列表和统计信息"""
    
    sheet_result_map = {
        '通 过': '通过',
        '放弃': '放弃',
        '孵化及观察跟踪类': '孵化及观察跟踪',
    }
    
    sheets_to_parse = list(sheet_result_map.keys())
    
    all_projects = []
    stats = {
        'total_rows_scanned': 0,
        'valid_projects': 0,
        'date_parsed': 0,
        'date_missing': 0,
        'name_len_distribution': {},
        'duplicates': {},
    }
    
    name_counter = Counter()
    
    xls = pd.ExcelFile(excel_path)
    
    for sheet_name in sheets_to_parse:
        if sheet_name not in xls.sheet_names:
            print(f"⚠ Sheet '{sheet_name}' 未找到，跳过")
            continue
        
        result = sheet_result_map.get(sheet_name, sheet_name)
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        for row_idx in range(3, df.shape[0]):
            stats['total_rows_scanned'] += 1
            raw_name = df.iloc[row_idx, 1]
            
            if not is_valid_project_name(raw_name):
                continue
            
            name = normalize_name(raw_name)
            if not name or len(name) < 2:
                continue
            
            # 提取各列
            established_raw = df.iloc[row_idx, 2]
            established = parse_established_date(established_raw)
            
            city = str(df.iloc[row_idx, 3]).strip() if pd.notna(df.iloc[row_idx, 3]) else ''
            if city in ('nan', 'NaN', ''):
                city = ''
            city = city[:50]
            
            ipo_expect = truncate(df.iloc[row_idx, 4], 100) if pd.notna(df.iloc[row_idx, 4]) else ''
            prev_val = truncate(df.iloc[row_idx, 5], 100) if pd.notna(df.iloc[row_idx, 5]) else ''
            investors = truncate(df.iloc[row_idx, 6], 200) if pd.notna(df.iloc[row_idx, 6]) else ''
            valuation = truncate(df.iloc[row_idx, 7], 100) if pd.notna(df.iloc[row_idx, 7]) else ''
            deadline = truncate(df.iloc[row_idx, 9], 100) if pd.notna(df.iloc[row_idx, 9]) else ''
            business = truncate(df.iloc[row_idx, 10], 300) if pd.notna(df.iloc[row_idx, 10]) else ''
            value_desc = truncate(df.iloc[row_idx, 11], 300) if pd.notna(df.iloc[row_idx, 11]) else ''
            last_year = truncate(df.iloc[row_idx, 12], 100) if pd.notna(df.iloc[row_idx, 12]) else ''
            
            # Col14: 来源 + 录入时间
            col14_raw = str(df.iloc[row_idx, 14]) if pd.notna(df.iloc[row_idx, 14]) else ''
            entry_date = parse_date(col14_raw)
            entry_source = extract_entry_source(col14_raw)
            
            # Col15 fallback (备用)
            col15_raw = str(df.iloc[row_idx, 15]) if pd.notna(df.iloc[row_idx, 15]) else ''
            if not entry_date and col15_raw:
                entry_date = parse_date(col15_raw)
            if not entry_source and col15_raw:
                entry_source = extract_entry_source(col15_raw)
            
            project = {
                'name': name,
                'name_2char': name[:2],
                'result': result,
                'city': city,
                'established': established,
                'business': business,
                'entry_date': entry_date,
                'entry_source': entry_source,
                'valuation': valuation,
                'prev_round_val': prev_val,
                'investors': investors,
                'ipo_expect': ipo_expect,
                'deadline': deadline,
                'last_year_finance': last_year,
                'value_desc': value_desc,
                'sheet': sheet_name,
                'row': row_idx + 1,
            }
            
            all_projects.append(project)
            name_counter[name] += 1
            
            if entry_date:
                stats['date_parsed'] += 1
            else:
                stats['date_missing'] += 1
    
    stats['valid_projects'] = len(all_projects)
    stats['duplicates'] = {k: v for k, v in name_counter.items() if v > 1}
    stats['name_len_distribution'] = dict(Counter(len(p['name']) for p in all_projects))
    
    return all_projects, stats


def build_output(all_projects, stats, excel_path, output_path):
    """生成JSON输出文件"""
    
    source_mtime = os.path.getmtime(excel_path)
    source_mtime_str = datetime.fromtimestamp(source_mtime).strftime('%Y-%m-%d %H:%M:%S')
    
    output = {
        'meta': {
            'built_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'source_file': os.path.basename(excel_path),
            'source_file_mtime': source_mtime_str,
            'total_projects': stats['valid_projects'],
            'date_coverage': f"{stats['date_parsed']}/{stats['valid_projects']} ({stats['date_parsed']*100//max(stats['valid_projects'],1)}%)",
            'name_len_distribution': stats['name_len_distribution'],
        },
        'projects': all_projects,
    }
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    return output


def print_stats(stats, output):
    """打印统计信息"""
    print(f"\n{'='*60}")
    print(f"  项目初筛索引构建完成")
    print(f"{'='*60}")
    print(f"  扫描行数:       {stats['total_rows_scanned']}")
    print(f"  有效项目数:     {stats['valid_projects']}")
    print(f"  日期解析成功:   {stats['date_parsed']} ({stats['date_parsed']*100//max(stats['valid_projects'],1)}%)")
    print(f"  日期缺失:       {stats['date_missing']}")
    print(f"  重复项目:       {len(stats['duplicates'])} 个名称出现多次")
    if stats['duplicates']:
        for name, count in sorted(stats['duplicates'].items()):
            print(f"    - {name} ({count}次)")
    print(f"  名称长度分布:   {stats['name_len_distribution']}")
    print(f"  输出文件:       {os.path.abspath(output_path)}")
    
    # 按sheet统计
    sheet_counts = Counter(p['result'] for p in output['projects'])
    print(f"\n  按结果分类:")
    for result, count in sheet_counts.most_common():
        print(f"    {result}: {count}")


if __name__ == '__main__':
    # 默认路径
    default_excel = r'C:\Users\class\Desktop\初筛项目信息表20260524.xlsx'
    default_output = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'project_index.json')
    
    excel_path = sys.argv[1] if len(sys.argv) > 1 else default_excel
    output_path = sys.argv[2] if len(sys.argv) > 2 else default_output
    
    if not os.path.exists(excel_path):
        print(f"❌ 文件不存在: {excel_path}")
        sys.exit(1)
    
    all_projects, stats = parse_excel(excel_path)
    output = build_output(all_projects, stats, excel_path, output_path)
    print_stats(stats, output)
    
    # 抽样检查日期解析
    missing_dates = [p for p in all_projects if not p['entry_date']]
    if missing_dates and len(missing_dates) <= 10:
        print(f"\n  日期缺失项目明细:")
        for p in missing_dates:
            print(f"    [{p['result']}] {p['name']}")
    elif missing_dates:
        print(f"\n  日期缺失项目数: {len(missing_dates)} (已超过10条，不逐一列出)")

```

## Embedded Source: query_project.cmd

```bat
@echo off
setlocal
set "PYTHONIOENCODING=utf-8"
"C:\Users\27851\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "D:\OpenClaw\tools\query_project\query_project.py" %*

```
