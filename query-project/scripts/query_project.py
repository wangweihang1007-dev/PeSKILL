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

DEFAULT_INDEX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'project_index.json')
INDEX_PATH = os.environ.get('PROJECT_INDEX_PATH', DEFAULT_INDEX_PATH)


def normalize_name(name):
    s = str(name).strip().replace('\n', '').replace('\r', '')
    s = s.replace(' ', '')
    s = s.replace('（', '(').replace('）', ')')
    return s


def load_index():
    if not os.path.exists(INDEX_PATH):
        raise FileNotFoundError(
            f"索引文件不存在: {INDEX_PATH}\n"
            "请先运行 parse_and_index.py 生成 project_index.json，"
            "或设置 PROJECT_INDEX_PATH 指向已有索引。"
        )
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
