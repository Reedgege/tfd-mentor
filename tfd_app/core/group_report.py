# -*- coding: utf-8 -*-
"""导师版「全组共性问题」聚合。

把每篇论文的体检报告（markdown）里的问题，按 checker 的**固定中文措辞**归到稳定的
问题类别，再跨篇统计「出现篇数 / 占比 / 涉及论文」，供导师组会一句话汇报。

设计要点（守模板驱动铁律）：
- 归类关键词取自 format_checker 各检查项的固定输出措辞（结构性类别），
  **不写死任何学校的格式数值**（如上边距必须 2.5cm、正文必须宋体小四等），
  因此新增学校模板也不会让这里失真。
- 完全独立模块，零侵入 format_checker / engine / xlsx_report，降低改动风险。
"""
import re
from collections import defaultdict

# 归类规则：按优先级（先匹配先归类）。关键词来自 checker 固定措辞。
_CLASSIFY = [
    ("参考文献", "参考文献格式"),
    ("三线表", "三线表规范"),
    ("题注", "图表题注"),
    ("半角标点", "半角标点"),
    ("首行缩进", "正文缩进"),
    ("缩进", "正文缩进"),
    ("页边距", "页边距"),
    ("编号跳号", "章节标题层级"),
    ("缺少第", "章节标题层级"),
    ("章标题", "章节标题层级"),
    ("章节", "章节标题层级"),
    ("样式", "标题样式"),
    ("字体", "字体规范"),
    ("字号", "字号分层"),
    ("对齐", "对齐方式"),
    ("段前", "段间距"),
    ("段后", "段间距"),
    ("封面", "封面要素"),
    ("摘要", "摘要"),
    ("声明", "声明页"),
    ("致谢", "致谢"),
    ("附录", "附录"),
]

# format_checker 输出的问题行形如：- 🔴 **[高]** 缺少第1章标题……
_ISSUE_RE = re.compile(r"\*\*\[(高|中|低|info)\]\*\*\s*(.+?)\s*$")


def classify_issue(msg):
    """把一条问题文本归到稳定类别；未命中返回「其他」。"""
    for kw, cat in _CLASSIFY:
        if kw in msg:
            return cat
    return "其他"


def extract_issues_from_markdown(md):
    """从 format_checker 的 markdown 报告抽出每条问题文本。"""
    out = []
    for line in (md or "").splitlines():
        m = _ISSUE_RE.search(line.strip())
        if m:
            out.append(m.group(2).strip())
    return out


def aggregate(paper_categories, total_papers=None):
    """paper_categories: [(论文标签, [类别, ...]), ...]。
    返回 (summary_lines, columns, rows)，rows 按出现篇数降序。
    """
    total = total_papers if total_papers is not None else len(paper_categories)
    cat_papers = defaultdict(set)
    for label, cats in paper_categories:
        for c in set(cats):
            cat_papers[c].add(label)

    rows = []
    for cat, papers in cat_papers.items():
        n = len(papers)
        pct = ("%d%%" % round(n * 100.0 / total)) if total else "0%"
        rows.append([cat, str(n), pct, "、".join(sorted(papers))])
    rows.sort(key=lambda r: int(r[1]), reverse=True)

    if rows:
        top = "；".join("%s %s篇(%s)" % (r[0], r[1], r[2]) for r in rows[:3])
        summary = [
            "本组共 %d 篇，最突出的共性问题（前 3）：%s。" % (total, top),
            "说明：类别按 checker 固定措辞自动归类，仅供导师把握全组共性、组会汇报参考。",
        ]
    else:
        summary = ["本组 %d 篇均未检出明显共性问题，可放心。" % total]

    columns = [("问题类别", 2400), ("出现篇数", 1000), ("占比", 900), ("涉及论文", 4200)]
    return summary, columns, rows


if __name__ == "__main__":
    # 冒烟自测
    demo = [
        ("张三_论文.docx", ["参考文献格式", "章节标题层级", "正文缩进"]),
        ("李四_论文.docx", ["参考文献格式", "字体规范"]),
        ("王五_论文.docx", ["章节标题层级", "页边距"]),
    ]
    s, c, r = aggregate(demo, 3)
    print("\n".join(s))
    for row in r:
        print(row)
