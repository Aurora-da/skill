#!/usr/bin/env python3
"""实验报告自检：结构、图表编号、格式、语言、占位。

用法：python3 check_report.py 实验报告.docx
输出 PASS / WARN / INFO 列表，不修改文件。
"""

import re
import sys

from docx import Document
from docx.oxml.ns import qn

REQUIRED = [
    ("实验目的", r"实验目的|目的"),
    ("原理/环境", r"实验原理|原理|实验环境|相关技术"),
    ("步骤/设计", r"实验步骤|步骤|实验内容|实验设计|实现"),
    ("结果", r"实验结果|结果|运行结果"),
    ("分析/讨论", r"分析|讨论"),
    ("结论", r"结论|总结"),
]

CJK_RE = re.compile(r"[\u4e00-\u9fff]")
FIG_RE = re.compile(r"^(图)\s*(\d+)")
TAB_RE = re.compile(r"^(表)\s*(\d+)")
CAPTION_STYLE = "LabCaption"


def style_east_asian(style):
    rpr = style.element.rPr
    if rpr is None or rpr.rFonts is None:
        return None
    return rpr.rFonts.get(qn("w:eastAsia"))


def main():
    if len(sys.argv) < 2:
        sys.exit("用法：python3 check_report.py 实验报告.docx")
    path = sys.argv[1]
    doc = Document(path)
    warns, infos, passes = [], [], []

    paras = doc.paragraphs
    headings = [(p.style.name, p.text.strip()) for p in paras
                if p.style.name.lower().startswith("heading") and p.text.strip()]
    h1 = [t for s, t in headings if s.lower() in ("heading 1", "标题 1")]

    # 1 章节结构
    if not headings:
        warns.append("未发现标题样式段落：正文可能未使用标题 1/2/3")
    found = {name: any(re.search(pat, t) for t in h1) for name, pat in REQUIRED}
    missing = [name for name, ok in found.items() if not ok]
    if missing:
        warns.append("缺少章节：" + "、".join(missing))
    else:
        passes.append("章节结构齐全（%d 个一级标题）" % len(h1))

    # 2 空章节
    empty = []
    for idx, p in enumerate(paras):
        if p.style.name.lower().startswith("heading") and p.text.strip():
            nxt = paras[idx + 1] if idx + 1 < len(paras) else None
            if nxt is not None and (nxt.style.name.lower().startswith("heading") or not nxt.text.strip()):
                empty.append(p.text.strip())
    if empty:
        warns.append("空标题（无内容）：" + "、".join(empty[:4]))

    # 3 图表编号
    fig_nums, tab_nums, fig_caps, tab_caps = [], [], [], []
    for p in paras:
        t = p.text.strip()
        m = FIG_RE.match(t)
        if m:
            fig_nums.append(int(m.group(2)))
            fig_caps.append(t)
        m = TAB_RE.match(t)
        if m:
            tab_nums.append(int(m.group(2)))
            tab_caps.append(t)
    for label, nums in (("图", fig_nums), ("表", tab_nums)):
        if nums and nums != list(range(1, len(nums) + 1)):
            warns.append(f"{label}编号不连续：{nums}")
        elif nums:
            passes.append(f"{label}编号连续（共 {len(nums)} 个）")
    if fig_nums and len(doc.inline_shapes) < len(fig_nums):
        warns.append(f"图题 {len(fig_nums)} 个，实际图片 {len(doc.inline_shapes)} 张，可能缺图")
    if tab_nums and len(doc.tables) < len(tab_nums):
        warns.append(f"表题 {len(tab_nums)} 个，实际表格 {len(doc.tables)} 个，可能缺表")

    # 4 图表正文引用
    body_text = "\n".join(p.text for p in paras
                          if not FIG_RE.match(p.text.strip()) and not TAB_RE.match(p.text.strip()))
    unreferenced = []
    for n in fig_nums:
        if not re.search(rf"图\s*{n}(?![0-9])", body_text):
            unreferenced.append(f"图{n}")
    for n in tab_nums:
        if not re.search(rf"表\s*{n}(?![0-9])", body_text):
            unreferenced.append(f"表{n}")
    if unreferenced:
        warns.append("正文未引用：" + "、".join(unreferenced))
    elif fig_nums or tab_nums:
        passes.append("图表均在正文被引用")

    # 5 正文格式
    normal = doc.styles["Normal"]
    size = normal.font.size.pt if normal.font.size else None
    spacing = normal.paragraph_format.line_spacing
    indent = normal.paragraph_format.first_line_indent
    ea = style_east_asian(normal)
    fmt_issues = []
    if size and abs(size - 12) > 0.5:
        fmt_issues.append(f"正文字号 {size} pt（通用规范 12 pt，模板要求优先）")
    if spacing and abs(float(spacing) - 1.5) > 0.05:
        fmt_issues.append(f"正文行距 {spacing}（通用规范 1.5 倍，模板要求优先）")
    if indent is not None and indent.pt < 20:
        fmt_issues.append("正文首行缩进不足 2 字符")
    if ea and ea not in ("宋体", "SimSun"):
        fmt_issues.append(f"正文中文字体 {ea}（通用规范宋体，模板要求优先）")
    if fmt_issues:
        warns.extend(fmt_issues)
    else:
        passes.append("正文字体、字号、行距、缩进符合规范")

    # 6 标题格式（模板报告可能不同，仅提示）
    h1_style = next((doc.styles[s] for s, _ in headings if s.lower() in ("heading 1", "标题 1")), None)
    if h1_style is not None:
        h1_size = h1_style.font.size.pt if h1_style.font.size else None
        h1_ea = style_east_asian(h1_style)
        if h1_size and abs(h1_size - 14) > 0.5:
            infos.append(f"一级标题 {h1_size} pt（通用规范为四号 14 pt，模板要求优先）")
        if h1_ea and h1_ea not in ("黑体", "SimHei"):
            infos.append(f"一级标题中文字体 {h1_ea}（通用规范为黑体，模板要求优先）")

    # 7 语言
    bad_words = ["我觉得", "我们", "我认为", "非常有趣", "收获很大", "搞定", "跑通", "特别棒"]
    hits = {}
    for p in paras:
        for w in bad_words:
            if w in p.text:
                hits.setdefault(w, p.text.strip()[:30])
    if hits:
        detail = "；".join(f"{w}（{s}…）" for w, s in list(hits.items())[:3])
        warns.append("口语/第一人称：" + detail)
    else:
        passes.append("未发现第一人称与口语表达")

    # 8 占位与待补
    placeholders = [p.text.strip()[:30] for p in paras
                    if re.search(r"TODO|＿＿|待补充|待填|缺图", p.text)]
    if placeholders:
        infos.append(f"存在 {len(placeholders)} 处待填/占位：" + "；".join(placeholders[:3]))

    # 9 参考文献
    refs = [p for p in paras if re.match(r"^\[\d+\]", p.text.strip())]
    if any("参考文献" in t for t in h1) and not refs:
        warns.append("有参考文献章节但未列出条目")
    elif refs:
        passes.append(f"参考文献 {len(refs)} 条")

    # 10 字数
    text = "\n".join(p.text for p in paras)
    char_count = len(CJK_RE.findall(text))
    infos.append(f"中文字数约 {char_count} 字，段落 {len([p for p in paras if p.text.strip()])} 段，图片 {len(doc.inline_shapes)} 张，表格 {len(doc.tables)} 个")

    print(f"检查文件：{path}")
    for item in passes:
        print(f"[PASS] {item}")
    for item in infos:
        print(f"[INFO] {item}")
    for item in warns:
        print(f"[WARN] {item}")
    print(f"小结：{len(warns)} 处需修正，{len(infos)} 条提示。")


if __name__ == "__main__":
    main()
