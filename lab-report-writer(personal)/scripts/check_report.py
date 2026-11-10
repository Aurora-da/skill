#!/usr/bin/env python3
"""实验报告自检：结构、图表编号、格式、语言、占位。

用法：python check_report.py 实验报告.docx [--strict]
输出 PASS / WARN / INFO 列表，不修改文件；--strict 时存在 WARN 则退出码为 1。
"""

import argparse
import re
import sys

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

try:  # Windows 控制台默认 GBK，中文提示会乱码
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

REQUIRED = [
    ("实验目的", r"实验目的|目的"),
    ("原理/环境", r"实验原理|原理|实验环境|相关技术"),
    ("步骤/设计", r"实验步骤|步骤|实验内容|实验设计|实现"),
    ("结果", r"实验结果|结果|运行结果"),
    ("分析/讨论", r"分析|讨论"),
    ("结论", r"结论|总结"),
]

CJK_RE = re.compile(r"[\u4e00-\u9fff]")
FIG_RE = re.compile(r"^图\s*(\d+)(?:\s*[-–—.]\s*(\d+))?")
TAB_RE = re.compile(r"^表\s*(\d+)(?:\s*[-–—.]\s*(\d+))?")
CAPTION_STYLE = "LabCaption"
REF_NUM_RE = re.compile(r"^\[(\d+)\]")
REF_TYPE_RE = re.compile(r"\[[A-Z]{1,2}(?:/(?:OL|EB))?\]")
CITE_RE = re.compile(r"\[(\d+(?:\s*[-,，]\s*\d+)*)\]")
DRAWING_XPATH = ".//w:drawing | .//w:pict"


def style_east_asian(style):
    rpr = style.element.rPr
    if rpr is None or rpr.rFonts is None:
        return None
    return rpr.rFonts.get(qn("w:eastAsia"))


def iter_block_items(doc):
    """按文档顺序产出段落与表格（doc.paragraphs 不含表格）。"""
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def is_heading(block):
    return isinstance(block, Paragraph) and block.style.name.lower().startswith("heading")


def has_visible_content(block):
    if isinstance(block, Table):
        return True
    return bool(block.text.strip()) or bool(block._p.xpath(DRAWING_XPATH))


def empty_headings(doc):
    """标题后直到下一个标题之间没有任何文字/图片/表格，才算空标题。"""
    blocks = list(iter_block_items(doc))
    empty = []
    for idx, block in enumerate(blocks):
        if not is_heading(block) or not block.text.strip():
            continue
        filled = False
        for nxt in blocks[idx + 1:]:
            if is_heading(nxt):
                break
            if has_visible_content(nxt):
                filled = True
                break
        if not filled:
            empty.append(block.text.strip()[:24])
    return empty


def parse_caption(text, regex):
    m = regex.match(text.strip())
    if not m:
        return None
    if m.group(2) is not None:
        return (int(m.group(1)), int(m.group(2)))
    return (int(m.group(1)), None)


def is_caption_paragraph(p):
    """题注样式，或短的“图1 …/表1 …”段落；正文句子（以句号结尾、过长）不算。"""
    t = p.text.strip()
    if p.style.name != CAPTION_STYLE and (len(t) > 40 or re.search(r"[。！？；]$", t)):
        return False
    return bool(FIG_RE.match(t) or TAB_RE.match(t))


def check_numbering(label, items):
    """items 为 (章号, 序号) 或 (编号, None)。返回问题描述或 None。"""
    if not items:
        return None
    if all(seq is None for _, seq in items):
        nums = [ch for ch, _ in items]
        if nums != list(range(1, len(nums) + 1)):
            return f"{label}编号不连续：{nums}"
        return None
    if any(seq is None for _, seq in items):
        return f"{label}编号混用“图1”与“图1-1”两种体例"
    prev = None
    for ch, seq in items:
        if prev is not None:
            pch, pseq = prev
            if ch < pch:
                return f"{label}编号章号倒退：{pch}-{pseq} → {ch}-{seq}"
            if ch == pch and seq != pseq + 1:
                return f"{label}{ch}-{seq} 与上一编号 {pch}-{pseq} 不连续"
        prev = (ch, seq)
    return None


def caption_label(label, entry):
    ch, seq = entry
    return f"{label}{ch}-{seq}" if seq is not None else f"{label}{ch}"


def expand_citations(text):
    nums = set()
    for m in CITE_RE.finditer(text):
        for token in re.split(r"[,，]", m.group(1)):
            token = token.strip()
            if not token:
                continue
            rng = re.match(r"^(\d+)\s*[-–—]\s*(\d+)$", token)
            if rng:
                a, b = int(rng.group(1)), int(rng.group(2))
                if a <= b and b - a <= 50:
                    nums.update(range(a, b + 1))
            elif token.isdigit():
                nums.add(int(token))
    return nums


def main():
    ap = argparse.ArgumentParser(description="实验报告 DOCX 自检")
    ap.add_argument("path", help="实验报告 docx 路径")
    ap.add_argument("--strict", action="store_true",
                    help="存在 WARN 时退出码为 1（默认始终为 0）")
    args = ap.parse_args()

    doc = Document(args.path)
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

    # 2 空章节（表格、图片也算内容）
    empty = empty_headings(doc)
    if empty:
        warns.append("空标题（无内容）：" + "、".join(empty[:4]))

    # 3 图表编号（全文连续 或 按章编号）
    fig_entries, tab_entries, fig_caps, tab_caps = [], [], [], []
    for p in paras:
        t = p.text.strip()
        if p.style.name != CAPTION_STYLE and (len(t) > 40 or re.search(r"[。！？；]$", t)):
            continue
        entry = parse_caption(t, FIG_RE)
        if entry is not None:
            fig_entries.append(entry)
            fig_caps.append(t)
        entry = parse_caption(t, TAB_RE)
        if entry is not None:
            tab_entries.append(entry)
            tab_caps.append(t)
    for label, entries in (("图", fig_entries), ("表", tab_entries)):
        issue = check_numbering(label, entries)
        if issue:
            warns.append(issue)
        elif entries:
            passes.append(f"{label}编号连续（共 {len(entries)} 个）")
    if fig_entries and len(doc.inline_shapes) < len(fig_entries):
        warns.append(f"图题 {len(fig_entries)} 个，实际图片 {len(doc.inline_shapes)} 张，可能缺图")
    if tab_entries and len(doc.tables) < len(tab_entries):
        warns.append(f"表题 {len(tab_entries)} 个，实际表格 {len(doc.tables)} 个，可能缺表")

    # 4 图表正文引用
    body_text = "\n".join(p.text for p in paras if not is_caption_paragraph(p))
    unreferenced = []
    for label, entries in (("图", fig_entries), ("表", tab_entries)):
        for entry in entries:
            if entry[1] is not None:
                pat = rf"{label}\s*{entry[0]}\s*[-–—.]\s*{entry[1]}(?![0-9])"
            else:
                pat = rf"{label}\s*{entry[0]}(?![0-9\-–—.])"
            if not re.search(pat, body_text):
                unreferenced.append(caption_label(label, entry))
    if unreferenced:
        warns.append("正文未引用：" + "、".join(unreferenced))
    elif fig_entries or tab_entries:
        passes.append("图表均在正文被引用")

    # 5 正文格式
    normal = doc.styles["Normal"]
    size = normal.font.size.pt if normal.font.size else None
    pf = normal.paragraph_format
    spacing = pf.line_spacing
    indent = pf.first_line_indent
    ea = style_east_asian(normal)
    fmt_issues = []
    if size and abs(size - 12) > 0.5:
        fmt_issues.append(f"正文字号 {size} pt（通用规范 12 pt，模板要求优先）")
    if spacing is not None:
        if pf.line_spacing_rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
            pt = getattr(spacing, "pt", None)
            value = f"{pt:.1f} pt" if pt is not None else str(spacing)
            fmt_issues.append(f"正文行距为固定值/最小值 {value}（通用规范 1.5 倍，模板要求优先）")
        else:
            try:
                multiple = float(spacing)
            except (TypeError, ValueError):
                multiple = None
            if multiple is not None and abs(multiple - 1.5) > 0.05:
                fmt_issues.append(f"正文行距 {multiple} 倍（通用规范 1.5 倍，模板要求优先）")
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

    # 9 参考文献：条数、类型标识、与正文引用对应
    ref_paras = [p for p in paras if REF_NUM_RE.match(p.text.strip())]
    ref_nums = [int(REF_NUM_RE.match(p.text.strip()).group(1)) for p in ref_paras]
    if any("参考文献" in t for t in h1) and not ref_paras:
        warns.append("有参考文献章节但未列出条目")
    elif ref_paras:
        passes.append(f"参考文献 {len(ref_paras)} 条")
        for p, num in zip(ref_paras, ref_nums):
            if not REF_TYPE_RE.search(p.text):
                warns.append(f"参考文献 [{num}] 缺少文献类型标识（如 [J]/[M]/[EB/OL]）")
        body_only = "\n".join(
            p.text for p in paras
            if not REF_NUM_RE.match(p.text.strip()) and not is_caption_paragraph(p))
        cited = expand_citations(body_only)
        not_listed = sorted(cited - set(ref_nums))
        if not_listed:
            warns.append("正文引用未在文献表列出：" + "、".join(f"[{n}]" for n in not_listed))
        uncited = sorted(set(ref_nums) - cited)
        if uncited:
            infos.append("文献未在正文引用：" + "、".join(f"[{n}]" for n in uncited))

    # 10 字数
    text = "\n".join(p.text for p in paras)
    char_count = len(CJK_RE.findall(text))
    infos.append(f"中文字数约 {char_count} 字，段落 {len([p for p in paras if p.text.strip()])} 段，"
                 f"图片 {len(doc.inline_shapes)} 张，表格 {len(doc.tables)} 个")

    print(f"检查文件：{args.path}")
    for item in passes:
        print(f"[PASS] {item}")
    for item in infos:
        print(f"[INFO] {item}")
    for item in warns:
        print(f"[WARN] {item}")
    print(f"小结：{len(warns)} 处需修正，{len(infos)} 条提示。")
    if args.strict and warns:
        sys.exit(1)


if __name__ == "__main__":
    main()
