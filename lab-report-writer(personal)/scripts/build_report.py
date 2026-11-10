#!/usr/bin/env python3
"""实验报告 Markdown -> 规范 DOCX。

用法：
    python3 build_report.py report.md -o 实验报告.docx
    python3 build_report.py report.md -o 实验报告.docx --template 模板.docx

report.md 结构：
    title: 实验三 进程调度
    course: 操作系统
    name: 唐浚淇
    student_id: 202403213037
    ...

    # 1 实验目的
    正文段落……
    [[FIG 图1 系统结构 | img/a.png | 12]]
    [[TAB 表1 实验数据]]
    | 序号 | 输入 | 输出 |
    | 1 | 10 | 20 |
    [[CODE 代码1 关键函数]]
    ```python
    print(1)
    ```
    [[EQ (1)]] T = P × C
"""

import argparse
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BODY_SIZE = 12          # 小四
CAPTION_SIZE = 10.5     # 五号
CODE_SIZE = 10.5
PAGE_W, PAGE_H = 21.0, 29.7
MARGIN_TB, MARGIN_LR = 2.54, 3.17
USABLE_W = PAGE_W - 2 * MARGIN_LR
META_KEYS = {
    "title", "course", "name", "student_id", "class_name", "class",
    "major", "teacher", "date", "cover", "subject", "school",
}


# ---------------------------------------------------------------- utilities
def set_run_font(run, size=BODY_SIZE, ascii_font="Times New Roman",
                 ea_font="宋体", bold=False, italic=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = ascii_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea_font)
    rfonts.set(qn("w:cs"), ascii_font)
    if color is not None:
        run.font.color.rgb = color


def set_style_font(style, size, ea_font="宋体", ascii_font="Times New Roman", bold=None):
    style.font.size = Pt(size)
    style.font.name = ascii_font
    if bold is not None:
        style.font.bold = bold
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea_font)
    rfonts.set(qn("w:cs"), ascii_font)


CJK_PUNCT = set("，。；：、（）！？“”‘’《》【】—…·")


def is_cjk(ch):
    return bool(ch) and ("\u4e00" <= ch <= "\u9fff" or ch in CJK_PUNCT)


def smart_join(parts):
    out = ""
    for p in parts:
        if not out:
            out = p
        elif is_cjk(out[-1]) or is_cjk(p[0]):
            out += p
        else:
            out += " " + p
    return out


def norm_heading(text):
    s = text.strip().replace("\u3000", " ")
    s = re.sub(r"^[0-9]+(\.[0-9]+)*[.、)]?\s*", "", s)
    s = re.sub(r"^[一二三四五六七八九十]+[、.]\s*", "", s)
    return re.sub(r"\s+", "", s)


# ---------------------------------------------------------------- parsing
def parse_meta(lines):
    meta, i = {}, 0
    while i < len(lines):
        s = lines[i].strip()
        if not s or s == "---":
            i += 1
            continue
        m = re.match(r"^([A-Za-z_][\w]*)\s*[:：]\s*(.*)$", s)
        if m and m.group(1).lower() in META_KEYS:
            key = m.group(1).lower()
            meta["class_name" if key == "class" else key] = m.group(2).strip()
            i += 1
            continue
        break
    return meta, lines[i:]


def parse_blocks(lines):
    blocks, i, n = [], 0, len(lines)
    while i < n:
        raw = lines[i]
        s = raw.strip()
        if not s:
            i += 1
            continue

        m = re.match(r"^#{1,6}\s+(.*)$", s)
        if m:
            level = len(s) - len(s.lstrip("#"))
            blocks.append({"type": "h", "level": min(level, 3), "text": m.group(1).strip()})
            i += 1
            continue

        if s == "[[PAGEBREAK]]":
            blocks.append({"type": "pagebreak"})
            i += 1
            continue

        m = re.match(r"^\[\[FIG\s+(.*?)\]\]$", s)
        if m:
            parts = [p.strip() for p in m.group(1).split("|")]
            caption = parts[0]
            path = parts[1] if len(parts) > 1 else ""
            width = None
            if len(parts) > 2 and parts[2]:
                try:
                    width = float(parts[2])
                except ValueError:
                    width = None
            blocks.append({"type": "figure", "caption": caption, "path": path, "width": width})
            i += 1
            continue

        m = re.match(r"^\[\[TAB\s+(.*?)\]\]$", s)
        if m:
            caption, rows = m.group(1).strip(), []
            i += 1
            while i < n and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            rows = [r for r in rows if not all(re.fullmatch(r":?-{2,}:?", c or "-") for c in r)]
            blocks.append({"type": "table", "caption": caption, "rows": rows})
            continue

        m = re.match(r"^\[\[CODE\s+(.*?)\]\]$", s)
        if m:
            caption, code = m.group(1).strip(), []
            i += 1
            if i < n and lines[i].strip().startswith("```"):
                i += 1
                while i < n and not lines[i].strip().startswith("```"):
                    code.append(lines[i].rstrip())
                    i += 1
                i += 1
            blocks.append({"type": "code", "caption": caption, "lines": code})
            continue

        m = re.match(r"^\[\[EQ\s+(.*?)\]\]\s*(.*)$", s)
        if m:
            blocks.append({"type": "eq", "number": m.group(1).strip(), "text": m.group(2).strip()})
            i += 1
            continue

        if s.startswith("```"):
            code = []
            i += 1
            while i < n and not lines[i].strip().startswith("```"):
                code.append(lines[i].rstrip())
                i += 1
            i += 1
            blocks.append({"type": "code", "caption": "", "lines": code})
            continue

        if s.startswith("|"):
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            rows = [r for r in rows if not all(re.fullmatch(r":?-{2,}:?", c or "-") for c in r)]
            blocks.append({"type": "table", "caption": "", "rows": rows})
            continue

        if re.match(r"^[（(]\d+[)）]", s) or re.match(r"^[（(][一二三四五六七八九十]+[)）]", s) \
                or re.match(r"^\d+、", s):
            blocks.append({"type": "p", "text": s})
            i += 1
            continue

        if re.match(r"^[-*+]\s+", s) or re.match(r"^\d+[.)]\s+", s):
            items, ordered = [], None
            while i < n:
                t = lines[i].strip()
                mo = re.match(r"^[-*+]\s+(.*)$", t)
                mn = re.match(r"^(\d+)[.)]\s+(.*)$", t)
                if mo:
                    ordered = False if ordered is None else ordered
                    items.append(mo.group(1).strip())
                elif mn:
                    ordered = True if ordered is None else ordered
                    items.append(mn.group(2).strip())
                elif t and t.startswith("  ") and items:
                    items[-1] += t
                else:
                    break
                i += 1
            blocks.append({"type": "list", "ordered": bool(ordered), "items": items})
            continue

        if re.match(r"^\[\d+\]", s) and re.search(r"\[(J|M|D|S|C|N|R|P|EB|DB|CP)\]", s):
            blocks.append({"type": "ref", "text": s})
            i += 1
            continue

        para = [s]
        i += 1
        while i < n:
            t = lines[i].strip()
            if (not t or t.startswith("#") or t.startswith("[[") or t.startswith("|")
                    or t.startswith("```") or re.match(r"^[-*+]\s+", t)
                    or re.match(r"^\d+[.)]\s+", t)
                    or re.match(r"^[（(]\d+[)）]", t)
                    or re.match(r"^[（(][一二三四五六七八九十]+[)）]", t)
                    or re.match(r"^\d+、", t)):
                break
            para.append(t)
            i += 1
        blocks.append({"type": "p", "text": smart_join(para)})
    return blocks


def split_sections(blocks):
    sections, cur = [], None
    for b in blocks:
        if b["type"] == "h":
            cur = {"heading": b["text"], "level": b["level"], "blocks": []}
            sections.append(cur)
        elif cur is not None:
            cur["blocks"].append(b)
        else:
            sections.append({"heading": "", "level": 0, "blocks": [b]})
    return sections


# ---------------------------------------------------------------- doc setup
def setup_document(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(PAGE_W), Cm(PAGE_H)
    sec.top_margin = sec.bottom_margin = Cm(MARGIN_TB)
    sec.left_margin = sec.right_margin = Cm(MARGIN_LR)

    normal = doc.styles["Normal"]
    set_style_font(normal, BODY_SIZE)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(24)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for name, size, before in (("Heading 1", 14, 12), ("Heading 2", 12, 6), ("Heading 3", 12, 6)):
        style = doc.styles[name]
        set_style_font(style, size, ea_font="黑体", bold=True)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.first_line_indent = Pt(0)
        style.paragraph_format.keep_with_next = True
        try:
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        except Exception:
            pass

    ensure_aux_styles(doc)


def ensure_aux_styles(doc):
    """补建图表题注、代码、参考文献三个辅助样式（模板模式下同样适用）。"""
    existing = set()
    try:
        for st in doc.styles:
            existing.add(st.name)
    except Exception:
        pass
    if "LabCaption" not in existing:
        caption = doc.styles.add_style("LabCaption", 1)
        caption.base_style = doc.styles["Normal"]
        set_style_font(caption, CAPTION_SIZE)
        caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.paragraph_format.first_line_indent = Pt(0)
        caption.paragraph_format.space_before = Pt(6)
        caption.paragraph_format.space_after = Pt(6)
        caption.paragraph_format.line_spacing = 1.5
    if "LabCode" not in existing:
        code = doc.styles.add_style("LabCode", 1)
        code.base_style = doc.styles["Normal"]
        set_style_font(code, CODE_SIZE, ea_font="Consolas", ascii_font="Consolas")
        code.paragraph_format.first_line_indent = Pt(0)
        code.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        code.paragraph_format.line_spacing = 1.0
        code.paragraph_format.space_before = Pt(6)
        code.paragraph_format.space_after = Pt(6)
    if "LabRef" not in existing:
        ref = doc.styles.add_style("LabRef", 1)
        ref.base_style = doc.styles["Normal"]
        set_style_font(ref, CAPTION_SIZE)
        ref.paragraph_format.left_indent = Pt(0)
        ref.paragraph_format.first_line_indent = Pt(0)
        ref.paragraph_format.line_spacing = 1.5


def add_page_number(doc):
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    set_run_font(run, CAPTION_SIZE)

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr, fld_end):
        run._element.append(el)


def add_cover(doc, meta, mode):
    if mode == "none":
        return
    title = meta.get("title", "")
    info = []
    line1 = []
    if meta.get("course"):
        line1.append("课程名称：" + meta["course"])
    if title:
        line1.append("实验名称：" + title)
    if line1:
        info.append("　　".join(line1))
    line2 = []
    if meta.get("name"):
        line2.append("姓　　名：" + meta["name"])
    if meta.get("student_id"):
        line2.append("学　　号：" + meta["student_id"])
    if line2:
        info.append("　　".join(line2))
    line3 = []
    if meta.get("class_name"):
        line3.append("班　　级：" + meta["class_name"])
    if meta.get("major"):
        line3.append("专　　业：" + meta["major"])
    if line3:
        info.append("　　".join(line3))
    line4 = []
    if meta.get("teacher"):
        line4.append("指导教师：" + meta["teacher"])
    if meta.get("date"):
        line4.append("日　　期：" + meta["date"])
    if line4:
        info.append("　　".join(line4))

    if mode == "page":
        for _ in range(4):
            doc.add_paragraph()
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(24)
    set_run_font(p.add_run(title or "实验报告"), 16, ea_font="黑体", bold=True)

    for line in info:
        q = doc.add_paragraph()
        q.paragraph_format.first_line_indent = Pt(0)
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(q.add_run(line), BODY_SIZE)

    if mode == "page":
        doc.add_page_break()


# ---------------------------------------------------------------- rendering
def add_caption(doc, text):
    if not text:
        return
    p = doc.add_paragraph(style="LabCaption")
    set_run_font(p.add_run(text), CAPTION_SIZE)
    return p


def add_paragraph(doc, text):
    p = doc.add_paragraph()
    set_run_font(p.add_run(text), BODY_SIZE)
    return p


def add_list(doc, items, ordered):
    for idx, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Pt(24)
        prefix = f"（{idx}）" if ordered else "• "
        set_run_font(p.add_run(prefix + item), BODY_SIZE)


def add_figure(doc, caption, path, width, base_dir):
    full = path if os.path.isabs(path) else os.path.join(base_dir, path)
    if not os.path.exists(full):
        add_paragraph(doc, f"［缺图：{path}］")
        add_caption(doc, caption)
        return
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    run.add_picture(full, width=Cm(min(width or 12.0, USABLE_W)))
    add_caption(doc, caption)


def _clear_borders(table):
    tbl_pr = table._tbl.tblPr
    for old in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "12")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
        el.set(qn("w:space"), "0")
        borders.append(el)
    tbl_pr.append(borders)
    for cell in table.rows[0].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_b = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:color"), "000000")
        tc_b.append(bottom)
        tc_pr.append(tc_b)


def add_table(doc, caption, rows):
    add_caption(doc, caption)
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    lens = [max((len(r[c]) if c < len(r) else 0) for r in rows) for c in range(cols)]
    total = sum(lens) or cols
    ratio = [max(0.16, v / total) for v in lens]
    ratio = [v / sum(ratio) for v in ratio]
    for c in range(cols):
        width = Cm(USABLE_W * ratio[c])
        table.columns[c].width = width
        for r in range(len(rows)):
            table.cell(r, c).width = width
    tbl_pr = table._tbl.tblPr
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(int(USABLE_W * 567)))
    tbl_pr.append(tbl_w)
    for r, row in enumerate(rows):
        for c in range(cols):
            text = row[c] if c < len(row) else ""
            cell = table.cell(r, c)
            p = cell.paragraphs[0]
            for old_run in list(p.runs):
                old_run._element.getparent().remove(old_run._element)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0
            set_run_font(p.add_run(text), CAPTION_SIZE, bold=(r == 0))
    _clear_borders(table)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.first_line_indent = Pt(0)
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = Pt(6)
    set_run_font(spacer.add_run(""), 6)


def add_code(doc, caption, lines):
    add_caption(doc, caption)
    p = doc.add_paragraph(style="LabCode")
    p.paragraph_format.keep_together = True
    for idx, line in enumerate(lines):
        run = p.add_run(line if line else " ")
        if idx < len(lines) - 1:
            run.add_break()
        set_run_font(run, CODE_SIZE, ea_font="Consolas", ascii_font="Consolas")
    ppr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F2F2F2")
    ppr.append(shd)
    pbdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "4")
        el.set(qn("w:color"), "BFBFBF")
        pbdr.append(el)
    ppr.append(pbdr)
    add_paragraph(doc, "")


def add_equation(doc, number, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.tab_stops.add_tab_stop(Cm(USABLE_W / 2), WD_TAB_ALIGNMENT.CENTER)
    pf.tab_stops.add_tab_stop(Cm(USABLE_W), WD_TAB_ALIGNMENT.RIGHT)
    run = p.add_run("\t" + text + ("\t" + number if number else ""))
    set_run_font(run, BODY_SIZE, ascii_font="Cambria Math", ea_font="Cambria Math")


def add_ref(doc, text):
    p = doc.add_paragraph(style="LabRef")
    p.paragraph_format.left_indent = Pt(24)
    p.paragraph_format.first_line_indent = Pt(-24)
    set_run_font(p.add_run(text), CAPTION_SIZE)


def render_block(doc, b, base_dir):
    t = b["type"]
    if t == "h":
        try:
            doc.add_heading(b["text"], level=b["level"])
        except KeyError:
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(0)
            set_run_font(p.add_run(b["text"]),
                         BODY_SIZE if b["level"] > 1 else 14,
                         ea_font="黑体", bold=True)
    elif t == "p":
        add_paragraph(doc, b["text"])
    elif t == "list":
        add_list(doc, b["items"], b["ordered"])
    elif t == "figure":
        add_figure(doc, b["caption"], b["path"], b["width"], base_dir)
    elif t == "table":
        add_table(doc, b["caption"], b["rows"])
    elif t == "code":
        add_code(doc, b["caption"], b["lines"])
    elif t == "eq":
        add_equation(doc, b["number"], b["text"])
    elif t == "pagebreak":
        doc.add_page_break()
    elif t == "ref":
        add_ref(doc, b["text"])


def find_template_heading(doc, text):
    target = norm_heading(text)
    if not target:
        return None
    candidates = []
    for p in doc.paragraphs:
        style = (p.style.name or "").lower()
        candidates.append(p)
        if target == norm_heading(p.text):
            return p
    for p in candidates:
        if target and target in norm_heading(p.text):
            return p
    return None


def find_tail_anchor(doc):
    """模板末尾的固定栏目（评语、成绩、签字等），补充章节应插在它之前。"""
    for p in reversed(doc.paragraphs):
        if re.search(r"评语|成绩|评分|意见|签字", p.text):
            return p
    return None


def append_section(doc, section, base_dir, before_p=None):
    body = doc.element.body
    before = list(body)
    try:
        doc.add_heading(section["heading"], level=section["level"] or 1)
    except KeyError:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(0)
        set_run_font(p.add_run(section["heading"]), 14, ea_font="黑体", bold=True)
    for b in section["blocks"]:
        render_block(doc, b, base_dir)
    new_els = [el for el in body if el not in before]
    if before_p is not None:
        for el in new_els:
            before_p.addprevious(el)


def render_section_into_template(doc, section, base_dir):
    anchor = find_template_heading(doc, section["heading"])
    if anchor is None:
        tail = find_tail_anchor(doc)
        print(f"[warn] 模板未找到章节：{section['heading']}，插入到"
              + ("固定栏目之前" if tail is not None else "文末"))
        append_section(doc, section, base_dir, tail._p if tail is not None else None)
        return
    body = doc.element.body
    before = list(body)
    for b in section["blocks"]:
        render_block(doc, b, base_dir)
    new_els = [el for el in body if el not in before]
    cursor = anchor._p
    for el in new_els:
        cursor.addnext(el)
        cursor = el


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="实验报告 Markdown 转 DOCX")
    ap.add_argument("source", help="report.md 路径")
    ap.add_argument("-o", "--output", required=True, help="输出 docx 路径")
    ap.add_argument("--template", help="按模板 docx 填充")
    args = ap.parse_args()

    if not os.path.exists(args.source):
        sys.exit(f"源文件不存在：{args.source}")
    with open(args.source, encoding="utf-8") as f:
        lines = f.read().replace("\r\n", "\n").split("\n")

    meta, body_lines = parse_meta(lines)
    blocks = parse_blocks(body_lines)
    base_dir = os.path.dirname(os.path.abspath(args.source))

    if args.template:
        if not os.path.exists(args.template):
            sys.exit(f"模板不存在：{args.template}")
        doc = Document(args.template)
        ensure_aux_styles(doc)
        for section in split_sections(blocks):
            if section["heading"]:
                render_section_into_template(doc, section, base_dir)
            else:
                for b in section["blocks"]:
                    render_block(doc, b, base_dir)
    else:
        doc = Document()
        setup_document(doc)
        cover = meta.get("cover", "auto")
        if cover == "auto":
            cover = "simple" if meta.get("title") and (
                meta.get("name") or meta.get("course") or meta.get("student_id")) else "none"
        if cover not in ("none", "simple", "page"):
            cover = "simple"
        add_cover(doc, meta, cover)
        for b in blocks:
            render_block(doc, b, base_dir)
        add_page_number(doc)

    if meta.get("title"):
        doc.core_properties.title = meta["title"]
    out_dir = os.path.dirname(os.path.abspath(args.output))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(args.output)
    print(f"已生成：{args.output}（{len(blocks)} 个内容块）")


if __name__ == "__main__":
    main()
