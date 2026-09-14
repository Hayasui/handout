# -*- coding: utf-8 -*-
"""render_handout —— 把一份结构化 markdown 渲染成 A4 打印用的 PDF 与可编辑的 docx。

一次解析，两条渲染线：
    markdown → 元素列表 ─┬─→ HTML → Edge 无头打印 → PDF
                        └─→ python-docx → docx

两条线的分页不保证逐页一致，**打印以 PDF 为准**，docx 当可编辑副本。

约定（渲染器认哪些写法）见 references/markdown-conventions.md，
版式的默认值与色值见 references/layout-spec.md。

本机渲染链的特殊之处（Edge 路径、纯 ASCII 中转目录）见 references/environment.md。
"""

import argparse
import io
import json
import os
import re
import shutil
import subprocess
import sys

# ================================================================ 默认值

DEFAULT_STYLE = {
    "font_en": "Segoe UI",
    "font_cn": "Microsoft YaHei",
    "size_body": 10,
    "size_h1": 19,
    "size_h2": 16,
    "size_h3": 13.5,
    "size_h4": 12,
    "size_h5": 10.5,
    "size_field": 9.5,
    "size_pair_en": 11.5,
    "size_pair_cn": 9,
    "size_bullet": 11,
    "size_bullet_cn": 9.5,
    "size_warn": 9.5,
    "size_table": 9.5,
    "size_page_number": 8,
    "page_width": 21.0,
    "page_height": 29.7,
    "margin_top": 1.4,
    "margin_bottom": 1.3,
    "margin_left": 1.5,
    "margin_right": 1.5,
    "html_page_margin": "14mm 13mm 12mm 13mm",
    "rule_height_pt": 22,
    "rule_line_pt": 23,
    "ink": "141414",
    "grey": "6B6B6B",
    "ink_body": "333333",
    "accent": "1F4E79",
    "field": "2E6DA4",
    "stage": "808080",
    "h4_ink": "17365D",
    "h5_ink": "44546A",
    "h3_bar": "9DC3E6",
    "head_bg": "EAF1F8",
    "table_border": "CADCED",
    "rule_color": "BDBDBD",
    "warn_ink": "9E2B25",
    "warn_body": "6B2A26",
    "warn_bg": "FDF2F2",
    "warn_bar": "C0392B",
    "warn_hair": "F0C8C4",
    "warn_head_bg": "F7E3E1",
    "warn_table_hair": "E8C4C0",
    "page_number_ink": "9A9A9A",
}

# 引用块里出现的「主持人栏」——整块套红框
DEFAULT_QUOTE_LABELS = ["内部提醒", "内部参考词库"]
# 正文里以 `**栏名**：…` 起头、自带清单的「主持人栏」
DEFAULT_BOX_LABELS = ["记录清单（主持人用）", "提示清单（主持人用）"]
# 见到这个栏名，就在它后面的内容底下补手写横线
LINES_FIELD = "观察记录点"

DEFAULTS = {
    "md": None,
    "out_dir": None,
    "stem": None,
    "title": None,
    "formats": ["docx", "pdf"],
    "lines": 4,
    "lines_rules": {},
    "quote_labels": list(DEFAULT_QUOTE_LABELS),
    "box_labels": list(DEFAULT_BOX_LABELS),
    "tmp_dir": r"C:\Agent\_handout_tmp",
    "edge": None,
    "report": None,
    "verify": False,
    "refs": [],
    "q_range": None,
    "style": {},
}

EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


# ================================================================ 配置

def load_config(args):
    cfg = json.loads(json.dumps(DEFAULTS))
    base_dir = os.getcwd()

    if args.config:
        p = os.path.abspath(args.config)
        with io.open(p, encoding="utf-8") as f:
            user = json.load(f)
        base_dir = os.path.dirname(p)
        for k, v in user.items():
            if k == "style":
                cfg["style"].update(v or {})
            elif k in cfg:
                cfg[k] = v
            else:
                raise SystemExit("配置里有不认识的键：%s" % k)

    # 命令行覆盖配置
    def rel(p):
        return p if os.path.isabs(p) else os.path.join(base_dir, p)

    for key, attr in (("out_dir", "out_dir"), ("stem", "stem"), ("title", "title"),
                      ("report", "report"), ("edge", "edge"), ("tmp_dir", "tmp_dir")):
        v = getattr(args, attr)
        if v:
            cfg[key] = v
    if args.formats:
        cfg["formats"] = [x.strip().lower() for x in args.formats.split(",") if x.strip()]
    if args.lines is not None:
        cfg["lines"] = args.lines
    for rule in args.lines_rule or []:
        if "=" not in rule:
            raise SystemExit("--lines-rule 要写成「前缀=条数」，收到：%s" % rule)
        k, v = rule.rsplit("=", 1)
        cfg["lines_rules"][k.strip()] = int(v)
    if args.ref:
        cfg["verify"] = True
        cfg["refs"] = list(args.ref)
    if args.verify:
        cfg["verify"] = True
    if args.q_range:
        cfg["q_range"] = args.q_range
    for kv in args.style or []:
        if "=" not in kv:
            raise SystemExit("--style 要写成 key=value，收到：%s" % kv)
        k, v = kv.split("=", 1)
        cfg["style"][k.strip()] = v.strip()

    md = args.md or cfg.get("md")
    if not md:
        raise SystemExit("必须给 --md，或在配置里写 md")
    cfg["md"] = rel(md)
    if not os.path.exists(cfg["md"]):
        raise SystemExit("找不到输入文件：%s" % cfg["md"])
    cfg["out_dir"] = rel(cfg["out_dir"]) if cfg["out_dir"] else os.path.dirname(cfg["md"])
    cfg["stem"] = cfg["stem"] or os.path.splitext(os.path.basename(cfg["md"]))[0]
    cfg["title"] = cfg["title"] or cfg["stem"]
    cfg["refs"] = [rel(r) for r in cfg["refs"]]
    cfg["tmp_dir"] = rel(cfg["tmp_dir"])
    cfg["report"] = rel(cfg["report"]) if cfg["report"] else os.path.join(cfg["tmp_dir"], "render_report.txt")
    cfg["edge"] = cfg["edge"] or find_edge()
    cfg["style"] = dict(DEFAULT_STYLE, **cfg["style"])
    return cfg


def find_edge():
    for p in EDGE_CANDIDATES:
        if os.path.exists(p):
            return p
    found = shutil.which("msedge")
    if found:
        return found
    raise SystemExit("找不到 Edge，请用 --edge 指定 msedge.exe 的完整路径")


# ================================================================ 解析 markdown

RICH = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*)")
TAG = re.compile(r"^\*\*(EN|中文)\*\*[　\s]*(.*)$")


def label_re(labels, colon_required):
    """按栏名拼正则：整名优先，匹配不上再剥掉尾部括号。"""
    names = sorted(labels, key=len, reverse=True)
    alt = "|".join(re.escape(n) for n in names)
    tail = r"[:：]" if colon_required else r"[:：]?"
    return re.compile(r"^\*\*(?P<label>" + alt + r")\*\*"
                      r"\s*(?P<paren>（[^）]*）)?\s*" + tail + r"\s*(?P<body>.*)$")


class Parser(object):
    def __init__(self, cfg):
        self.cfg = cfg
        self.quote_re = label_re(cfg["quote_labels"], colon_required=False)
        self.box_re = label_re(cfg["box_labels"], colon_required=True)

    def parse(self, path):
        lines = io.open(path, encoding="utf-8").read().replace("\ufeff", "").split("\n")
        els, i, n = [], 0, len(lines)
        while i < n:
            st = lines[i].strip()
            if st == "":
                i += 1
                continue
            if st.startswith("<!--"):
                while i < n and "-->" not in lines[i]:
                    i += 1
                i += 1
                continue
            for pre, kind in (("##### ", "h5"), ("#### ", "h4"), ("### ", "h3"),
                              ("## ", "h2"), ("# ", "h1")):
                if st.startswith(pre):
                    els.append((kind, st[len(pre):].strip()))
                    break
            else:
                if st.startswith(">"):
                    run = []
                    while i < n and lines[i].strip().startswith(">"):
                        run.append(lines[i].strip()[1:].strip())
                        i += 1
                    els.extend(self.parse_quote(run))
                    continue
                if st.startswith("|"):
                    rows = []
                    while i < n and lines[i].strip().startswith("|"):
                        s = lines[i].strip()
                        if not re.match(r"^\|[\s:|-]+\|$", s):
                            rows.append([c.strip() for c in s.strip("|").split("|")])
                        i += 1
                    els.append(("table", rows))
                    continue
                if st.startswith("[[lines:"):
                    # 兼容旧稿。新稿不再写这种标记，横线一律由栏名触发。
                    els.append(("lines", int(re.search(r"\d+", st).group())))
                    i += 1
                    continue
                if st.startswith("- "):
                    items = []
                    while i < n and lines[i].strip().startswith("- "):
                        items.append(lines[i].strip()[2:].strip())
                        i += 1
                    els.extend(parse_bullets(items))
                    continue
                if st.startswith("**") and st.endswith("**") and st.count("**") == 2:
                    els.append(("field", st.strip("*").strip()))
                    i += 1
                    continue
                els.append(("para", st))
                i += 1
                continue
            i += 1
        return self.postprocess(els)

    def parse_quote(self, run):
        m = self.quote_re.match(run[0]) if run else None
        if m:
            title = m.group("label") + (m.group("paren") or "")
            body = []
            if m.group("body").strip():
                body.append(m.group("body").strip())
            for l in run[1:]:
                if l.strip():
                    body.append(l)
            return [("warn", title, body)]
        en, cn = [], []
        for l in run:
            mm = TAG.match(l)
            if not mm:
                continue
            (en if mm.group(1) == "EN" else cn).append(mm.group(2).strip())
        return [("pair", "\n".join([x for x in en if x]), "\n".join([x for x in cn if x]))]

    def postprocess(self, els):
        """按栏名补出红框与手写横线——这些都不写进 markdown。"""
        out, i, seg, first = [], 0, "", True
        while i < len(els):
            el = els[i]
            if el[0].startswith("h"):
                if el[0] == "h3":
                    seg = el[1]
                out.append(el)
                i += 1
                continue
            if el[0] == "para":
                m = self.box_re.match(el[1])
                if m:
                    title = m.group("label") + (m.group("paren") or "")
                    body = [m.group("body")] if m.group("body").strip() else []
                    i += 1
                    while i < len(els):
                        e2 = els[i]
                        if e2[0].startswith("h") or e2[0] in ("field", "warn"):
                            break
                        if e2[0] == "para":
                            body.append(e2[1])
                        elif e2[0] == "bcn":
                            body.append("- " + e2[1])
                        elif e2[0] == "table":
                            for row in e2[1]:
                                body.append("| " + " | ".join(row) + " |")
                        elif e2[0] == "lines":
                            body.append("[[lines:%d]]" % e2[1])
                        else:
                            break
                        i += 1
                    out.append(("warn", title, body))
                    continue
            if el[0] == "field" and el[1] == LINES_FIELD:
                out.append(el)
                i += 1
                while i < len(els):
                    e2 = els[i]
                    if e2[0].startswith("h") or e2[0] in ("field", "warn"):
                        break
                    out.append(e2)
                    i += 1
                out.append(("lines", self.lines_for(seg)))
                first = False
                continue
            out.append(el)
            i += 1
        return out

    def lines_for(self, seg):
        for prefix, n in self.cfg["lines_rules"].items():
            if seg.startswith(prefix):
                return int(n)
        return int(self.cfg["lines"])


def parse_bullets(items):
    out, i = [], 0
    while i < len(items):
        it = items[i]
        m = TAG.match(it)
        if m and m.group(1) == "EN":
            cn = ""
            if i + 1 < len(items):
                m2 = TAG.match(items[i + 1])
                if m2 and m2.group(1) == "中文":
                    cn = m2.group(2).strip()
                    i += 1
            out.append(("bpair", m.group(2).strip(), cn))
        elif m and m.group(1) == "中文":
            out.append(("bcn", m.group(2).strip()))
        else:
            out.append(("bcn", it))
        i += 1
    return out


# ================================================================ DOCX

def build_docx(els, cfg, out_path):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    S = cfg["style"]
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(S["page_width"]), Cm(S["page_height"])
    sec.top_margin, sec.bottom_margin = Cm(S["margin_top"]), Cm(S["margin_bottom"])
    sec.left_margin, sec.right_margin = Cm(S["margin_left"]), Cm(S["margin_right"])

    normal = doc.styles["Normal"]
    normal.font.name = S["font_en"]
    normal.font.size = Pt(S["size_body"])
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), S["font_cn"])

    def style_run(r, size, color, bold=False, italic=False, font=None):
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = RGBColor.from_string(color)
        r.font.name = font or S["font_en"]
        rpr = r._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        rf.set(qn("w:ascii"), font or S["font_en"])
        rf.set(qn("w:hAnsi"), font or S["font_en"])
        rf.set(qn("w:eastAsia"), S["font_cn"])

    def rich(p, text, size, color, font=None, italic=False):
        for part in RICH.split(text):
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                style_run(p.add_run(part[2:-2]), size, color, bold=True, font=font)
            elif part.startswith("*") and part.endswith("*") and len(part) > 2:
                style_run(p.add_run(part[1:-1]), size, color, italic=True, font=font)
            else:
                style_run(p.add_run(part), size, color, italic=italic, font=font)

    def P(space_before=0, space_after=3, left=0, hanging=0, line=None):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        if left:
            pf.left_indent = Cm(left)
        if hanging:
            pf.first_line_indent = Cm(-hanging)
        if line:
            pf.line_spacing = line
        return p

    def shade(p, fill):
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), fill)
        pPr.append(shd)

    def left_bar(p, color, sz=18):
        pPr = p._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), str(sz))
        left.set(qn("w:space"), "6")
        left.set(qn("w:color"), color)
        bdr.append(left)
        pPr.append(bdr)

    def rule_line():
        p = P(0, 0, line=Pt(S["rule_line_pt"]))
        p.add_run("")
        pPr = p._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "1")
        b.set(qn("w:color"), S["rule_color"])
        bdr.append(b)
        pPr.append(bdr)

    def set_cell_bg(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), fill)
        tcPr.append(shd)

    def table_borders(t, left=None, hair=None):
        hair = hair or S["table_border"]
        tblPr = t._tbl.tblPr
        bd = OxmlElement("w:tblBorders")
        for tag, spec in (("top", ("single", 4, hair)), ("left", left),
                          ("bottom", ("single", 4, hair)), ("right", ("single", 4, hair)),
                          ("insideH", ("single", 4, hair)), ("insideV", ("single", 4, hair))):
            el = OxmlElement("w:" + tag)
            if spec is None:
                el.set(qn("w:val"), "none")
                el.set(qn("w:sz"), "0")
            else:
                el.set(qn("w:val"), spec[0])
                el.set(qn("w:sz"), str(spec[1]))
                el.set(qn("w:color"), spec[2])
            el.set(qn("w:space"), "0")
            bd.append(el)
        tblPr.append(bd)

    def cell_margins(t, v=80, h=110):
        tblPr = t._tbl.tblPr
        mar = OxmlElement("w:tblCellMar")
        for tag, val in (("top", v), ("left", h), ("bottom", v), ("right", h)):
            el = OxmlElement("w:" + tag)
            el.set(qn("w:w"), str(val))
            el.set(qn("w:type"), "dxa")
            mar.append(el)
        tblPr.append(mar)

    def add_table(rows):
        if not rows:
            return
        t = doc.add_table(rows=len(rows), cols=len(rows[0]))
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        t.autofit = True
        table_borders(t)
        cell_margins(t)
        for ri, row in enumerate(rows):
            for ci, cell_text in enumerate(row):
                if ci >= len(rows[0]):
                    continue
                cell = t.cell(ri, ci)
                cell.text = ""
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                rich(p, cell_text, S["size_table"], S["ink"] if ri == 0 else S["ink_body"])
                if ri == 0:
                    for r in p.runs:
                        r.font.bold = True
                    set_cell_bg(cell, S["head_bg"])
        # 嵌套表格之后必须补一个空段落，否则 Word 打不开
        sp = P(2, 6)
        sp.add_run("")
        return t

    def warn_box(title, body):
        t = doc.add_table(rows=1, cols=1)
        t.autofit = True
        table_borders(t, left=("single", 24, S["warn_bar"]), hair=S["warn_hair"])
        cell_margins(t, 90, 130)
        cell = t.cell(0, 0)
        set_cell_bg(cell, S["warn_bg"])
        cell.text = ""
        p0 = cell.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(3)
        r = p0.add_run("▲ " + title)
        style_run(r, S["size_warn"], S["warn_ink"], bold=True)
        for line in body:
            if line == "":
                continue
            if line.startswith("- "):
                p = cell.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.left_indent = Cm(0.45)
                p.paragraph_format.first_line_indent = Cm(-0.35)
                rich(p, "· " + line[2:].strip(), S["size_warn"], S["warn_body"])
            elif line.startswith("[["):
                for _ in range(int(re.search(r"\d+", line).group())):
                    p = cell.add_paragraph()
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = Pt(S["rule_line_pt"])
                    pPr = p._p.get_or_add_pPr()
                    bdr = OxmlElement("w:pBdr")
                    b = OxmlElement("w:bottom")
                    b.set(qn("w:val"), "single")
                    b.set(qn("w:sz"), "4")
                    b.set(qn("w:space"), "1")
                    b.set(qn("w:color"), S["rule_color"])
                    bdr.append(b)
                    pPr.append(bdr)
            elif line.startswith("|"):
                pass
            else:
                p = cell.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                rich(p, line, S["size_warn"], S["warn_body"])
        boxed = [l for l in body if l.startswith("|")]
        if boxed:
            rows = []
            for l in boxed:
                if re.match(r"^\|[\s:|-]+\|$", l):
                    continue
                rows.append([c.strip() for c in l.strip("|").split("|")])
            inner = cell.add_table(rows=len(rows), cols=len(rows[0]))
            table_borders(inner, hair=S["warn_table_hair"])
            for ri, row in enumerate(rows):
                for ci, ct in enumerate(row):
                    ic = inner.cell(ri, ci)
                    ic.text = ""
                    ip = ic.paragraphs[0]
                    ip.paragraph_format.space_before = Pt(0)
                    ip.paragraph_format.space_after = Pt(0)
                    rich(ip, ct, S["size_table"] - 0.5, S["warn_body"])
                    if ri == 0:
                        for r2 in ip.runs:
                            r2.font.bold = True
            cell.add_paragraph()
        sp = P(2, 6)
        sp.add_run("")

    # ---- 页脚页码
    footer = sec.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run()
    style_run(fr, S["size_page_number"], S["page_number_ink"])
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = "PAGE"
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "end")
    fr._r.append(f1)
    fr._r.append(it)
    fr._r.append(f2)

    first_h2 = True
    for el in els:
        kind = el[0]
        if kind == "h1":
            p = P(0, 4)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            style_run(p.add_run(el[1]), S["size_h1"], S["accent"], bold=True)
        elif kind == "h2":
            p = P(0 if first_h2 else 16, 8)
            if not first_h2:
                # 用段落属性分页，不要插空段落的硬分页
                p.paragraph_format.page_break_before = True
            first_h2 = False
            shade(p, S["head_bg"])
            style_run(p.add_run("  " + el[1]), S["size_h2"], S["accent"], bold=True)
        elif kind == "h3":
            p = P(12, 5)
            style_run(p.add_run(el[1]), S["size_h3"], S["accent"], bold=True)
            left_bar(p, S["h3_bar"], 12)
        elif kind == "h4":
            p = P(10, 3)
            style_run(p.add_run(el[1]), S["size_h4"], S["h4_ink"], bold=True)
        elif kind == "h5":
            p = P(8, 2)
            style_run(p.add_run("› " + el[1]), S["size_h5"], S["h5_ink"], bold=True)
        elif kind == "field":
            p = P(7, 3)
            style_run(p.add_run(el[1]), S["size_field"], S["field"], bold=True)
        elif kind == "pair":
            en, cn = el[1], el[2]
            if en:
                p = P(1, 1, left=0.3, line=1.12)
                rich(p, en, S["size_pair_en"], S["ink"])
            if cn:
                p = P(0, 4, left=0.3, line=1.0)
                rich(p, cn, S["size_pair_cn"], S["grey"])
        elif kind == "bpair":
            en, cn = el[1], el[2]
            p = P(0, 3, left=0.75, hanging=0.35, line=1.1)
            style_run(p.add_run("• "), S["size_bullet"], S["ink"])
            rich(p, en, S["size_bullet"], S["ink"])
            if cn:
                p2 = P(0, 3, left=0.75, line=1.0)
                rich(p2, cn, S["size_pair_cn"], S["grey"])
        elif kind == "bcn":
            p = P(0, 2, left=0.5, hanging=0.35, line=1.1)
            rich(p, "· " + el[1], S["size_bullet_cn"], S["ink_body"])
        elif kind == "para":
            t = el[1]
            is_stage = t.startswith("（") and t.endswith("）")
            p = P(2, 3, line=1.15)
            rich(p, t, S["size_pair_cn"] if is_stage else S["size_body"],
                 S["stage"] if is_stage else S["ink_body"], italic=is_stage)
        elif kind == "lines":
            for _ in range(el[1]):
                rule_line()
            P(0, 2).add_run("")
        elif kind == "table":
            add_table(el[1])
        elif kind == "warn":
            warn_box(el[1], el[2])

    doc.save(out_path)


# ================================================================ HTML -> PDF

HTML_TEMPLATE = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>{title}</title>
<style>
@page {{ size: A4; margin: {html_page_margin}; }}
* {{ box-sizing: border-box; }}
body {{ font-family: "{font_en}", Arial, "{font_cn}", sans-serif;
       font-size: {size_body}pt; color: #{ink_body}; margin: 0; line-height: 1.35; }}
h1 {{ font-size: {size_h1}pt; color: #{accent}; text-align: center; margin: 0 0 4pt; }}
h2 {{ font-size: {size_h2}pt; color: #{accent}; background: #{head_bg}; padding: 5pt 7pt;
     margin: 0 0 8pt; page-break-before: always; }}
h2.first {{ page-break-before: auto; }}
h3 {{ font-size: {size_h3}pt; color: #{accent}; margin: 14pt 0 5pt;
     border-left: 4pt solid #{h3_bar}; padding-left: 6pt; }}
h4 {{ font-size: {size_h4}pt; color: #{h4_ink}; margin: 12pt 0 4pt; }}
h5 {{ font-size: {size_h5}pt; color: #{h5_ink}; margin: 10pt 0 3pt; }}
.f {{ font-size: {size_field}pt; font-weight: 700; color: #{field}; margin: 8pt 0 3pt; }}
.en {{ font-size: {size_pair_en}pt; color: #{ink}; margin: 1pt 0 1pt 3pt; }}
.cn {{ font-size: {size_pair_cn}pt; color: #{grey}; margin: 0 0 5pt 3pt; }}
.bp {{ margin: 0 0 3pt 8pt; font-size: {size_bullet}pt; color: #{ink};
      text-indent: -8pt; padding-left: 8pt; }}
.bc {{ font-size: {size_pair_cn}pt; color: #{grey}; margin: 0 0 3pt 22pt; }}
.bcn {{ font-size: {size_bullet_cn}pt; color: #{ink_body}; margin: 0 0 2pt 6pt;
       text-indent: -7pt; padding-left: 7pt; }}
.p {{ font-size: {size_body}pt; color: #{ink_body}; margin: 2pt 0 3pt; }}
.stage {{ font-size: {size_pair_cn}pt; color: #{stage}; font-style: italic; margin: 2pt 0 3pt; }}
.warn {{ background: #{warn_bg}; border-left: 3pt solid #{warn_bar}; padding: 5pt 8pt 6pt;
        margin: 6pt 0 8pt; font-size: {size_warn}pt; color: #{warn_body}; }}
.warn .t {{ font-weight: 700; color: #{warn_ink}; margin-bottom: 3pt; }}
.warn p {{ margin: 0 0 3pt; }}
.warn ul {{ margin: 0 0 3pt; padding-left: 14pt; }}
.warn li {{ margin-bottom: 1pt; }}
.rule {{ border-bottom: 1px solid #{rule_color}; height: {rule_height_pt}pt; }}
table {{ border-collapse: collapse; width: 100%; margin: 4pt 0 8pt; font-size: {size_table}pt; }}
th, td {{ border: 0.6pt solid #{table_border}; padding: 3pt 5pt; text-align: left; }}
th {{ background: #{head_bg}; color: #{accent}; }}
.warn table {{ border-color: #{warn_table_hair}; }}
.warn th {{ background: #{warn_head_bg}; color: #{warn_ink}; }}
.tall td {{ height: {rule_height_pt}pt; }}
</style></head><body>
{html_body}
</body></html>"""


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inline(s):
    s = esc(s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"<em>\1</em>", s)
    return s


def build_html(els, cfg):
    out = []
    first_h2 = True
    for el in els:
        kind = el[0]
        if kind == "h1":
            out.append("<h1>%s</h1>" % inline(el[1]))
        elif kind == "h2":
            out.append('<h2%s>%s</h2>' % (' class="first"' if first_h2 else "", inline(el[1])))
            first_h2 = False
        elif kind == "h3":
            out.append("<h3>%s</h3>" % inline(el[1]))
        elif kind == "h4":
            out.append("<h4>%s</h4>" % inline(el[1]))
        elif kind == "h5":
            out.append("<h5>%s</h5>" % inline(el[1]))
        elif kind == "field":
            out.append('<div class="f">%s</div>' % inline(el[1]))
        elif kind == "pair":
            if el[1]:
                out.append('<div class="en">%s</div>' % inline(el[1]))
            if el[2]:
                out.append('<div class="cn">%s</div>' % inline(el[2]))
        elif kind == "bpair":
            out.append('<div class="bp">&#8226; %s</div>' % inline(el[1]))
            if el[2]:
                out.append('<div class="bc">%s</div>' % inline(el[2]))
        elif kind == "bcn":
            out.append('<div class="bcn">&#183; %s</div>' % inline(el[1]))
        elif kind == "para":
            t = el[1]
            cls = "stage" if (t.startswith("（") and t.endswith("）")) else "p"
            out.append('<div class="%s">%s</div>' % (cls, inline(t)))
        elif kind == "lines":
            out.append('<div class="rule"></div>' * el[1])
        elif kind == "table":
            out.append(table_html(el[1]))
        elif kind == "warn":
            out.append(warn_html(el[1], el[2]))
    style = dict(cfg["style"])
    return HTML_TEMPLATE.format(title=esc(cfg["title"]), html_body="\n".join(out), **style)


def table_html(rows):
    if not rows:
        return ""
    body = ['<table class="tall">']
    for ri, row in enumerate(rows):
        tag = "th" if ri == 0 else "td"
        body.append("<tr>" + "".join("<%s>%s</%s>" % (tag, inline(c), tag) for c in row) + "</tr>")
    body.append("</table>")
    return "\n".join(body)


def warn_html(title, body):
    parts = ['<div class="warn">', '<div class="t">&#9650; %s</div>' % inline(title)]
    bullets, boxed, rules = [], [], 0
    for line in body:
        if line.startswith("- "):
            bullets.append("<li>%s</li>" % inline(line[2:].strip()))
        elif line.startswith("[["):
            rules += int(re.search(r"\d+", line).group())
        elif line.startswith("|"):
            if not re.match(r"^\|[\s:|-]+\|$", line):
                boxed.append([c.strip() for c in line.strip("|").split("|")])
        elif line:
            parts.append("<p>%s</p>" % inline(line))
    if bullets:
        parts.append("<ul>" + "".join(bullets) + "</ul>")
    if boxed:
        parts.append(table_html(boxed))
    if rules:
        parts.append('<div class="rule"></div>' * rules)
    parts.append("</div>")
    return "\n".join(parts)


def html_to_pdf(html, cfg, out_pdf, log):
    tmp = cfg["tmp_dir"]
    os.makedirs(tmp, exist_ok=True)
    html_path = os.path.join(tmp, "handout_print.html")
    pdf_path = os.path.join(tmp, "handout_print.pdf")
    io.open(html_path, "w", encoding="utf-8", newline="\n").write(html)
    log.append("HTML 中转：%s（%d 字节）" % (html_path, len(html.encode("utf-8"))))

    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    cmd = [cfg["edge"], "--headless=new", "--disable-gpu", "--no-first-run",
           "--no-pdf-header-footer", "--user-data-dir=" + os.path.join(tmp, "edge_profile"),
           "--print-to-pdf=" + pdf_path, "file:///" + html_path.replace("\\", "/")]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=180)
    log.append("Edge 退出码：%s" % proc.returncode)
    tail = proc.stdout.decode("utf-8", "replace")[-400:]
    if tail.strip():
        log.append("Edge 输出：%s" % tail)
    if not os.path.exists(pdf_path):
        log.append("PDF 未生成（Edge 没吐出文件）")
        return False
    shutil.copyfile(pdf_path, out_pdf)
    log.append("PDF：%s（%d 字节）" % (out_pdf, os.path.getsize(out_pdf)))
    return True


# ================================================================ 保真校验

def run_verify(cfg, log):
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "verify_fidelity.py")
    if not os.path.exists(script):
        log.append("保真校验：跳过（同目录下没有 verify_fidelity.py）")
        return None
    cmd = [sys.executable, script, "--check", cfg["md"],
           "--report", os.path.join(cfg["tmp_dir"], "verify_report.txt")]
    for r in cfg["refs"]:
        cmd += ["--ref", r]
    for fmt in cfg["formats"]:
        if fmt == "docx":
            cmd += ["--docx", os.path.join(cfg["out_dir"], cfg["stem"] + ".docx")]
    if cfg["q_range"]:
        cmd += ["--q-range", cfg["q_range"]]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    out = proc.stdout.decode("utf-8", "replace").strip()
    log.append("保真校验：退出码 %s" % proc.returncode)
    if out:
        log.append(out)
    if os.path.exists(os.path.join(cfg["tmp_dir"], "verify_report.txt")):
        log.append("校验报告：%s" % os.path.join(cfg["tmp_dir"], "verify_report.txt"))
    return proc.returncode


# ================================================================ 主流程

def main(argv=None):
    args = parse_args(argv)
    cfg = load_config(args)
    os.makedirs(cfg["tmp_dir"], exist_ok=True)
    os.makedirs(cfg["out_dir"], exist_ok=True)
    log = []
    log.append("输入：%s" % cfg["md"])
    log.append("输出目录：%s" % cfg["out_dir"])
    log.append("输出文件名：%s" % cfg["stem"])

    els = Parser(cfg).parse(cfg["md"])
    counts = {}
    for e in els:
        counts[e[0]] = counts.get(e[0], 0) + 1
    log.append("元素统计：" + ", ".join("%s=%d" % kv for kv in sorted(counts.items())))
    h2 = [e[1] for e in els if e[0] == "h2"]
    log.append("分页锚点（二级标题）：" + " | ".join(h2))

    ok = True
    if "pdf" in cfg["formats"]:
        ok = html_to_pdf(build_html(els, cfg), cfg, os.path.join(cfg["out_dir"], cfg["stem"] + ".pdf"), log) and ok
    if "docx" in cfg["formats"]:
        out_docx = os.path.join(cfg["out_dir"], cfg["stem"] + ".docx")
        build_docx(els, cfg, out_docx)
        log.append("DOCX：%s（%d 字节）" % (out_docx, os.path.getsize(out_docx)))

    if cfg["verify"]:
        if not cfg["refs"]:
            log.append("保真校验：开了开关但没给源头文件（--ref 或配置里的 refs），已跳过")
            ok = False
        else:
            rc = run_verify(cfg, log)
            if rc not in (0, None):
                ok = False
    else:
        log.append("保真校验：未开启（需要时用 --ref 给出源头 md）")

    io.open(cfg["report"], "w", encoding="utf-8", newline="\r\n").write("\n".join(log))
    if not args.quiet:
        sys.stdout.write("\n".join(log) + "\n")
    return 0 if ok else 1


def parse_args(argv):
    ap = argparse.ArgumentParser(
        prog="render_handout",
        description="把结构化 markdown 渲染成 A4 打印用的 PDF 与可编辑的 docx")
    ap.add_argument("--md", help="输入 markdown 的路径")
    ap.add_argument("--config", help="JSON 配置文件；命令行参数优先于配置")
    ap.add_argument("--out-dir", help="输出目录，默认与输入同目录")
    ap.add_argument("--stem", help="输出文件名（不含扩展名），默认取输入文件名")
    ap.add_argument("--title", help="文档标题，用于 PDF 的页面标题，默认取 stem")
    ap.add_argument("--formats", help="要出的格式，逗号分隔：docx,pdf（默认两个都出）")
    ap.add_argument("--lines", type=int, help="「观察记录点」底下补几条手写横线，默认 4")
    ap.add_argument("--lines-rule", action="append", default=[], metavar="前缀=条数",
                    help="某一节的例外条数，可重复。例如「环节 4=8」")
    ap.add_argument("--ref", action="append", default=[], metavar="PATH",
                    help="保真校验的源头文件，可重复。给了即开启校验；"
                         "文件名以 _EN 结尾的按英文源处理，也可以用 en: / cn: 前缀指定")
    ap.add_argument("--verify", action="store_true", help="按配置里的 refs 开启保真校验")
    ap.add_argument("--q-range", metavar="A-B", help="校验题目编号是否连续，例如 2-33")
    ap.add_argument("--report", help="报告文件路径，默认写到中转目录")
    ap.add_argument("--tmp-dir", help="纯 ASCII 中转目录，HTML 与 PDF 先落在这里")
    ap.add_argument("--edge", help="msedge.exe 的完整路径")
    ap.add_argument("--style", action="append", default=[], metavar="key=value",
                    help="覆盖版式，如 size_body=10（可重复）")
    ap.add_argument("--quiet", action="store_true", help="不往标准输出打报告")
    return ap.parse_args(argv)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        import traceback
        tb = traceback.format_exc()
        try:
            os.makedirs(r"C:\Agent\_handout_tmp", exist_ok=True)
            io.open(r"C:\Agent\_handout_tmp\render_traceback.txt", "w",
                    encoding="utf-8", newline="\r\n").write(tb)
        except Exception:
            pass
        sys.stderr.write(tb)
        sys.exit(2)
