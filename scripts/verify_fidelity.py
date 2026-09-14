# -*- coding: utf-8 -*-
"""保真校验 —— 待校文件里的每一句中英双语行，必须能在给定的源头文件里逐字找到。

用途：当一份 markdown 是从别的文件（定稿、需求、旧版）提取重组的，就要证明它
零改写、零新增。做法是把待校文件里的双语行规范化之后，拿到源头文件的全文里做
子串查找；找不到的连行号一起报出来。

它只查「找不到」，不查「意思变了」。同义替换、语序调整、语义走样一律看不出来，
只有逐字照抄才能过关。

用法：
    python verify_fidelity.py --check 手册.md \\
        --ref 定稿_中文.md --ref 定稿_EN.md \\
        --docx 手册.docx --q-range 2-33 \\
        --report verify_report.txt

退出码：0 = 全部找到；1 = 有找不到的行或开了 --q-range 而题号不连续；2 = 出错。
"""

import argparse
import io
import os
import re
import sys

DEFAULT_REPORT = "verify_report.txt"


def norm(s):
    s = s.replace("　", " ")
    s = re.sub(r"[*`]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def source_blob(path):
    raw = io.open(path, encoding="utf-8").read().replace("\ufeff", "")
    lines = []
    for ln in raw.split("\n"):
        st = ln.strip()
        if st.startswith(">"):
            st = st[1:].strip()
        lines.append(st)
    return norm(" ".join(lines))


def classify_ref(spec):
    """返回 (语言, 路径)。语言取 cn / en；没写前缀时按文件名 _EN 后缀判定。"""
    m = re.match(r"^(cn|en)\s*[:：]\s*(.+)$", spec, re.I)
    if m:
        return m.group(1).lower(), m.group(2).strip()
    path = spec.strip()
    stem = os.path.splitext(os.path.basename(path))[0]
    return ("en" if stem.upper().endswith("_EN") else "cn"), path


LINE_RE = re.compile(r"^(?:>\s*|-\s*)\*\*(EN|中文)\*\*[　\s]*(.*)$")


def check(check_path, refs, docx_path, q_range):
    sources = {"cn": [], "en": []}
    notes = []
    for spec in refs:
        lang, path = classify_ref(spec)
        if not os.path.exists(path):
            raise SystemExit("找不到源头文件：%s" % path)
        sources[lang].append(source_blob(path))
        notes.append("%s → %s" % (path, "英文源" if lang == "en" else "中文源"))

    blob = {}
    for lang in ("cn", "en"):
        if sources[lang]:
            blob[lang] = "  ".join(sources[lang])
        else:
            # 只给了一侧源头时，退回到「所有源头合起来」，并在报告里说明
            blob[lang] = "  ".join(sources["cn"] + sources["en"])
            notes.append("没有单独的%s源头，该语言的双语行改为在全部源头里查找"
                         % ("英文" if lang == "en" else "中文"))

    out, missing, counts = [], [], {"cn": 0, "en": 0}
    text = io.open(check_path, encoding="utf-8").read()
    for i, ln in enumerate(text.split("\n"), 1):
        m = LINE_RE.match(ln.strip())
        if not m:
            continue
        lang = "en" if m.group(1) == "EN" else "cn"
        body = norm(m.group(2))
        if not body:
            continue
        counts[lang] += 1
        if body not in blob[lang]:
            missing.append((i, m.group(1), m.group(2)[:90]))

    out.append("待校文件：%s" % check_path)
    out.append("源头文件：")
    for n in notes:
        out.append("    " + n)
    out.append("")
    out.append("双语行：英文 %d 条，中文 %d 条" % (counts["en"], counts["cn"]))
    if missing:
        out.append("")
        out.append("★ 在源头里找不到原文的行（%d 条）：" % len(missing))
        for i, lang, s in missing:
            out.append("   行 %d [%s]：%s" % (i, lang, s))
    else:
        out.append("")
        out.append("✔ 全部双语行都能在源头里逐字找到，无改写、无新增。")

    bad_q = []
    if q_range:
        lo, hi = [int(x) for x in re.split(r"[-–~]", q_range)]
        qs = sorted(set(int(x) for x in re.findall(r"^#{2,5}\s*Q?(\d+)\s*[·．.]", text, re.M)))
        missing_q = [q for q in range(lo, hi + 1) if q not in qs]
        out.append("")
        out.append("题目编号：%s" % (",".join("Q%d" % q for q in qs) or "无"))
        out.append("期望 Q%d-Q%d 共 %d 道，实得 %d 道，缺失：%s"
                   % (lo, hi, hi - lo + 1, len([q for q in qs if lo <= q <= hi]),
                      ",".join("Q%d" % q for q in missing_q) or "无"))
        bad_q = missing_q

    if docx_path:
        out.append("")
        if os.path.exists(docx_path):
            try:
                from docx import Document
                d = Document(docx_path)
                pbb = sum(1 for p in d.paragraphs if p.paragraph_format.page_break_before)
                out.append("docx：段落 %d 个、表格 %d 张、分页前段落 %d 处"
                           % (len(d.paragraphs), len(d.tables), pbb))
            except Exception as e:
                out.append("docx 检查失败：%r" % (e,))
        else:
            out.append("docx 不存在：%s" % docx_path)

    return "\n".join(out), (len(missing) == 0 and not bad_q)


def parse_args(argv):
    ap = argparse.ArgumentParser(prog="verify_fidelity",
                                 description="逐字回查：待校文件的双语行是否原样来自源头文件")
    ap.add_argument("--check", required=True, help="待校验的 markdown")
    ap.add_argument("--ref", action="append", default=[], metavar="PATH",
                    help="源头文件，可重复。文件名以 _EN 结尾的按英文源处理，"
                         "也可以用 en: / cn: 前缀指定")
    ap.add_argument("--docx", help="顺带统计这份 docx 的结构")
    ap.add_argument("--q-range", metavar="A-B", help="校验题号是否连续，例如 2-33")
    ap.add_argument("--report", help="报告写到哪，默认与待校文件同目录")
    ap.add_argument("--quiet", action="store_true")
    return ap.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if not args.ref:
        raise SystemExit("至少要给一个 --ref")
    report, ok = check(args.check, args.ref, args.docx, args.q_range)
    path = args.report or os.path.join(os.path.dirname(os.path.abspath(args.check)),
                                       DEFAULT_REPORT)
    if os.path.dirname(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
    io.open(path, "w", encoding="utf-8", newline="\r\n").write(report)
    if not args.quiet:
        sys.stdout.write(report + "\n报告：%s\n" % path)
    return 0 if ok else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        import traceback
        sys.stderr.write(traceback.format_exc())
        sys.exit(2)
