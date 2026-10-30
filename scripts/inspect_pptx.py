# -*- coding: utf-8 -*-
"""Inspect a folder of math-course pptx files and report defects.

Pure XML: no PowerPoint or WPS needed, works on any platform. Every finding
carries a code from the defect catalogue so a report can be triaged and a repair
script selected.

Usage:
    python inspect_pptx.py --decks DIR [--out DIR] [--json-only]
"""
import argparse
import collections
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from pptlib import (Deck, A, P, M, A14, MC, deck_files, iter_shapes, shape_paras,
                    shape_name, shape_geo, is_hidden, slide_size, omml_linear,
                    para_segments, para_text, para_plain, shape_text)

# --------------------------------------------------------------- patterns
SUP_CHARS = ('\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079'
             '\u207a\u207b\u207f\u2071\u1d43\u1d47\u1d9c\u1d48\u1d49\u1da0'
             '\u02b0\u02b2\u02b3\u02b7\u02e3\u02e2\u1d57\u1d58\u1d5b')
RE_UNICODE_SUP = re.compile('[%s]+' % re.escape(SUP_CHARS))
RE_HALF_EXP = re.compile('[\u00bd\u00bc\u00be]')
RE_CARET = re.compile(r'\^')
RE_PLACEHOLDER = re.compile('\u5355\u51fb\u6b64\u5904|\u70b9\u51fb\u6b64\u5904'
                            '|TODO|XXX|\u5f85\u8865|\u5360\u4f4d')
# a fraction written with a slash between math-ish tokens in prose
RE_LINEAR_FRAC = re.compile(r'(?<![0-9A-Za-z_/])'
                            r'([0-9A-Za-z\u0391-\u03c9][0-9A-Za-z\u0391-\u03c9]{0,6})'
                            r'\s*/\s*'
                            r'([0-9A-Za-z\u0391-\u03c9(][0-9A-Za-z\u0391-\u03c9)]{0,6})')
FRAC_ALLOWED = {'0/0', '\u221e/\u221e', '1/2', '1/3', '1/4'}
RE_QNUM = re.compile(r'^\s*(\d{1,2})\s*[.\uff0e]\s*(?!\d)\S')

ISSUES = []


def add(deck, page, shape, family, code, severity, detail, sample):
    ISSUES.append(dict(deck=deck, page=page, shape=shape, family=family,
                       code=code, severity=severity,
                       detail=detail, sample=(sample or '')[:160]))


# ------------------------------------------------------- alternate content
def check_alternate_content(deck_name, page, root):
    """M-BLANKFB: content lives only in mc:Choice, mc:Fallback is an empty shell.

    Any consumer that takes the fallback branch shows nothing at all. WPS does
    this in practice while PowerPoint reads Choice, which is why the slide looks
    fine in one editor and blank in the other.
    """
    for ac in root.iter(MC + 'AlternateContent'):
        choice = ac.find(MC + 'Choice')
        fallback = ac.find(MC + 'Fallback')
        if choice is None:
            continue
        ct = ''.join((e.text or '') for e in choice.iter(A + 't')).strip()
        cm = ''.join((e.text or '') for e in choice.iter(M + 't')).strip()
        if not (ct or cm):
            continue
        ft = ''.join((e.text or '') for e in fallback.iter(A + 't')).strip() \
            if fallback is not None else ''
        fm = ''.join((e.text or '') for e in fallback.iter(M + 't')).strip() \
            if fallback is not None else ''
        if ft or fm:
            continue
        parent = etree.QName(ac.getparent()).localname
        if fallback is None:
            add(deck_name, page, parent, 'INVISIBLE', 'M-NOFALLBACK', 'P0',
                'AlternateContent 无 Fallback 分支，非 a14 阅读器整块空白',
                (ct or cm))
        else:
            add(deck_name, page, parent, 'INVISIBLE', 'M-BLANKFB', 'P0',
                'Fallback 分支为空，WPS 会整块不显示（PowerPoint 正常）',
                (ct or cm))


# ------------------------------------------------------------ omml nesting
def check_omml_nesting(deck_name, page, root):
    """M-BADNEST: <m:oMathPara> may only hold m:oMathParaPr and m:oMath."""
    for para in root.iter(M + 'oMathPara'):
        if any(c.tag == A14 + 'm' for c in para if isinstance(c.tag, str)):
            txt = ''.join((e.text or '') for e in para.iter(M + 't'))
            add(deck_name, page, 'oMathPara', 'INVISIBLE', 'M-BADNEST', 'P0',
                'm:oMathPara 内非法嵌套 a14:m，独立公式不渲染', txt)


# ---------------------------------------------------------------- math body
def _empty_container(node):
    if ''.join((t.text or '') for t in node.iter(M + 't')).strip():
        return False
    for child in node:
        if not isinstance(child.tag, str):
            continue
        if etree.QName(child).localname.endswith('Pr'):
            continue
        return False
    return True


def check_math(deck_name, page, sname, para):
    for seg in para_segments(para):
        if seg['kind'] != 'math':
            continue
        om = seg['el']
        flat = seg['text']

        if not ''.join((t.text or '') for t in om.iter(M + 't')).strip():
            add(deck_name, page, sname, 'EMPTY', 'M-EMPTYMATH', 'P0',
                '空公式对象（无任何 m:t 文本）', flat)

        # function name with no argument
        for fn in om.iter(M + 'func'):
            name = fn.find(M + 'fName')
            arg = fn.find(M + 'e')
            nt = ''.join((t.text or '') for t in name.iter(M + 't')).strip() \
                if name is not None else ''
            if arg is not None and _empty_container(arg):
                add(deck_name, page, sname, 'EMPTY', 'M-FUNCARG', 'P0',
                    '函数 %s 缺少自变量' % (nt or '?'), flat)

        # fraction with an empty side
        for fr in om.iter(M + 'f'):
            num, den = fr.find(M + 'num'), fr.find(M + 'den')
            nt = ''.join((t.text or '') for t in num.iter(M + 't')).strip() \
                if num is not None else ''
            dt = ''.join((t.text or '') for t in den.iter(M + 't')).strip() \
                if den is not None else ''
            if num is None or den is None or not nt or not dt:
                add(deck_name, page, sname, 'EMPTY', 'M-EMPTYFRAC', 'P0',
                    '分式分子或分母为空（分子=%r 分母=%r）' % (nt, dt), flat)

        # limit written as a plain subscript
        for sub in om.iter(M + 'sSub'):
            base = sub.find(M + 'e')
            bt = ''.join((t.text or '') for t in base.iter(M + 't')).strip() \
                if base is not None else ''
            if bt == 'lim':
                node = sub.find(M + 'sub')
                cond = ''.join((t.text or '') for t in node.iter(M + 't')).strip() \
                    if node is not None else ''
                add(deck_name, page, sname, 'LIMIT', 'L-AS-SUB', 'P1',
                    'lim 用普通下标搭建，趋近条件会缩成右下脚标（应为 m:limLow）',
                    'lim_(%s)' % cond)

        # limit object missing its condition or operand
        for node in om.iter(M + 'limLow'):
            lim, base = node.find(M + 'lim'), node.find(M + 'e')
            lt = ''.join((t.text or '') for t in lim.iter(M + 't')).strip() \
                if lim is not None else ''
            bt = ''.join((t.text or '') for t in base.iter(M + 't')).strip() \
                if base is not None else ''
            if lim is None or base is None or not lt or not bt:
                add(deck_name, page, sname, 'LIMIT', 'L-INCOMPLETE', 'P0',
                    'lim 结构残缺（条件=%r 被求极限的式子=%r）' % (lt, bt), flat)

        # a slash left inside math text where a fraction belongs
        for run in om.iter(M + 'r'):
            t = run.find(M + 't')
            if t is None or not t.text:
                continue
            txt = t.text
            if not re.search(r'[0-9A-Za-z\u0391-\u03c9)\]]\s*/\s*'
                             r'[0-9A-Za-z\u0391-\u03c9(]', txt):
                continue
            norm = txt.strip().strip('\u00a0')
            if norm in ('0/0', '\u221e/\u221e'):
                continue
            anc = [etree.QName(a).localname for a in run.iterancestors()]
            if any(x in anc for x in ('sSup', 'sSub', 'sSubSup')):
                continue      # a/b as an exponent is conventional
            add(deck_name, page, sname, 'FRAC', 'F-IN-TEXT', 'P1',
                '公式内部仍以斜杠写分式：%r' % txt[:40], flat)

        # whole expression used as a superscript
        for node in om.iter(M + 'sSup'):
            sup = node.find(M + 'sup')
            if sup is None:
                continue
            st = ''.join((t.text or '') for t in sup.iter(M + 't')).strip()
            if ('/' in st or ' ' in st) and len(st) > 3:
                add(deck_name, page, sname, 'EXP', 'E-SUPEXPR', 'P1',
                    '整个式子被做成上标（%r），版面会挤成一小块' % st, flat)


# --------------------------------------------------------------- plain text
def check_text(deck_name, page, sname, para):
    segs = para_segments(para)
    full = para_text(para)
    for i, seg in enumerate(segs):
        if seg['kind'] != 'text':
            continue
        t = seg['text']
        if not t.strip():
            continue

        # exponent split into a separate run: "…e" + "−x"  /  "…x" + "3/2"
        if i + 1 < len(segs) and segs[i + 1]['kind'] == 'text':
            nxt = segs[i + 1]['text']
            tail = t.rstrip()[-1:]
            if tail in 'eE' and re.match(r'^[\u2212+\-][0-9A-Za-z]', nxt.lstrip()):
                add(deck_name, page, sname, 'EXP', 'E-SPLITRUN', 'P0',
                    '指数被拆成独立 run（"%s" + "%s"），会排成 e 减 x 而非上标'
                    % (tail, nxt.lstrip()[:10]), full)
            elif tail in 'xXyY' and re.match(r'^-?[0-9]', nxt.lstrip()) and '/' in nxt:
                add(deck_name, page, sname, 'EXP', 'E-SPLITRUN', 'P0',
                    '分数指数被拆成独立 run（"%s" + "%s"），会排成 x3/2'
                    % (tail, nxt.lstrip()[:10]), full)

        if RE_CARET.search(t):
            add(deck_name, page, sname, 'EXP', 'E-CARET', 'P0',
                '正文残留 ^ 线性指数写法', full)
        if RE_HALF_EXP.search(t):
            add(deck_name, page, sname, 'EXP', 'E-HALF', 'P1',
                '分数指数写成 Unicode 半字符（x½ 之类）', full)
        us = RE_UNICODE_SUP.findall(t)
        if us:
            add(deck_name, page, sname, 'EXP', 'E-UNISUP', 'P2',
                '正文用 Unicode 上标字符（%r），与 OMML 上标样式不统一'
                % ''.join(us), full)

        for m in RE_LINEAR_FRAC.finditer(t):
            frag = m.group(0).replace(' ', '')
            if frag in FRAC_ALLOWED or len(frag) > 12:
                continue
            add(deck_name, page, sname, 'FRAC', 'F-LINEAR', 'P1',
                '分式写成线性斜杠 %r，未做成分式对象' % m.group(0), full)

        if RE_PLACEHOLDER.search(t):
            add(deck_name, page, sname, 'EMPTY', 'M-PLACEHOLDER', 'P0',
                '残留占位符文字', full)


# ------------------------------------------------------------------ layout
def check_layout(deck_name, page, root, deck, hidden_unreachable_report):
    width, height = slide_size(deck, page)
    targets = {t.get('spid') for t in root.iter(P + 'spTgt')}
    for sp in iter_shapes(root):
        geo = shape_geo(sp)
        if geo is None:
            continue
        x, y, w, h = geo
        sname = shape_name(sp)
        txt = shape_text(sp)

        if txt and (x >= width - 0.05 or y >= height - 0.05
                    or x + w <= 0.05 or y + h <= 0.05):
            add(deck_name, page, sname, 'LAYOUT', 'O-OFFCANVAS', 'P1',
                '形状整体在画布外（放映不显示）x=%.2f y=%.2f' % (x, y), txt)

        cnv = sp.find('.//' + P + 'cNvPr')
        if cnv is not None and cnv.get('hidden') == '1' and txt:
            sid = cnv.get('id')
            if sid not in targets:
                add(deck_name, page, sname, 'LAYOUT', 'O-HIDDENDEAD', 'P0',
                    '形状标记 hidden 但没有任何动画指向它，永远不显示', txt)


def check_answers(deck_name, page, items):
    """A numbered question must own at least one answer shape in its band.

    Several answers per question (option letter + explanation) are normal, so we
    band by vertical order inside each column instead of pairing 1:1.
    """
    anims = [(n, t, g) for n, t, g in items if n.startswith('anim') and t and g]
    if not anims:
        return
    questions = []
    for n, t, g in items:
        if n.startswith('anim') or not g:
            continue
        m = RE_QNUM.match(t)
        if m and len(t) > 6:
            questions.append((int(m.group(1)), n, t, g))
    if not questions:
        return

    xs = sorted({round(q[3][0], 1) for q in questions})
    edges = []
    for x in xs:
        if not edges or x - edges[-1] > 2.0:
            edges.append(x)

    for ci, edge in enumerate(edges):
        lo = edge - 0.40
        hi = edges[ci + 1] - 0.40 if ci + 1 < len(edges) else 1e9
        col_q = sorted([q for q in questions if lo <= q[3][0] < hi],
                       key=lambda q: q[3][1])
        col_a = [a for a in anims if lo - 1.0 <= a[2][0] < hi + 1.0]
        for i, (qn, qname, qtxt, g) in enumerate(col_q):
            top = g[1] - 0.35
            bottom = col_q[i + 1][3][1] - 0.35 if i + 1 < len(col_q) else g[1] + 2.60
            if not [a for a in col_a if top <= a[2][1] < bottom]:
                add(deck_name, page, qname, 'EMPTY', 'M-NOANSWER', 'P1',
                    '第 %d 题在其版面区间内没有任何答案形状' % qn, qtxt)


# -------------------------------------------------------------------- main
def inspect_deck(path):
    deck = Deck(path)
    name = os.path.basename(path)
    for page, root in deck.slides:
        check_alternate_content(name, page, root)
        check_omml_nesting(name, page, root)
        items = []
        for sp in iter_shapes(root):
            sname = shape_name(sp)
            paras = shape_paras(sp)
            for p in paras:
                check_math(name, page, sname, p)
                check_text(name, page, sname, p)
            items.append((sname, shape_text(sp), shape_geo(sp)))
        check_layout(name, page, root, deck, None)
        if not any(i.startswith('anim') for i, _, _ in items):
            pass
        check_answers(name, page, items)
    deck.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--out', default=None)
    ap.add_argument('--json-only', action='store_true')
    args = ap.parse_args()

    files = deck_files(args.decks)
    if not files:
        raise SystemExit('没有找到 .pptx: %s' % args.decks)
    for f in files:
        inspect_deck(os.path.join(args.decks, f))

    out_dir = args.out or os.getcwd()
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, 'issues.json'), 'w', encoding='utf-8') as fh:
        json.dump(ISSUES, fh, ensure_ascii=False, indent=1)

    counter = collections.Counter((i['family'], i['code']) for i in ISSUES)
    print('共 %d 份，%d 条问题' % (len(files), len(ISSUES)))
    for (family, code), n in sorted(counter.items(), key=lambda kv: -kv[1]):
        print('  %-10s %-14s %d' % (family, code, n))

    if not args.json_only:
        write_report(os.path.join(out_dir, 'report.md'), files, counter)


SEV_LABEL = {'P0': '高', 'P1': '中', 'P2': '低'}


def write_report(path, files, counter):
    by_deck = collections.defaultdict(list)
    for i in ISSUES:
        by_deck[i['deck']].append(i)
    lines = ['# 数学课件检查报告', '',
             '检查 %d 份课件，发现 %d 条问题。' % (len(files), len(ISSUES)), '',
             '## 概览', '', '| 族 | 代码 | 数量 | 严重度 |', '|---|---|---:|---|']
    for (family, code), n in sorted(counter.items(), key=lambda kv: -kv[1]):
        sev = next((i['severity'] for i in ISSUES
                    if i['code'] == code), '')
        lines.append('| %s | %s | %d | %s |' % (family, code, n, SEV_LABEL.get(sev, sev)))
    lines.append('')
    for f in files:
        rows = by_deck.get(f, [])
        lines.append('## %s' % f)
        lines.append('')
        if not rows:
            lines.append('（未发现结构性问题）')
            lines.append('')
            continue
        lines.append('| 页 | 形状 | 代码 | 级别 | 说明 | 片段 |')
        lines.append('|---:|---|---|---|---|---|')
        seen = set()
        for i in sorted(rows, key=lambda r: (r['page'], r['code'])):
            key = (i['page'], i['code'], i['sample'][:40])
            if key in seen:
                continue
            seen.add(key)
            lines.append('| P%02d | %s | %s | %s | %s | %s |' % (
                i['page'], i['shape'][:16], i['code'],
                SEV_LABEL.get(i['severity'], ''), i['detail'].replace('|', '/'),
                i['sample'].replace('|', '/')[:40]))
        lines.append('')
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines) + '\n')
    print('-> %s' % path)


if __name__ == '__main__':
    main()
