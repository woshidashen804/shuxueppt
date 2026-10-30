# -*- coding: utf-8 -*-
"""Turn split exponent runs into real OMML superscripts.

Pattern handled:
    text run "…x"  +  text run "3/2" (baseline=30000)
becomes
    text run "…"   +  <a14:m><m:oMath><m:sSup>x^(3/2)</m:sSup></m:oMath></a14:m>

The exponent text is copied verbatim into m:sup, so e^(−x), e^(sin x) and
x^(3/2) all keep their intended reading. A fraction exponent becomes a real
stacked m:f inside the superscript.

The run element inside oMath must be m:r with m:t. Using a:r there makes
PowerPoint refuse to open the file, which is checked after writing.

Usage:
    python fix_exponent.py --decks DIR [--write]
"""
import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from pptlib import A, M, A14, P, deck_files

A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

RE_EXP_OK = re.compile(r'^[\w\u0370-\u03ff\u0394\u03b4\u2212+\-/.\s()'
                       r'\u00b2\u00b3\u207f\u25a1]{1,16}$')
RE_BASE_CHAR = re.compile(r'([A-Za-z\u0370-\u03ff\u0394\u03b4\)])$')
RE_FRACTION = re.compile(r'^(-?)([0-9.]+|[A-Za-z\u0391-\u03c9\u25a1])'
                         r'\s*/\s*([0-9.]+|[A-Za-z\u0391-\u03c9\u25a1])$')


def mk_math_run(text, size, color, italic=False):
    """<m:r><a:rPr/><m:t>text</m:t></m:r> — a:r here would break the file."""
    run = etree.Element(M + 'r')
    rpr = etree.SubElement(run, A + 'rPr')
    rpr.set('lang', 'zh-CN')
    if size:
        rpr.set('sz', str(size))
    rpr.set('i', '1' if italic else '0')
    rpr.set('dirty', '0')
    fill = etree.SubElement(rpr, A + 'solidFill')
    etree.SubElement(fill, A + 'srgbClr').set('val', color or '1F1F1F')
    for tag in ('latin', 'ea', 'cs'):
        etree.SubElement(rpr, A + tag).set('typeface', 'Cambria Math')
    t = etree.SubElement(run, M + 't')
    t.text = text
    return run


def build_superscript(base_char, exponent, size, color):
    a14 = etree.Element(A14 + 'm', nsmap={'a14': A14_NS, 'm': M_NS})
    omath = etree.SubElement(a14, M + 'oMath', nsmap={'m': M_NS})
    sup = etree.SubElement(omath, M + 'sSup')
    e = etree.SubElement(sup, M + 'e')
    e.append(mk_math_run(base_char, size, color, italic=True))
    sp = etree.SubElement(sup, M + 'sup')

    frac = RE_FRACTION.match(exponent.strip())
    if frac:
        f = etree.SubElement(sp, M + 'f')
        fpr = etree.SubElement(f, M + 'fPr')
        etree.SubElement(fpr, M + 'type').set(M + 'val', 'bar')
        num = etree.SubElement(f, M + 'num')
        num.append(mk_math_run((frac.group(1) or '') + frac.group(2), size, color))
        den = etree.SubElement(f, M + 'den')
        den.append(mk_math_run(frac.group(3), size, color))
    else:
        sp.append(mk_math_run(exponent.replace('\u2212', '-'), size, color))
    return a14


def find_size_and_color(*rprs):
    size, color = None, '1F1F1F'
    for rpr in rprs:
        if rpr is None:
            continue
        if size is None and rpr.get('sz'):
            size = rpr.get('sz')
        c = rpr.find(A + 'solidFill/' + A + 'srgbClr')
        if c is not None:
            color = c.get('val')
            break
    return size, color


def convert(root):
    done = skipped = 0
    for para in list(root.iter(A + 'p')):
        kids = [k for k in para if isinstance(k.tag, str)]
        i = 0
        while i < len(kids):
            run = kids[i]
            if run.tag != A + 'r':
                i += 1
                continue
            rpr = run.find(A + 'rPr')
            if rpr is None or rpr.get('baseline') != '30000':
                i += 1
                continue
            t = run.find(A + 't')
            exp = (t.text or '').strip()
            if not exp or not RE_EXP_OK.match(exp):
                skipped += 1
                i += 1
                continue
            prev = kids[i - 1] if i >= 1 else None
            if prev is None or prev.tag != A + 'r':
                skipped += 1
                i += 1
                continue
            ppr = prev.find(A + 'rPr')
            if ppr is not None and ppr.get('baseline'):
                skipped += 1
                i += 1
                continue
            pt = prev.find(A + 't')
            ptxt = (pt.text or '') if pt is not None else ''
            m = RE_BASE_CHAR.search(ptxt)
            if not m:
                skipped += 1
                i += 1
                continue

            size, color = find_size_and_color(ppr, rpr)
            pt.text = ptxt[:m.start()]
            node = build_superscript(m.group(1), exp, size, color)
            idx = list(para).index(run)
            para.insert(idx, node)
            para.remove(run)
            done += 1

            kids = [k for k in para if isinstance(k.tag, str)]
            i = kids.index(node) + 1
    return done, skipped


def validate_math_runs(data):
    """Ensure no a:r / a:t leaked inside oMath (PowerPoint would reject it)."""
    root = etree.fromstring(data)
    for om in root.iter(M + 'oMath'):
        if any(True for _ in om.iter(A + 'r')):
            return False
        if any(True for _ in om.iter(A + 't')):
            return False
    for mr in root.iter(M + 'r'):
        if mr.find(A + 't') is not None:
            return False
    return True


def process(path, write):
    zin = zipfile.ZipFile(path)
    tmp = tempfile.mkdtemp()
    out = os.path.join(tmp, 'o.zip')
    done = skipped = 0
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') \
                    and item.filename.endswith('.xml'):
                root = etree.fromstring(data)
                d, s = convert(root)
                if d:
                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
                    if not validate_math_runs(data):
                        raise SystemExit('ABORT %s %s: 公式内出现 a:r'
                                         % (os.path.basename(path), item.filename))
                    done += d
                skipped += s
            zout.writestr(item, data)
    zin.close()
    if write:
        shutil.copyfile(out, path)
    shutil.rmtree(tmp, ignore_errors=True)
    return done, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--write', action='store_true')
    args = ap.parse_args()

    td = ts = 0
    for f in deck_files(args.decks):
        d, s = process(os.path.join(args.decks, f), args.write)
        td += d
        ts += s
        if d or s:
            print('%-44s 转换=%d 跳过=%d' % (f, d, s))
    print('合计 转换=%d 跳过=%d  %s'
          % (td, ts, 'WRITE' if args.write else 'DRY-RUN'))


if __name__ == '__main__':
    main()
