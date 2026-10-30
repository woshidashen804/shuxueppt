# -*- coding: utf-8 -*-
"""Remove empty math wrappers that render as a placeholder box.

Two cases:

1. m:func whose m:e is empty — "sin" / "cos" / "tan" used as a noun rather than
   applied to something. The function name is spliced back as plain math runs so
   the words survive and the empty argument box disappears.

2. m:f with an empty numerator or denominator cannot be repaired without knowing
   the intended values, so it is only reported.

Usage:
    python fix_empty_math.py --decks DIR [--write]
"""
import argparse
import os
import shutil
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from pptlib import A, M, deck_files


def text_of(node):
    return ''.join((t.text or '') for t in node.iter(M + 't')).strip()


def is_empty(node):
    if text_of(node):
        return False
    for child in node:
        if not isinstance(child.tag, str):
            continue
        if etree.QName(child).localname.endswith('Pr'):
            continue
        return False
    return True


def strip_empty_func(root):
    n = 0
    for fn in list(root.iter(M + 'func')):
        arg, name = fn.find(M + 'e'), fn.find(M + 'fName')
        if arg is None or name is None or not is_empty(arg):
            continue
        parent = fn.getparent()
        if parent is None:
            continue
        idx = list(parent).index(fn)
        for j, child in enumerate(list(name)):
            parent.insert(idx + j, child)
        parent.remove(fn)
        n += 1
    return n


def report_empty_frac(root):
    out = []
    for fr in root.iter(M + 'f'):
        num, den = fr.find(M + 'num'), fr.find(M + 'den')
        if num is None or den is None or not text_of(num) or not text_of(den):
            out.append('分子=%r 分母=%r' % (
                text_of(num) if num is not None else None,
                text_of(den) if den is not None else None))
    return out


def process(path, write):
    zin = zipfile.ZipFile(path)
    tmp = tempfile.mkdtemp()
    out = os.path.join(tmp, 'o.zip')
    fixed = 0
    warnings = []
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') \
                    and item.filename.endswith('.xml'):
                root = etree.fromstring(data)
                c = strip_empty_func(root)
                warnings.extend(report_empty_frac(root))
                if c:
                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
                    fixed += c
            zout.writestr(item, data)
    zin.close()
    if write:
        shutil.copyfile(out, path)
    shutil.rmtree(tmp, ignore_errors=True)
    return fixed, warnings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--write', action='store_true')
    args = ap.parse_args()

    total = 0
    for f in deck_files(args.decks):
        c, warn = process(os.path.join(args.decks, f), args.write)
        total += c
        if c:
            print('%-44s 去掉空函数壳 %d 处' % (f, c))
        for w in warn:
            print('  [需人工] %s 空分式：%s' % (f, w))
    print('合计 %d  %s' % (total, 'WRITE' if args.write else 'DRY-RUN'))


if __name__ == '__main__':
    main()
