# -*- coding: utf-8 -*-
"""Rebuild lim subscripts as real limit objects.

    <m:sSub><m:e>lim</m:e><m:sub>x→0</m:sub></m:sSub>
 -> <m:limLow><m:e>lim</m:e><m:lim>x→0</m:lim></m:limLow>

Only the case where the base text is exactly "lim" is touched. Everything else
is reported as skipped so nothing is silently rewritten.

Usage:
    python fix_limit.py --decks DIR [--write]
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


def rebuild(root):
    n = 0
    for sub in list(root.iter(M + 'sSub')):
        base, cond = sub.find(M + 'e'), sub.find(M + 'sub')
        if base is None or cond is None:
            continue
        if text_of(base) != 'lim':
            continue

        pr = sub.find(M + 'sSubPr')
        ctrl = pr.find(M + 'ctrlPr') if pr is not None else None

        low = etree.Element(M + 'limLow')
        pr2 = etree.SubElement(low, M + 'limLowPr')
        if ctrl is not None:
            pr2.append(ctrl)
        e2 = etree.SubElement(low, M + 'e')
        for child in list(base):
            e2.append(child)
        lim = etree.SubElement(low, M + 'lim')
        for child in list(cond):
            lim.append(child)

        parent = sub.getparent()
        idx = list(parent).index(sub)
        parent.remove(sub)
        parent.insert(idx, low)
        n += 1
    return n


def process(path, write):
    zin = zipfile.ZipFile(path)
    tmp = tempfile.mkdtemp()
    out = os.path.join(tmp, 'o.zip')
    total = 0
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') \
                    and item.filename.endswith('.xml'):
                root = etree.fromstring(data)
                c = rebuild(root)
                if c:
                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
                    total += c
            zout.writestr(item, data)
    zin.close()
    if write:
        shutil.copyfile(out, path)
    shutil.rmtree(tmp, ignore_errors=True)
    return total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--write', action='store_true')
    args = ap.parse_args()

    grand = 0
    for f in deck_files(args.decks):
        c = process(os.path.join(args.decks, f), args.write)
        grand += c
        if c:
            print('%-44s lim 重建 %d 处' % (f, c))
    print('合计 %d  %s' % (grand, 'WRITE' if args.write else 'DRY-RUN'))


if __name__ == '__main__':
    main()
