# -*- coding: utf-8 -*-
"""Structural integrity check that needs no PowerPoint.

Two invariants:

1. Every slide XML parses, and no shape is left with a dangling animation target.
2. The <m:oMath> count per slide never drops below the baseline.

That second one is the important lesson: a14:m count is NOT an invariant, because
flattening an illegal nest legitimately removes one. Using a14:m as the guard hides
real math loss.

Usage:
    python verify_integrity.py --decks DIR [--baseline DIR]
"""
import argparse
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from pptlib import P, deck_files

RE_OMATH = re.compile(r'<m:oMath[ >]')


def slide_map(path):
    """page -> (oMath count, slide xml length)."""
    zf = zipfile.ZipFile(path)
    out = {}
    for name in zf.namelist():
        if not (name.startswith('ppt/slides/slide') and name.endswith('.xml')):
            continue
        page = int(''.join(c for c in name.split('slide')[-1] if c.isdigit()))
        raw = zf.read(name).decode('utf-8')
        out[page] = len(RE_OMATH.findall(raw))
    zf.close()
    return out


def structural_problems(path):
    """Parse failures and animation targets that point at nothing."""
    problems = []
    zf = zipfile.ZipFile(path)
    for name in zf.namelist():
        if not (name.startswith('ppt/slides/slide') and name.endswith('.xml')):
            continue
        page = int(''.join(c for c in name.split('slide')[-1] if c.isdigit()))
        try:
            root = etree.fromstring(zf.read(name))
        except Exception as exc:
            problems.append('P%02d XML 解析失败：%s' % (page, exc))
            continue
        ids = {c.get('id') for c in root.iter(P + 'cNvPr')}
        targets = {t.get('spid') for t in root.iter(P + 'spTgt')}
        orphan = sorted(t for t in targets if t and t not in ids)
        if orphan:
            problems.append('P%02d 动画指向不存在的形状：%s' % (page, orphan))
    zf.close()
    return problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--baseline', default=None)
    args = ap.parse_args()

    bad = 0
    for f in deck_files(args.decks):
        cur_path = os.path.join(args.decks, f)
        cur = slide_map(cur_path)
        issues = structural_problems(cur_path)

        lost = []
        if args.baseline:
            base_path = os.path.join(args.baseline, f)
            if os.path.exists(base_path):
                base = slide_map(base_path)
                lost = [(p, base[p], cur.get(p, 0))
                        for p in sorted(base)
                        if cur.get(p, 0) < base[p]]
                missing = sorted(set(base) - set(cur))
                if missing:
                    issues.append('缺失页：%s' % missing)
            else:
                issues.append('基线中找不到同名文件')

        status = 'OK' if not issues and not lost else 'FAIL'
        if status == 'FAIL':
            bad += 1
        print('%-44s %-4s oMath=%d' % (f, status, sum(cur.values())))
        for i in issues:
            print('     %s' % i)
        for p, b, c in lost:
            print('     P%02d 公式减少 %d -> %d' % (p, b, c))

    print('\n问题课件数：%d' % bad)
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()
