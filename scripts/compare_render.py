# -*- coding: utf-8 -*-
"""Render decks with PowerPoint and/or WPS and compare ink page by page.

Why this exists: XML inspection cannot tell you whether WPS actually draws a
formula. Two behaviours need measuring separately.

  * PowerPoint delta  (baseline -> current): must stay near zero.
    A large negative delta means the repair changed what users already saw.
  * WPS delta         (baseline -> current): must be positive where content was
    missing. That is the proof that the invisible-content fix worked.

It also separates true deletion from a mere layout shift:
    lost  = dark in baseline, white now
    extra = white in baseline, dark now
A deletion gives lost >> extra; a shift gives lost ~= extra.

Windows only (win32com + Pillow).

Usage:
    python compare_render.py --decks DIR --baseline DIR [--engines ppt,wps]
                             [--pages 5,17] [--out DIR]
"""
import argparse
import os
import shutil
import sys
import tempfile

try:
    import win32com.client as win32
except ImportError:
    raise SystemExit('需要 win32com（pip install pywin32）')
try:
    from PIL import Image
except ImportError:
    raise SystemExit('需要 Pillow（pip install pillow）')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptlib import deck_files

PROGIDS = {'ppt': 'PowerPoint.Application', 'wps': 'KWPP.Application'}


def render_deck(progid, path, outdir, tag, width=1280, height=720):
    os.makedirs(outdir, exist_ok=True)
    app = win32.Dispatch(progid)
    try:
        app.Visible = False
    except Exception:
        pass
    try:
        pres = app.Presentations.Open(path)
    except Exception as exc:
        raise SystemExit('无法打开 %s（%s）：%s' % (os.path.basename(path), tag, exc))
    images = {}
    try:
        for i in range(1, pres.Slides.Count + 1):
            out = os.path.join(outdir, '%s_%03d.png' % (tag, i))
            pres.Slides(i).Export(out, 'PNG', width, height)
            images[i] = out
    finally:
        pres.Close()
        try:
            app.Quit()
        except Exception:
            pass
    return images


def ink(path):
    im = Image.open(path).convert('L')
    px = im.load()
    w, h = im.size
    n = 0
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            if px[x, y] < 200:
                n += 1
    im.close()
    return n


def lost_extra(a_path, b_path):
    a = Image.open(a_path).convert('L')
    b = Image.open(b_path).convert('L')
    if a.size != b.size:
        b = b.resize(a.size)
    pa, pb = a.load(), b.load()
    w, h = a.size
    lost = extra = 0
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            da = pa[x, y] < 200
            db = pb[x, y] < 200
            if da and not db:
                lost += 1
            elif db and not da:
                extra += 1
    a.close()
    b.close()
    return lost * 4, extra * 4


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--baseline', required=True)
    ap.add_argument('--engines', default='ppt,wps')
    ap.add_argument('--pages', default=None,
                    help='只比较这些页，例如 5,17（默认全部）')
    ap.add_argument('--out', default=None)
    args = ap.parse_args()

    engines = [e.strip() for e in args.engines.split(',') if e.strip()]
    only_pages = None
    if args.pages:
        only_pages = {int(x) for x in args.pages.split(',') if x.strip()}

    tmp = tempfile.mkdtemp()
    summary = []
    for f in deck_files(args.decks):
        cur_path = os.path.join(args.decks, f)
        base_path = os.path.join(args.baseline, f)
        if not os.path.exists(base_path):
            print('%-40s 跳过（基线缺失）' % f)
            continue
        print('### %s' % f)
        for eng in engines:
            progid = PROGIDS[eng]
            old = render_deck(progid, base_path, os.path.join(tmp, 'old'), eng)
            new = render_deck(progid, cur_path, os.path.join(tmp, 'new'), eng)
            regress, recovered, dead = [], [], []
            for page in sorted(set(old) & set(new)):
                if only_pages and page not in only_pages:
                    continue
                a, b = ink(old[page]), ink(new[page])
                if a < 300:
                    continue
                ratio = b / max(a, 1)
                if ratio < 0.98:
                    regress.append((page, round(ratio, 2)))
                elif ratio > 1.02:
                    recovered.append((page, round(ratio, 2)))
                lost, extra = lost_extra(old[page], new[page])
                if lost > 1500 and lost > extra * 2:
                    dead.append((page, lost, extra))
            print('  %-3s 回归页=%d 恢复页=%d 真删除页=%d'
                  % (eng, len(regress), len(recovered), len(dead)))
            if regress:
                print('       回归：%s' % regress[:12])
            if recovered:
                print('       恢复：%s' % recovered[:12])
            if dead:
                print('       真删除：%s' % dead[:12])
            summary.append((f, eng, len(regress), len(recovered), len(dead)))
        for d in ('old', 'new'):
            shutil.rmtree(os.path.join(tmp, d), ignore_errors=True)
    shutil.rmtree(tmp, ignore_errors=True)

    print('\n==== 汇总 ====')
    for f, eng, r, g, d in summary:
        print('%-40s %-3s 回归=%d 恢复=%d 删除=%d' % (f, eng, r, g, d))


if __name__ == '__main__':
    main()
