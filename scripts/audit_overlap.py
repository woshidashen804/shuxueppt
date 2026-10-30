# -*- coding: utf-8 -*-
"""Locate text-on-text collisions using PowerPoint's own glyph bounds.

Shape boxes overlap harmlessly all the time (a card behind its label, an answer
sitting inside the blank of its own question). What matters is whether the DRAWN
text rectangles collide, and whether the collision is with another question.

Two traps this script avoids:

  * TextRange2.Bound* can fail on some shapes; fall back to TextRange.
  * A short answer inside its own （　） blank is intended. The check therefore
    requires the hit text to be long enough to be a real sentence, and reports the
    exact characters that would be covered.

Windows only.

Usage:
    python audit_overlap.py --decks DIR [--pages 7,10] [--min-hit 0.12]
"""
import argparse
import os
import sys

try:
    import win32com.client as win32
except ImportError:
    raise SystemExit('需要 win32com（pip install pywin32）')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptlib import deck_files

PT_PER_INCH = 72.0
FREE_CHARS = set('\uff08\u3000\uff09()  ')


def text_bounds(shape):
    for get in (lambda: shape.TextFrame2.TextRange,
                lambda: shape.TextFrame.TextRange):
        try:
            tr = get()
            return (tr.BoundLeft / PT_PER_INCH, tr.BoundTop / PT_PER_INCH,
                    tr.BoundWidth / PT_PER_INCH, tr.BoundHeight / PT_PER_INCH)
        except Exception:
            continue
    return None


def char_spans(shape):
    tr = shape.TextFrame.TextRange
    n = len(tr.Text)
    spans = []
    for i in range(1, n + 1):
        c = tr.Characters(i, 1)
        try:
            left = c.BoundLeft / PT_PER_INCH
            spans.append((c.Text, left, left + c.BoundWidth / PT_PER_INCH))
        except Exception:
            spans.append((c.Text, None, None))
    return spans


def audit_slide(slide, deck, page, min_hit):
    items = []
    for j in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(j)
        try:
            if not (shape.HasTextFrame and shape.TextFrame.HasText):
                continue
            txt = shape.TextFrame.TextRange.Text.strip()
        except Exception:
            continue
        if not txt:
            continue
        g = text_bounds(shape)
        if g is None or g[2] <= 0 or g[3] <= 0:
            continue
        items.append((shape, txt, g))

    findings = []
    for a in range(len(items)):
        for b in range(len(items)):
            if a == b:
                continue
            q_shape, q_txt, q_geo = items[a]
            a_shape, a_txt, a_geo = items[b]
            if len(q_txt) < len(a_txt):
                continue          # q must be the longer (question) text
            if not (a_geo[1] < q_geo[1] + q_geo[3] - 0.05
                    and a_geo[1] + a_geo[3] > q_geo[1] + 0.05):
                continue
            if not (q_geo[0] - 0.05 <= a_geo[0] <= q_geo[0] + q_geo[2]):
                continue
            left, right = a_geo[0], a_geo[0] + a_geo[2]
            covered = []
            for ch, l, r in char_spans(q_shape):
                if l is None or r is None:
                    continue
                overlap = min(right, r) - max(left, l)
                if overlap <= 0 or ch in FREE_CHARS:
                    continue
                covered.append((ch, round(overlap, 3)))
            real = sum(v for _, v in covered)
            if real > min_hit:
                findings.append((page, a_shape.Name, a_txt[:28],
                                 q_shape.Name, covered[:6]))
    return findings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--pages', default=None)
    ap.add_argument('--min-hit', type=float, default=0.12)
    args = ap.parse_args()

    only = None
    if args.pages:
        only = {int(x) for x in args.pages.split(',') if x.strip()}

    app = win32.Dispatch('PowerPoint.Application')
    try:
        app.Visible = False
    except Exception:
        pass

    total = 0
    for f in deck_files(args.decks):
        pres = app.Presentations.Open(os.path.join(args.decks, f))
        try:
            for page in range(1, pres.Slides.Count + 1):
                if only and page not in only:
                    continue
                found = audit_slide(pres.Slides(page), f, page, args.min_hit)
                for page_no, name, txt, qname, covered in found:
                    total += 1
                    print('%s P%02d  %-14s %-30r 压在 %s 上'
                          % (f, page_no, name, txt, qname))
                    print('      命中字符：%s' % covered)
        finally:
            pres.Close()
    try:
        app.Quit()
    except Exception:
        pass
    print('\n文字压字实例：%d' % total)


if __name__ == '__main__':
    main()
