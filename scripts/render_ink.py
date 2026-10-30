# -*- coding: utf-8 -*-
"""Render one deck with one engine and write per-page ink to JSON.

Kept as a separate process on purpose: PowerPoint and WPS COM instances collide
when they coexist in one process ("RPC 服务器不可用").

Usage:
    python render_ink.py <progid> <pptx> <out.json>

    progid: KWPP.Application        (WPS 演示)
            PowerPoint.Application  (PowerPoint)

Output: {"1": 5065, "2": 10325, ...}  page -> dark pixel count
"""
import json
import os
import sys

try:
    import win32com.client as win32
except ImportError:
    raise SystemExit('需要 win32com（pip install pywin32）')
try:
    from PIL import Image
except ImportError:
    raise SystemExit('需要 Pillow（pip install pillow）')


def main():
    if len(sys.argv) != 4:
        raise SystemExit('用法: python render_ink.py <progid> <pptx> <out.json>')
    progid, path, out_json = sys.argv[1], sys.argv[2], sys.argv[3]

    app = win32.Dispatch(progid)
    try:
        app.Visible = False
    except Exception:
        pass
    pres = app.Presentations.Open(path)
    workdir = os.path.dirname(os.path.abspath(out_json))
    os.makedirs(workdir, exist_ok=True)
    png = os.path.join(workdir, 'page.png')

    result = {}
    try:
        for i in range(1, pres.Slides.Count + 1):
            pres.Slides(i).Export(png, 'PNG', 1280, 720)
            im = Image.open(png).convert('L')
            px = im.load()
            w, h = im.size
            n = 0
            for y in range(0, h, 2):
                for x in range(0, w, 2):
                    if px[x, y] < 200:
                        n += 1
            result[i] = n
            im.close()
    finally:
        pres.Close()
        try:
            app.Quit()
        except Exception:
            pass

    with open(out_json, 'w', encoding='utf-8') as fh:
        json.dump(result, fh)


if __name__ == '__main__':
    main()
