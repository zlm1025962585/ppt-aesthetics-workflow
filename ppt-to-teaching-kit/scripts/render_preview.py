# -*- coding: utf-8 -*-
"""用本机 PowerPoint 把指定页导出为 PNG，便于肉眼验收（Windows 可选工具）。"""
import os
import sys

import win32com.client as win32


def render(pptx, out_dir, pages=None, width=1600, height=900):
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    app = win32.Dispatch('PowerPoint.Application')
    pres = None
    out = []
    try:
        pres = app.Presentations.Open(os.path.abspath(pptx), WithWindow=False)
        total = pres.Slides.Count
        targets = pages or list(range(1, total + 1))
        for i in targets:
            if i < 1 or i > total:
                continue
            p = os.path.join(out_dir, 'p%02d.png' % i)
            pres.Slides(i).Export(p, 'PNG', width, height)
            out.append(p)
    finally:
        if pres is not None:
            pres.Close()
        app.Quit()
    return out


if __name__ == '__main__':
    src = sys.argv[1]
    dst = sys.argv[2]
    pages = [int(x) for x in sys.argv[3:]] or None
    for f in render(src, dst, pages):
        print(f)
