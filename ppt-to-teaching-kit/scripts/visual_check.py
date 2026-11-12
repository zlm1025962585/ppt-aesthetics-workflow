#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""视觉底线确定性校验：把"审美"里可判定的部分变成断言。

用法:
  python visual_check.py --tokens visual_tokens.json --pptx 课件.pptx --html a.html b.html
  python visual_check.py --tokens visual_tokens.json --pptx 课件.pptx --min-slides 24 --max-slides 30

检查 PPTX：16:9、页数区间、每页可编辑文本框下限、显式字号下限、色板集中度、
            形状是否越出画布；图片页是否有足够文字（防止题干被栅格化）。
检查 HTML：正文字体首项、12pt 正文、2em 缩进、每个 table 的 thead/tbody、
            @page 存在且注释内无该字样、h1 上限、封面标题标签、CSS 裸样式值。

设计原则：取不到的属性一律 SKIP 而不是判违规，避免假阳性；能取到的才断言。
输出 PASS / FAIL / SKIP 明细，任一 FAIL 退出码 1。
"""
import argparse
import json
import os
import re
import sys


def norm_color(rgb, palette):
    return ('#' + str(rgb).upper().lstrip('#')) in palette


def check_pptx(path, tokens, min_slides, max_slides):
    from pptx import Presentation
    from pptx.oxml.ns import qn
    prs = Presentation(path)
    w, h = prs.slide_width, prs.slide_height
    out = []

    def add(name, ok, detail, skipped=False):
        out.append(('SKIP' if skipped else ('PASS' if ok else 'FAIL'), name, detail))

    add('画布 16:9', abs(w / h - 16 / 9) < 0.02, '%.4f (%.2f x %.2f in)' % (w / h, w / 914400, h / 914400))
    n = len(prs.slides)
    add('页数区间 %d–%d' % (min_slides, max_slides), min_slides <= n <= max_slides, '实际 %d 页' % n)

    palette = set(c.upper() for c in tokens['palette'].values())
    palette |= set(c.upper() for c in tokens.get('neutral_allowed', []))
    min_shapes = tokens['layout_rules'].get('min_text_shapes_per_slide', 2)
    min_pt = tokens['typography'].get('min_font_pt', 10)
    weak, small, overflow, color_all, color_hit, pic_only = [], [], [], 0, 0, []
    for idx, slide in enumerate(prs.slides, 1):
        text_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        pics = [s for s in slide.shapes if s.shape_type == 13]
        if len(text_shapes) < min_shapes:
            weak.append(idx)
        if pics and len(text_shapes) < min_shapes:
            pic_only.append(idx)
        for s in slide.shapes:
            if s.has_text_frame:
                for p in s.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size is not None and r.font.size.pt < min_pt:
                            small.append((idx, round(r.font.size.pt, 1)))
                        try:
                            if r.font.color and r.font.color.rgb is not None:
                                color_all += 1
                                color_hit += norm_color(r.font.color.rgb, palette)
                        except (AttributeError, ValueError, TypeError):
                            pass
            try:
                if s.fill.type is not None and s.fill.fore_color and s.fill.fore_color.rgb is not None:
                    color_all += 1
                    color_hit += norm_color(s.fill.fore_color.rgb, palette)
            except (AttributeError, ValueError, TypeError):
                pass
            try:
                if (s.left is not None and s.top is not None and s.width is not None and s.height is not None
                        and (s.left < -1 or s.top < -1 or s.left + s.width > w + 1 or s.top + s.height > h + 1)):
                    overflow.append(idx)
            except (TypeError, ValueError):
                pass
    add('每页可编辑文本框 ≥%d' % min_shapes, not weak, '薄弱页: %s' % (weak or '无'))
    add('图片页仍保有文字', not pic_only, '疑似图文替代页: %s' % (pic_only or '无'))
    add('显式字号 ≥%s pt' % min_pt, not small, '过小: %s' % (small[:5] or '无'))
    add('形状不越出画布', not overflow, '越界页: %s' % (sorted(set(overflow)) or '无'))

    # 视觉资产是否真的落到页面上（防止退化成纯色框）
    rules = tokens.get('layout_rules', {})
    min_pics = rules.get('min_pictures_per_deck', 3)
    cover_ratio_min = rules.get('cover_picture_min_area_ratio', 0.85)
    pic_pages, cover_ratio, grads = [], 0.0, 0
    for idx, slide in enumerate(prs.slides, 1):
        for s in slide.shapes:
            if s.shape_type == 13:
                pic_pages.append(idx)
                if idx == 1 and s.width and s.height:
                    cover_ratio = max(cover_ratio, (s.width * s.height) / float(w * h))
            try:
                spPr = s._element.spPr
            except AttributeError:
                spPr = None
            if spPr is not None and spPr.find(qn('a:gradFill')) is not None:
                grads += 1
    add('封面满版插画 ≥%d%%' % int(cover_ratio_min * 100), cover_ratio >= cover_ratio_min,
        '封面图片覆盖 %.0f%%' % (cover_ratio * 100))
    add('含插画页数 ≥%d' % min_pics, len(set(pic_pages)) >= min_pics,
        '插画页: %s' % (sorted(set(pic_pages)) or '无'))
    add('渐变蒙层 ≥2 处', grads >= 2, 'gradFill 数量 %d' % grads)
    ratio = (color_hit / color_all) if color_all else None
    add('色板集中度 ≥70%', ratio is None or ratio >= 0.70,
        '命中 %d/%d = %s' % (color_hit, color_all, '%.0f%%' % (ratio * 100) if ratio else 'n/a'),
        skipped=ratio is None)
    return out


def check_html(path, tokens):
    raw = open(path, encoding='utf-8').read()
    out = []

    def add(name, ok, detail):
        out.append(('PASS' if ok else 'FAIL', name, os.path.basename(path) + ': ' + detail))

    rule = tokens['docx_html_rules']
    first = rule['body_font_first']
    add('正文首项字体 = %s' % first,
        bool(re.search(r'--ff-body\s*:\s*["\']?%s' % first, raw, re.I)),
        '--ff-body 首项')
    add('正文 %s pt' % rule['body_size_pt'],
        bool(re.search(r'--fs-body\s*:\s*%spt' % rule['body_size_pt'], raw)),
        '--fs-body')
    add('正文首行缩进 %s' % rule['body_indent'],
        ('text-indent:' + rule['body_indent']).replace(' ', '') in raw.replace(' ', ''),
        'text-indent')
    tables = re.findall(r'<table\b[^>]*>(.*?)</table>', raw, re.S | re.I)
    bad_tables = [i + 1 for i, t in enumerate(tables)
                  if not all(tag in t for tag in rule['table_requires'])]
    add('每个 table 含 thead/tbody', not bad_tables,
        '%d 个表，缺结构: %s' % (len(tables), bad_tables or '无'))
    no_comments = re.sub(r'/\*.*?\*/', '', raw, flags=re.S)
    add('@page 页面规则', '@page' in no_comments,
        '含页面规则 %d 处' % no_comments.count('@page'))
    comments = re.findall(r'/\*.*?\*/', raw, flags=re.S)
    add('CSS 注释不含 @page', not any('@page' in c for c in comments), '注释 %d 段' % len(comments))
    h1 = len(re.findall(r'<h1\b', raw, re.I))
    add('h1 ≤ %d' % rule['max_h1'], h1 <= rule['max_h1'], 'h1 = %d' % h1)
    add('封面标题用 <p>', bool(re.search(r'<p[^>]*class="cover-title"', raw)), 'cover-title')
    css = re.search(r'<style>(.*?)</style>', raw, re.S | re.I)
    naked = []
    if css:
        block = re.sub(r':root\s*\{.*?\}', '', css.group(1), flags=re.S)
        for line in block.split(';'):
            if 'var(' in line:
                continue
            if re.match(r'\s*text-indent\s*:', line):
                continue
            if re.search(r':\s*(#[0-9a-fA-F]{3,8}|\d+(?:\.\d+)?(?:px|pt|em|cm))\b', line):
                naked.append(line.strip()[:70])
    add('CSS 无裸样式值', not naked, '裸值 %s' % (naked[:3] or '无'))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--tokens', required=True)
    ap.add_argument('--pptx')
    ap.add_argument('--html', nargs='*', default=[])
    ap.add_argument('--min-slides', type=int, default=24)
    ap.add_argument('--max-slides', type=int, default=30)
    a = ap.parse_args()
    tokens = json.load(open(a.tokens, encoding='utf-8'))
    rows = []
    for path in ([a.pptx] if a.pptx else []):
        if not os.path.exists(path):
            print('文件不存在:', path)
            return 2
        rows.append(('PPTX', check_pptx(path, tokens, a.min_slides, a.max_slides)))
    for path in a.html:
        if not os.path.exists(path):
            print('文件不存在:', path)
            return 2
        rows.append(('HTML', check_html(path, tokens)))
    if not rows:
        print('未指定 --pptx 或 --html')
        return 2
    fails = 0
    for kind, checks in rows:
        print('---', kind, '---')
        for status, name, detail in checks:
            print('%-4s %-28s %s' % (status, name, detail))
            fails += (status == 'FAIL')
    print('结论:', 'PASS（视觉底线达标）' if not fails else 'FAIL（%d 项未达标）' % fails)
    return 0 if not fails else 1


if __name__ == '__main__':
    sys.exit(main())
