#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从结构化内容锁定表一次生成「课件 + 教案 + 任务单」三件套。

平台无关：只用 python-pptx / python-docx，不依赖任何 WorkBuddy 专有工具。
因为三份文件都从同一份 JSON 取值，题干、口诀、任务名、答案、分值天然一致。

视觉不是"让模型自由发挥"：封面满版插画＋渐变蒙层、章节页巨型编号、
方法页四卡、答案页分步卡＋评分胶囊、结尾横幅图，全部由本脚本按固定几何生成，
插画来自 assets/illustrations/。模型只填内容，不设计版式。

用法:
  python build_kit.py --content content_lock.json --out-dir ./out [--tokens visual_tokens.json] [--prefix 课题名]
  python build_kit.py --demo --out-dir ./out        # 用内置《巨人的花园》示例
  python build_kit.py --demo --out-dir ./out --no-images   # 无插画时的降级验证

生成: <out-dir>/<prefix>_学生课件.pptx、<prefix>_教案逐字稿.docx、<prefix>_学生任务单.docx
任务单生成时显式跳过答案字段（answer_steps / answer_full），不印示例答案。
"""
import argparse
import json
import math
import os
import re

from docx import Document
from docx.shared import Pt, Inches, RGBColor as DocxRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Emu, Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

DEFAULT_TOKENS = {
    "palette": {"primary": "#43A047", "secondary": "#A5D6A7", "accent": "#FB8C00",
                "background": "#FDFCF3", "text": "#263238", "title": "#1B5E20",
                "quote": "#5D4037", "flower": "#F8BBD0", "muted": "#78909C",
                "cream": "#FFF8E1", "light": "#E8F5E9", "border": "#C8E6C9"},
    "typography": {"heading_font": "Microsoft YaHei", "quote_font": "KaiTi"},
    "visual": {"illustrations": {"cover": "assets/illustrations/cover_garden.jpg",
                                 "cold": "assets/illustrations/winter_garden.jpg",
                                 "spring": "assets/illustrations/spring_garden.jpg"}}
}


def rgb(hexstr):
    h = hexstr.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


SLIDE_W_EMU = 12192000          # 13.333 in
SLIDE_H_EMU = 6858000           # 7.5 in（= 12192000 * 720 / 1280，整除无误差）


def px(v):
    """设计稿 1280x720 像素 -> EMU，按画布精确换算，避免 13.333in 的舍入越界。"""
    return Emu(int(round(v * SLIDE_W_EMU / 1280.0)))


def image_size(path):
    """不依赖 Pillow 读取 PNG/JPEG 尺寸，失败时退回 16:9。"""
    try:
        with open(path, 'rb') as f:
            head = f.read(26)
            if head[:8] == b'\x89PNG\r\n\x1a\n':
                return int.from_bytes(head[16:20], 'big'), int.from_bytes(head[20:24], 'big')
            f.seek(0)
            if f.read(2) == b'\xff\xd8':
                while True:
                    b = f.read(1)
                    while b and b != b'\xff':
                        b = f.read(1)
                    while b == b'\xff':
                        b = f.read(1)
                    if not b:
                        break
                    marker = b[0]
                    if marker in (0xD8, 0xD9):
                        continue
                    ln = int.from_bytes(f.read(2), 'big')
                    if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                        f.read(1)
                        h = int.from_bytes(f.read(2), 'big')
                        w = int.from_bytes(f.read(2), 'big')
                        return w, h
                    f.seek(ln - 2, 1)
    except OSError:
        pass
    return 16, 9


def task_short(k):
    """「挑战一 · 想象画面题」-> 「想象画面」"""
    tail = k['task'].split('·')[-1].strip()
    return tail[:-1] if tail.endswith('题') else tail


def unit_label(k):
    m = re.search(r'（([^）]*单元)）', k.get('element', ''))
    return m.group(1) if m else ''


def parse_rubric(item):
    """「① 圈出关键词 1 分」-> ("1分", "圈出关键词")"""
    m = re.search(r'(\d+(?:\.\d+)?)\s*分', item)
    label = (m.group(1) + '分') if m else '1分'
    text = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\s、，,]*', '', item)
    text = re.sub(r'\s*\d+(?:\.\d+)?\s*分', '', text)
    return label, text.strip(' ；;，,')


def set_line_none(shape):
    shape.line.fill.background()


def set_solid(shape, color, alpha=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    if alpha is not None:
        sf = shape._element.spPr.find(qn('a:solidFill'))
        srgb = sf.find(qn('a:srgbClr'))
        srgb.append(parse_xml('<a:alpha xmlns:a="%s" val="%d"/>' % (A_NS, int(alpha * 100000))))


def set_gradient(shape, stops, angle=90):
    """stops: [(pos 0-1, '#RRGGBB', alpha 0-1), ...]；angle 顺时针、0=从左到右。"""
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'):
        for e in spPr.findall(qn(tag)):
            spPr.remove(e)
    buf = ['<a:gradFill xmlns:a="%s" rotWithShape="1"><a:gsLst>' % A_NS]
    for pos, color, alpha in stops:
        buf.append('<a:gs pos="%d"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:gs>'
                   % (int(pos * 100000), color.lstrip('#').upper(), int(alpha * 100000)))
    buf.append('</a:gsLst><a:lin ang="%d" scaled="0"/></a:gradFill>' % int(angle * 60000))
    frag = parse_xml(''.join(buf))
    ln = spPr.find(qn('a:ln'))
    if ln is not None:
        ln.addprevious(frag)
    else:
        spPr.append(frag)


def apply_font(run, name, size=None, bold=None, color=None, spc=None, shadow=False):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn('a:latin'))
    if latin is None:
        latin = parse_xml('<a:latin xmlns:a="%s" typeface="%s"/>' % (A_NS, name))
        rPr.append(latin)
    latin.set('typeface', name)
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        latin.addnext(parse_xml('<a:ea xmlns:a="%s" typeface="%s"/>' % (A_NS, name)))
    else:
        ea.set('typeface', name)
    if size is not None:
        run.font.size = PPt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = rgb(color)
    if spc is not None:
        rPr.set('spc', str(int(spc * 100)))
    if shadow:
        rPr.append(parse_xml(
            '<a:effectLst xmlns:a="%s"><a:outerShdw blurRad="76200" dist="19050" dir="5400000" algn="ctr">'
            '<a:srgbClr val="000000"><a:alpha val="32000"/></a:srgbClr></a:outerShdw></a:effectLst>' % A_NS))


class Deck:
    def __init__(self, tokens, data):
        self.p = tokens['palette']
        self.t = tokens.get('typography', {})
        self.v = tokens.get('visual', {})
        self.data = data
        self.warnings = []
        prs = Presentation()
        prs.slide_width = Emu(SLIDE_W_EMU)
        prs.slide_height = Emu(SLIDE_H_EMU)
        self.prs = prs
        self.n = 0
        self.total = 30
        self.head_font = self.t.get('heading_font', 'Microsoft YaHei')
        self.quote_font = self.t.get('quote_font', 'KaiTi')

    # ---------- 基础构件 ----------
    def blank(self, bg=None, footer=True):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.n += 1
        if bg:
            rect = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, px(1280), px(720))
            rect.fill.solid()
            rect.fill.fore_color.rgb = rgb(bg)
            set_line_none(rect)
            rect.shadow.inherit = False
        else:
            f = s.background.fill
            f.solid()
            f.fore_color.rgb = rgb(self.p['background'])
        if footer:
            self.text(s, 60, 690, 640, 24, '%s · %s' % (self.data['scenario'], self.data['role']),
                      size=14, color=self.p['muted'])
            self.text(s, 900, 690, 320, 24, '%d / %d' % (self.n, self.total),
                      size=14, color=self.p['muted'], align=PP_ALIGN.RIGHT)
        return s

    def fit(self, value, w, h, base, min_size=10, line=1.35):
        """按字数估算最大可容纳字号（字号单位 pt，框尺寸单位 px；中文行高取 1.2 倍）。"""
        lines = str(value).split('\n')
        w_pt, h_pt = w * 0.75, h * 0.75
        pt = float(base)
        while pt > min_size:
            cpl = max(1, int(w_pt / pt))
            need = sum(max(1, math.ceil(len(t) / cpl)) for t in lines)
            if need * pt * 1.2 * line <= h_pt:
                return round(pt, 1)
            pt -= 0.5
        return min_size

    def rich_fit(self, pairs, w, h, base, line=1.9, min_size=10):
        """成对文本（加粗标签 + 正文）的最大可容纳字号。"""
        w_pt, h_pt = w * 0.75, h * 0.75
        pt = float(base)
        while pt > min_size:
            cpl = max(1, int(w_pt / pt))
            need = sum(max(1, math.ceil((len(a) + len(b)) / cpl)) for a, b in pairs)
            if need * pt * 1.2 * line <= h_pt:
                return round(pt, 1)
            pt -= 0.5
        return min_size

    def text(self, s, x, y, w, h, value, size=16, color=None, bold=False, font=None,
             align=PP_ALIGN.LEFT, line=1.35, valign=MSO_ANCHOR.TOP, spc=None, shadow=False):
        box = s.shapes.add_textbox(px(x), px(y), px(w), px(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = valign
        for i, ln in enumerate(str(value).split('\n')):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            para.line_spacing = line
            run = para.add_run()
            run.text = ln
            apply_font(run, font or self.head_font, size, bold, color or self.p['text'], spc, shadow)
        return box


    def rich_para(self, s, x, y, w, h, parts, size=15, line=1.55,
                  valign=MSO_ANCHOR.TOP, font=None):
        """单段多色：parts = [(文字, 颜色, 是否加粗), ...]"""
        box = s.shapes.add_textbox(px(x), px(y), px(w), px(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = valign
        para = tf.paragraphs[0]
        para.line_spacing = line
        for text, color, bold in parts:
            run = para.add_run()
            run.text = text
            apply_font(run, font or self.head_font, size, bold, color)
        return box

    def rect(self, s, x, y, w, h, fill=None, line=None, radius=None, alpha=None,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE):
        shp = s.shapes.add_shape(shape, px(x), px(y), px(w), px(h))
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            try:
                shp.adjustments[0] = radius
            except (IndexError, ValueError):
                pass
        if fill:
            set_solid(shp, fill, alpha)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = rgb(line)
            shp.line.width = PPt(1)
        else:
            set_line_none(shp)
        shp.shadow.inherit = False
        return shp

    def circle(self, s, x, y, dia, label, fill, size=18, color='#FFFFFF'):
        shp = s.shapes.add_shape(MSO_SHAPE.OVAL, px(x), px(y), px(dia), px(dia))
        set_solid(shp, fill)
        set_line_none(shp)
        shp.shadow.inherit = False
        tf = shp.text_frame
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = str(label)
        apply_font(run, self.head_font, size, True, color)
        return shp

    def card(self, s, x, y, w, h, heading=None, lines=None, fill='#FFFFFF', border=None,
             head_color=None, body_color=None, size=15, head_size=17, radius=0.08,
             pad=18, align=PP_ALIGN.LEFT, line=1.55, font=None):
        self.rect(s, x, y, w, h, fill=fill, line=border or self.p['border'], radius=radius)
        inner_w = w - pad * 2
        top = y + pad
        if heading:
            self.text(s, x + pad, top, inner_w, 28, heading, size=head_size, bold=True,
                      color=head_color or self.p['title'])
            top += head_size * 1.9
        if lines:
            body = '\n'.join(lines) if isinstance(lines, (list, tuple)) else str(lines)
            avail = (y + h - pad) - top
            fs = self.fit(body, inner_w, avail, size, min_size=10, line=line)
            self.text(s, x + pad, top, inner_w, avail, body, size=fs, color=body_color or self.p['text'],
                      align=align, line=line, font=font)

    def card_rich(self, s, x, y, w, h, heading, pairs, fill='#FFFFFF', border=None,
                  size=14, line=1.9, pad=18, label_color=None):
        """pairs: [(加粗标签, 正文), ...] 同段两色；字号按可用高度自适应。"""
        self.rect(s, x, y, w, h, fill=fill, line=border or self.p['border'], radius=0.08)
        top = y + pad
        self.text(s, x + pad, top, w - pad * 2, 28, heading, size=17, bold=True, color=self.p['title'])
        top += 32
        body_h = y + h - pad - top
        fs = self.rich_fit(pairs, w - pad * 2, body_h, size, line, min_size=9.5)
        box = s.shapes.add_textbox(px(x + pad), px(top), px(w - pad * 2), px(body_h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        for i, (label, body) in enumerate(pairs):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.line_spacing = line
            r1 = para.add_run()
            r1.text = label
            apply_font(r1, self.head_font, fs, True, label_color or self.p['accent'])
            r2 = para.add_run()
            r2.text = body
            apply_font(r2, self.head_font, fs, False, self.p['text'])

    def header(self, s, badge, title, tag=None, tag_fill=None):
        self.circle(s, 60, 46, 48, badge, self.p['accent'], 18)
        self.text(s, 124, 50, 900, 40, title, size=32, bold=True, color=self.p['title'])
        self.rect(s, 60, 108, 1160, 1, fill=self.p['primary'], shape=MSO_SHAPE.RECTANGLE)
        if tag:
            tw = max(120, 26 + len(tag) * 15)
            self.rect(s, 1220 - tw, 48, tw, 40, fill=tag_fill or self.p['cream'], radius=0.5)
            self.text(s, 1220 - tw, 48, tw, 40, tag, size=15, bold=True, color='#E65100',
                      align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    def label(self, s, x, y, w, text, color=None):
        self.text(s, x, y, w, 22, text, size=16, bold=True, color=color or self.p['accent'], spc=2)

    def band(self, s, x, y, w, h, label, text, fill=None, border=None, size=26, font=None):
        self.rect(s, x, y, w, h, fill=fill or self.p['cream'], line=border or '#FFD54F', radius=0.06)
        self.text(s, x + 24, y, 158, h, label, size=17, bold=True, color='#E65100',
                  valign=MSO_ANCHOR.MIDDLE)
        self.text(s, x + 186, y, w - 210, h, text,
                  size=self.fit(text, w - 210, h, size, 15, 1.35), bold=True, color=self.p['quote'],
                  valign=MSO_ANCHOR.MIDDLE, font=font or self.quote_font, line=1.35)

    # ---------- 图像 ----------
    def asset(self, name):
        path = (self.v.get('illustrations') or {}).get(name)
        if not path:
            return None
        path = path if os.path.isabs(path) else os.path.join(SKILL_ROOT, path)
        if not os.path.exists(path):
            self.warnings.append('插画缺失: %s' % path)
            return None
        return path

    def picture(self, s, path, x, y, w, h, bias=0.5):
        """按 cover 规则裁切填满框，不变形。"""
        pic = s.shapes.add_picture(path, px(x), px(y), px(w), px(h))
        iw, ih = image_size(path)
        box_ar, img_ar = w / h, iw / ih
        if img_ar > box_ar:
            frac = 1 - box_ar / img_ar
            pic.crop_left = frac * (1 - bias)
            pic.crop_right = frac * bias
        elif img_ar < box_ar:
            frac = 1 - img_ar / box_ar
            pic.crop_top = frac * bias
            pic.crop_bottom = frac * (1 - bias)
        return pic

    def scrim(self, s, x, y, w, h, stops, angle=90):
        shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(h))
        set_gradient(shp, stops, angle)
        set_line_none(shp)
        shp.shadow.inherit = False
        return shp


# ---------------- 页面 ----------------

def pg_cover(d, data, keys):
    s = d.blank(footer=False)
    cover = d.asset('cover')
    if cover:
        d.picture(s, cover, 0, 0, 1280, 720)
        d.scrim(s, 0, 0, 1280, 720,
                [(0.0, '#FDFCF3', 0.06), (0.34, '#1B5E20', 0.16), (1.0, '#1B5E20', 0.82)], 90)
    else:
        d.scrim(s, 0, 0, 1280, 720,
                [(0.0, '#E8F5E9', 1.0), (0.55, '#FDFCF3', 1.0), (1.0, '#1B5E20', 0.85)], 90)
    d.text(s, 80, 372, 1120, 30, '%s · 集 合 令' % data['role'], size=18, bold=True,
           color='#FFF8E1', align=PP_ALIGN.CENTER, spc=5, shadow=True)
    d.text(s, 60, 408, 1160, 80, data['scenario'], size=d.fit(data['scenario'], 1160, 80, 60, 42),
           bold=True, color='#FFFFFF', align=PP_ALIGN.CENTER, line=1.1, shadow=True)
    d.text(s, 60, 498, 1160, 40, '童话阅读答题技巧 · 四把钥匙等你来拿', size=26, color='#FFFFFF',
           align=PP_ALIGN.CENTER, shadow=True)
    d.rect(s, 330, 552, 620, 1, fill='#FFFFFF', shape=MSO_SHAPE.RECTANGLE, alpha=0.65)
    d.text(s, 60, 572, 1160, 34, '四项挑战：' + ' · '.join(task_short(k) for k in keys),
           size=19, color='#FFF8E1', align=PP_ALIGN.CENTER)
    d.text(s, 60, 618, 1160, 28,
           '阅读材料：%s　　适用年级：%s' % (data['textbook'], data.get('grade', '四年级')),
           size=15, color='#E8F5E9', align=PP_ALIGN.CENTER)


def pg_intro(d, data, keys):
    s = d.blank()
    cold = d.asset('cold')
    if cold:
        d.picture(s, cold, 760, 0, 520, 720)
        d.scrim(s, 760, 0, 520, 720, [(0.0, '#FDFCF3', 1.0), (0.72, '#FDFCF3', 0.0)], 0)
    d.label(s, 60, 78, 600, '情 境 导 入')
    d.text(s, 60, 106, 660, 70, '一封来自巨人的求助信', size=42, bold=True, color=d.p['title'])
    d.text(s, 60, 186, 660, 30, '花园又被冬天困住了，只有你能帮忙。', size=18, color=d.p['muted'])

    letter = data.get('intro_letter') or [
        '我的花园又被冬天困住了。',
        '只有集齐四把「阅读钥匙」，春天才会回来。',
        '今天，你们化身——%s！' % data['role'],
        '依次完成四项挑战，把春天带回花园。',
    ]
    d.rect(s, 60, 222, 660, 372, fill='#FFFFFF', line=d.p['border'], radius=0.05)
    body = '\n'.join(letter)
    d.text(s, 92, 248, 596, 272, body,
           size=d.fit(body, 596, 272, 20, 13, 2.0), color=d.p['text'], line=2.0, font=d.quote_font)
    d.text(s, 92, 528, 596, 26, '—— 被困在冬天里的巨人', size=15, color=d.p['muted'],
           align=PP_ALIGN.RIGHT, font=d.quote_font)
    d.text(s, 60, 622, 660, 30, '四把钥匙：' + ' · '.join(task_short(k) for k in keys),
           size=16, color=d.p['title'], bold=True)


def pg_overview(d, data, keys):
    s = d.blank()
    d.header(s, '★', '本课挑战：四把钥匙等你来拿', tag='每题 %d 分' % data.get('score_per_task', 5))
    colors = [d.p['primary'], d.p['accent'], d.p['primary'], d.p['accent']]
    for i, k in enumerate(keys):
        x = 60 + i * 288
        d.rect(s, x, 150, 272, 410, fill='#FFFFFF', line=d.p['border'], radius=0.06)
        d.rect(s, x, 150, 272, 8, fill=colors[i], radius=0.5)
        d.circle(s, x + 22, 180, 46, str(i + 1), colors[i], 20)
        d.text(s, x + 22, 240, 228, 34, k['name'], size=24, bold=True, color=d.p['title'])
        d.text(s, x + 22, 284, 228, 24, k['task'], size=14, bold=True, color=colors[i])
        d.rect(s, x + 22, 318, 40, 2, fill=colors[i], shape=MSO_SHAPE.RECTANGLE)
        d.text(s, x + 22, 334, 228, 110, k['element'], size=14, color=d.p['text'], line=1.55)
        d.text(s, x + 22, 452, 228, 100, '口诀：' + k['mnemonic'], size=13, color=d.p['quote'], line=1.5)
    d.text(s, 60, 586, 1160, 26,
           '集齐顺序：%s　→　春天回来' % '　→　'.join(k['name'] for k in keys),
           size=15, color=d.p['muted'], align=PP_ALIGN.CENTER)


def pg_goals(d, data, keys):
    s = d.blank()
    d.header(s, '◎', '我要学会：四项阅读本领')
    for i, k in enumerate(keys):
        y = 140 + i * 108
        d.rect(s, 60, y, 1160, 92, fill='#FFFFFF', line=d.p['border'], radius=0.08)
        d.rect(s, 60, y, 6, 92, fill=d.p['accent'] if i % 2 else d.p['primary'], radius=0.5)
        d.circle(s, 86, y + 24, 44, '0%d' % (i + 1), d.p['light'], 16, d.p['title'])
        d.text(s, 148, y + 14, 600, 30, k['task'],
               size=d.fit(k['task'], 600, 30, 20, 14, 1.2), bold=True, color=d.p['title'])
        elem = re.sub(r'（[^）]*单元）', '', k['element'])
        d.text(s, 148, y + 50, 800, 26, elem, size=14, color=d.p['text'])
        pill = unit_label(k)
        if pill:
            d.rect(s, 1060, y + 28, 136, 36, fill=d.p['cream'], radius=0.5)
            d.text(s, 1060, y + 28, 136, 36, pill, size=13, bold=True, color='#E65100',
                   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    d.rect(s, 60, 590, 1160, 54, fill=d.p['light'], line=None, radius=0.3)
    d.text(s, 84, 590, 1112, 54, '答题三金律：' + ' · '.join(data.get('golden_rules', [])),
           size=15, bold=True, color=d.p['title'], valign=MSO_ANCHOR.MIDDLE)


def pg_material(d, data, keys):
    s = d.blank()
    d.header(s, '▤', '阅读材料：%s' % data['textbook'])
    for i, k in enumerate(keys):
        y = 132 + i * 120
        d.rect(s, 60, y, 1160, 112, fill='#FFFDE7', line='#D7CCC8', radius=0.08)
        d.rect(s, 60, y, 5, 112, fill=d.p['accent'], radius=0.5)
        d.text(s, 84, y + 9, 500, 24, '%s · 原文节选' % k['task'], size=14, bold=True, color='#E65100')
        src = k['source']
        if len(src) > 86:
            src = src[:86] + '……'
        d.text(s, 84, y + 42, 1108, 58, src,
               size=d.fit(src, 1108, 58, 14, 11, 1.5), color=d.p['quote'], line=1.5, font=d.quote_font)


def pg_divider(d, data, k, idx):
    s = d.blank()
    d.scrim(s, 0, 0, 1280, 720,
            [(0.0, '#E8F5E9', 1.0), (0.55, '#FDFCF3', 1.0), (1.0, '#FFF8E1', 1.0)], 45)
    d.text(s, 60, 150, 250, 130, '%02d' % idx, size=110, bold=True, color=d.p['secondary'], line=1.0)
    d.rect(s, 300, 168, 1, 96, fill=d.p['secondary'], shape=MSO_SHAPE.RECTANGLE)
    d.text(s, 330, 158, 500, 28, '第 %s 关 · 钥匙收集' % '一二三四'[idx - 1], size=20,
           bold=True, color=d.p['accent'], spc=3)
    d.text(s, 330, 190, 700, 64, k['name'], size=42, bold=True, color=d.p['title'], line=1.05)
    d.text(s, 330, 262, 760, 32, k['task'], size=22, bold=True, color='#33691E')

    # 右侧钥匙进度栏：已收集 / 当前 / 待收集
    for i, kk in enumerate(d.data['keys']):
        y = 174 + i * 80
        done, now = i < idx - 1, i == idx - 1
        fill = d.p['accent'] if now else (d.p['light'] if done else '#FFFFFF')
        line = None if now else (d.p['secondary'] if done else '#CFD8DC')
        d.rect(s, 860, y, 360, 68, fill=fill, line=line, radius=0.16)
        d.text(s, 882, y + 9, 200, 24, kk['name'], size=16, bold=True,
               color='#FFFFFF' if now else (d.p['title'] if done else d.p['muted']))
        d.text(s, 882, y + 37, 200, 20, task_short(kk), size=11.5,
               color='#FFF8E1' if now else d.p['muted'])
        d.text(s, 1036, y, 108, 68, '待收集' if not done and not now else ('进行中' if now else '已收集'),
               size=13, bold=True, color='#FFFFFF' if now else (d.p['primary'] if done else d.p['muted']),
               align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE)
        d.text(s, 1156, y, 38, 68, '✓' if done else ('●' if now else '○'), size=15,
               color='#FFFFFF' if now else (d.p['primary'] if done else d.p['muted']),
               align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    d.rect(s, 60, 330, 760, 190, fill='#FFFFFF', line=d.p['border'], radius=0.06)
    d.label(s, 84, 348, 400, '情 境 导 入')
    intro = (data.get('task_intro') or {}).get(str(idx)) or '带着钥匙出发，完成这一关的挑战。'
    d.text(s, 84, 376, 712, 124, intro, size=d.fit(intro, 712, 124, 19, 14, 1.7),
           color=d.p['text'], line=1.7)

    d.rect(s, 60, 548, 118, 40, fill=d.p['accent'], radius=0.5)
    d.text(s, 60, 548, 118, 40, '我要学会', size=16, bold=True, color='#FFFFFF',
           align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    d.text(s, 194, 548, 1026, 40, k['element'], size=17, color='#546E7A', valign=MSO_ANCHOR.MIDDLE)


def pg_source_question(d, data, k, idx, per):
    s = d.blank()
    d.header(s, '%02d' % idx, '挑战%s：读原文，认领题目' % '一二三四'[idx - 1],
             tag='%s · %d 分' % (task_short(k) + '题', per))
    d.label(s, 60, 126, 760, '呈 现 原 文 · %s' % data['textbook'])
    d.rect(s, 60, 154, 1160, 222, fill='#FFFDE7', line='#D7CCC8', radius=0.05)
    d.text(s, 84, 174, 1112, 182, k['source'],
           size=d.fit(k['source'], 1112, 182, 17, 12, 1.75), color=d.p['quote'],
           line=1.75, font=d.quote_font)

    d.rect(s, 60, 396, 760, 252, fill='#FFFFFF', line=d.p['accent'], radius=0.05)
    d.text(s, 60, 396, 760, 252, '', size=12)
    d.label(s, 84, 414, 500, '出 示 题 目')
    d.text(s, 84, 442, 712, 190, k['question'],
           size=d.fit(k['question'], 712, 190, 20, 14, 1.7), bold=True, color=d.p['text'], line=1.7)

    d.rect(s, 844, 396, 376, 252, fill='#F1F8E9', line=d.p['secondary'], radius=0.05)
    d.text(s, 868, 414, 328, 24, '小使者提示', size=16, bold=True, color='#33691E')
    tip = '这一关考的是「%s」，对应我要学会的本领：%s。\n答案要写在任务单的书写区。' % (
        task_short(k) + '题', k['element'])
    d.text(s, 868, 444, 328, 184, tip, size=d.fit(tip, 328, 184, 15, 11, 1.7),
           color='#33691E', line=1.7)


def pg_attempt(d, k, idx):
    s = d.blank()
    d.header(s, '%02d' % idx, '挑战%s：我来闯关，动笔试一试' % '一二三四'[idx - 1],
             tag='独自闯关 · 3 分钟')
    d.text(s, 60, 126, 1160, 30, '先自己写，写不出来就顺着支架一格一格填——每格只写一句话。',
           size=17, color='#546E7A')
    rows = k['scaffold']
    top, bottom = 172, 560
    gap = 10
    rh = min(78, (bottom - top - gap * (len(rows) - 1)) / len(rows))
    for i, row in enumerate(rows):
        y = top + i * (rh + gap)
        d.rect(s, 60, y, 1160, rh, fill='#FFFFFF', line=d.p['border'], radius=0.12)
        d.text(s, 84, y, 300, rh, row, size=16, bold=True, color=d.p['title'], valign=MSO_ANCHOR.MIDDLE)
        d.rect(s, 400, y + rh * 0.62, 780, 1, fill='#CFD8DC', shape=MSO_SHAPE.RECTANGLE)
    d.rect(s, 60, 578, 1160, 56, fill=d.p['light'], line=None, radius=0.25)
    d.text(s, 84, 578, 1112, 56, '我的动笔步骤：① 回到原文　② 按口诀思考　③ 用支架写下来　④ 读一遍改通顺',
           size=15, bold=True, color=d.p['title'], valign=MSO_ANCHOR.MIDDLE)


def pg_method(d, k, idx):
    s = d.blank()
    d.header(s, '%02d' % idx, '挑战%s：方法指导 · %s' % ('一二三四'[idx - 1], k['method']))
    d.text(s, 60, 126, 1160, 30, '%s题不难——跟着四步法走，答案就有骨头也有肉。' % task_short(k),
           size=18, color='#546E7A')
    steps = k.get('method_steps') or []
    colors = [d.p['primary'], d.p['accent'], d.p['primary'], d.p['accent']]
    for i, step in enumerate(steps[:4]):
        x = 60 + i * 288
        d.rect(s, x, 176, 272, 240, fill=d.p['light'], line=None, radius=0.06)
        d.circle(s, x + 20, 196, 44, str(i + 1), colors[i], 20)
        d.text(s, x + 20, 252, 232, 32, step.get('name', ''), size=22, bold=True, color=d.p['title'])
        d.text(s, x + 20, 290, 232, 112, step.get('desc', ''),
               size=d.fit(step.get('desc', ''), 232, 112, 16, 12, 1.6), color=d.p['text'], line=1.6)
    d.band(s, 60, 446, 1160, 118, '口诀记一记', k['mnemonic'], size=26)


def pg_answer(d, k, idx, per):
    s = d.blank()
    d.header(s, '%02d' % idx, '挑战%s：示例答案与评分要点' % '一二三四'[idx - 1], tag='满分 %d 分' % per)
    d.label(s, 60, 126, 700, '示 例 答 案 · 分 步 呈 现')
    steps = k['answer_steps']
    names = [x.get('name', '') for x in (k.get('method_steps') or [])]
    top, bottom = 158, 648
    answer_h = 148
    step_area = bottom - top - answer_h - 16
    gap = 8
    step_h = (step_area - gap * (len(steps) - 1)) / len(steps)
    y = top
    for i, a in enumerate(steps):
        d.rect(s, 60, y, 690, step_h, fill='#FFFFFF', line=d.p['border'], radius=0.10)
        tag = '第%d步%s：' % (i + 1, '（%s）' % names[i] if i < len(names) and names[i] else '')
        d.rich_para(s, 78, y, 654, step_h,
                    [(tag, d.p['primary'], True), (a, d.p['text'], False)],
                    size=d.fit(tag + a, 654, step_h - 14, 15.5, 10, 1.55), line=1.55)
        y += step_h + gap
    d.rect(s, 60, y, 690, answer_h, fill=d.p['light'], line=d.p['primary'], radius=0.08)
    d.rich_para(s, 78, y, 654, answer_h,
                [('完整答案：', d.p['primary'], True), (k['answer_full'], d.p['title'], False)],
                size=d.fit('完整答案：' + k['answer_full'], 654, answer_h - 14, 15.5, 10, 1.6),
                line=1.6)

    d.label(s, 774, 126, 446, '评 分 要 点（共 %d 分）' % per)
    rubric = k['rubric']
    chip_top, chip_bottom = 158, 560
    chip_gap = 8
    chip_h = min(62, (chip_bottom - chip_top - chip_gap * (len(rubric) - 1)) / len(rubric))
    y = chip_top
    for item in rubric:
        label, text = parse_rubric(item)
        d.rect(s, 774, y, 446, chip_h, fill='#FFFDE7', line=None, radius=0.18)
        d.circle(s, 786, y + (chip_h - 42) / 2, 42, label, d.p['accent'], 12)
        d.text(s, 842, y + 6, 362, chip_h - 12, text,
               size=d.fit(text, 362, chip_h - 12, 15, 10.5, 1.45),
               color=d.p['text'], line=1.45, valign=MSO_ANCHOR.MIDDLE)
        y += chip_h + chip_gap
    tip = '对照提示：先看有没有原文依据，再看要点是否齐全，最后读一遍是否通顺。'
    d.rect(s, 774, 578, 446, 70, fill='#F1F8E9', line=d.p['secondary'], radius=0.12)
    d.text(s, 792, 578, 410, 70, tip, size=13, color='#33691E', line=1.5, valign=MSO_ANCHOR.MIDDLE)


def pg_pitfall(d, k, idx):
    s = d.blank()
    d.header(s, '%02d' % idx, '挑战%s：易错提醒与修改提升' % '一二三四'[idx - 1])
    d.rect(s, 60, 140, 570, 330, fill='#FFEBEE', line='#EF9A9A', radius=0.06)
    d.text(s, 84, 162, 522, 28, '易错提醒（避开这些坑）', size=18, bold=True, color='#C62828')
    d.text(s, 84, 202, 522, 244, '\n\n'.join(k['pitfalls']),
           size=d.fit('\n\n'.join(k['pitfalls']), 522, 244, 16, 11, 1.6), color='#37474F', line=1.6)

    d.rect(s, 650, 140, 570, 330, fill=d.p['light'], line=d.p['primary'], radius=0.06)
    d.text(s, 674, 162, 522, 28, '修改提升（任务单「修改区」）', size=18, bold=True, color=d.p['title'])
    d.text(s, 674, 202, 522, 190, k['revision'],
           size=d.fit(k['revision'], 522, 190, 16, 11, 1.6), color=d.p['text'], line=1.6)
    d.text(s, 674, 400, 522, 56, '口诀再念一遍：' + k['mnemonic'], size=14, bold=True,
           color='#E65100', line=1.45)

    d.rect(s, 60, 496, 1160, 112, fill=d.p['title'], line=None, radius=0.10)
    d.circle(s, 88, 524, 56, '★', d.p['accent'], 24)
    d.text(s, 164, 512, 700, 36, '恭喜收集第 %d 把钥匙' % idx, size=22, bold=True, color='#FFD54F')
    d.text(s, 164, 552, 1020, 30, '%s · %s' % (k['name'], k['method']), size=16, color='#FFFFFF')


def pg_summary(d, data, keys):
    s = d.blank(footer=False)
    spring = d.asset('spring')
    d.rect(s, 60, 48, 1160, 188, fill=d.p['title'], line=None, radius=0.04)
    if spring:
        d.picture(s, spring, 60, 48, 1160, 188)
        d.scrim(s, 60, 48, 1160, 188,
                [(0.0, '#1B5E20', 0.74), (0.65, '#1B5E20', 0.22), (1.0, '#1B5E20', 0.0)], 0)
    d.text(s, 96, 76, 760, 52, '春天回来了！', size=38, bold=True, color='#FFFFFF', shadow=True, line=1.1)
    d.text(s, 96, 148, 760, 30, '四把钥匙都集齐——你就是童话阅读小达人', size=19,
           color='#FFF8E1', shadow=True, line=1.3)

    pairs = [('%s · %s：' % (k['name'], k['method']), k['mnemonic']) for k in keys]
    d.card_rich(s, 60, 256, 630, 252, '四把钥匙 · 方法汇总', pairs, size=14, line=1.95)
    evals = [e.lstrip('0123456789.、 ') for e in data.get('self_eval', [])] or \
            ['%s：我能用对方法。' % k['name'] for k in keys]
    ev_pairs = [('%d. ' % (i + 1), e + '　☆☆☆') for i, e in enumerate(evals)]
    d.card_rich(s, 710, 256, 510, 252, '我的自评表（在任务单上打☆）', ev_pairs,
                fill=d.p['cream'], border='#FFD54F', size=13, line=1.75, label_color='#8D6E63')

    d.rect(s, 60, 528, 1160, 110, fill=d.p['title'], line=None, radius=0.08)
    d.text(s, 84, 528, 120, 110, '课后延伸', size=15, bold=True, color='#FFD54F', valign=MSO_ANCHOR.MIDDLE)
    hw = '　'.join(data.get('homework', []))
    d.text(s, 204, 528, 992, 110, hw,
           size=d.fit(hw, 992, 92, 14, 10, 1.55),
           color='#FFFFFF', line=1.55, valign=MSO_ANCHOR.MIDDLE)


def build_pptx(data, tokens, path):
    d = Deck(tokens, data)
    keys = data['keys']
    per = data.get('score_per_task', 5)
    pg_cover(d, data, keys)
    pg_intro(d, data, keys)
    pg_overview(d, data, keys)
    pg_goals(d, data, keys)
    pg_material(d, data, keys)
    for idx, k in enumerate(keys, 1):
        pg_divider(d, data, k, idx)
        pg_source_question(d, data, k, idx, per)
        pg_attempt(d, k, idx)
        pg_method(d, k, idx)
        pg_answer(d, k, idx, per)
        pg_pitfall(d, k, idx)
    pg_summary(d, data, keys)
    d.total = d.n
    d.prs.save(path)
    return d.n, d.warnings


# ---------------- DOCX ----------------

def docx_setup(doc, title):
    sec = doc.sections[0]
    sec.top_margin = Inches(0.8)
    sec.bottom_margin = Inches(0.8)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.name = 'SimHei'
    r.font.color.rgb = DocxRGB(0x1B, 0x5E, 0x20)


def add_para(doc, text, size=12, bold=False, indent=True):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = 'SimSun'
    return p


def add_table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = 'Table Grid'
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = str(h)
        for p in c.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.name = 'SimHei'
                r.font.size = Pt(10.5)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10.5)
                    r.font.name = 'SimSun'
    return t


def build_lesson(data, path):
    doc = Document()
    docx_setup(doc, '%s · 阅读答题技巧教案逐字稿' % data['scenario'])
    keys = data['keys']
    per = data.get('score_per_task', 5)
    total = data.get('total_score', len(keys) * per)
    add_table(doc, ['项目', '内容'], [
        ['课题', '%s（%s）' % (data['scenario'], data['textbook'])],
        ['年级／课时', '%s／%d 分钟' % (data.get('grade', '四年级'), data.get('duration_min', 40))],
        ['学生角色', data['role']],
        ['配套', '学生课件；学生任务单'],
        ['分值', '每题 %d 分，四题合计 %d 分' % (per, total)],
    ])
    doc.add_heading('一、教学目标', level=1)
    for k in keys:
        add_para(doc, '· 能用「%s」完成%s：%s' % (k['method'], k['task'], k['element']))
    doc.add_heading('二、教学重难点', level=1)
    add_para(doc, '重点：四类题型的答题方法与规范表达；难点：答案有原文依据、概括时抓住关键转折。')
    doc.add_heading('三、教学过程逐字稿（%d 分钟）' % data.get('duration_min', 40), level=1)
    add_para(doc, '【导入 · 情境发布（3 分钟）】师：同学们，今天老师收到一封求助信……'
                  '我们的花园又被冬天困住了。今天我们化身——预设：%s！' % data['role'])
    for idx, k in enumerate(keys, 1):
        doc.add_heading('（%s）%s（8 分钟）' % ('一二三四'[idx - 1], k['task']), level=2)
        add_para(doc, '师：第 %d 关，%s！这一关我要学会的是——预设：%s。' % (idx, k['name'], k['element']))
        add_para(doc, '师：请看原文（课件呈现）。题目是——预设：（齐读题干）%s' % k['question'])
        add_para(doc, '师：翻开任务单，按支架先自己写一写，计时 3 分钟，开始！'
                      '（巡视指导：先圈关键词，再按口诀一步一步写。）')
        add_para(doc, '师：时间到。方法口诀念起来——预设：（齐读）%s' % k['mnemonic'])
        for j, a in enumerate(k['answer_steps'], 1):
            add_para(doc, '师：第 %d 步，谁来说？预设：%s' % (j, a))
        add_para(doc, '师：完整答案就是——%s' % k['answer_full'])
        add_para(doc, '师：%d 分怎么评？对照评分要点：%s' % (per, '；'.join(k['rubric'])))
        add_para(doc, '师：最容易丢分的三个坑——预设：（齐读）%s' % '；'.join(k['pitfalls']))
        add_para(doc, '师：现在进入任务单「修改区」，%s。改完收集——预设：%s！' % (k['revision'], k['name']))
    doc.add_heading('四、汇总 · 评价 · 总结（5 分钟）', level=1)
    add_para(doc, '师：四把钥匙都到手了！方法大集合，四句口诀连起来读——')
    for k in keys:
        add_para(doc, '· %s · %s：%s' % (k['name'], k['method'], k['mnemonic']))
    add_para(doc, '师：翻开任务单自评表，对照四条标准给自己打星。'
                  '答题三金律永远别忘：%s。' % '、'.join(data.get('golden_rules', [])))
    doc.add_heading('五、板书设计', level=1)
    add_para(doc, '%s —— %s' % (data['scenario'], data['role']), indent=False)
    for k in keys:
        add_para(doc, '%s｜%s：%s' % (k['name'], k['method'], k['mnemonic']), indent=False)
    add_para(doc, '答题三金律：%s' % ' · '.join(data.get('golden_rules', [])), indent=False)
    doc.add_heading('六、时间分配总表', level=1)
    add_table(doc, ['环节', '时间', '主要活动'], [
        ['导入 · 情境发布', '3 分钟', '求助信、角色、目标与金律'],
    ] + [[k['task'], '8 分钟', '情境 1′＋原文题目 1′＋尝试 3′＋方法示例 2′＋评分修改 1′'] for k in keys] + [
        ['汇总评价总结', '5 分钟', '口诀汇总 2′＋自评 2′＋延伸 1′'],
        ['合计', '%d 分钟' % data.get('duration_min', 40), '——'],
    ])
    doc.add_heading('七、答案与评分要点总表（与课件、任务单一致）', level=1)
    for k in keys:
        add_para(doc, '【%s · %s（%d 分）】' % (k['task'], k['name'], per), bold=True, indent=False)
        add_para(doc, '题干：%s' % k['question'])
        add_para(doc, '方法口诀（%s）：%s' % (k['method'], k['mnemonic']))
        for j, a in enumerate(k['answer_steps'], 1):
            add_para(doc, '第%d步：%s' % (j, a))
        add_para(doc, '完整答案：%s' % k['answer_full'])
        add_para(doc, '评分要点：%s' % '；'.join(k['rubric']))
        add_para(doc, '易错提醒：%s' % '；'.join(k['pitfalls']))
    doc.add_heading('八、附：任务原文（%s）' % data['textbook'], level=1)
    for k in keys:
        add_para(doc, '%s：%s' % (k['task'], k['source']))
    doc.save(path)


def build_worksheet(data, path):
    """任务单：显式不写 answer_steps / answer_full。"""
    doc = Document()
    docx_setup(doc, '%s · 学生任务单' % data['scenario'])
    keys = data['keys']
    per = data.get('score_per_task', 5)
    add_para(doc, '姓名：__________　班级：__________　日期：__________', indent=False)
    add_para(doc, '亲爱的%s：%s！请依次完成四项挑战，收集四把钥匙，把春天带回花园。'
             % (data['role'], data.get('worksheet_intro', '花园又被冬天困住了')))
    add_table(doc, ['项目', '内容'], [
        ['阅读材料', data['textbook']],
        ['闯关任务', '；'.join(k['task'] for k in keys)],
        ['钥匙收集', '、'.join(k['name'] for k in keys)],
        ['分值', '每题 %d 分，四题合计 %d 分' % (per, data.get('total_score', len(keys) * per))],
        ['答题三金律', ' · '.join(data.get('golden_rules', []))],
    ])
    for k in keys:
        doc.add_heading(k['task'] + '（%s · %d 分）' % (k['name'], per), level=1)
        add_para(doc, '我要学会：%s' % k['element'], indent=False)
        add_para(doc, '原文（%s）：%s' % (data['textbook'], k['source']))
        add_para(doc, '题目（%d 分）：%s' % (per, k['question']))
        add_para(doc, '方法口诀（%s）：%s' % (k['method'], k['mnemonic']))
        add_para(doc, '我的答题支架：', bold=True, indent=False)
        add_table(doc, ['我的记录', '书写区'], [[x, ''] for x in k['scaffold']])
        add_para(doc, '完整作答区：', bold=True, indent=False)
        add_table(doc, ['书写区'], [[''] for _ in range(3)])
        add_para(doc, '评分要点（自查用）：%s' % '；'.join(k['rubric']))
        add_para(doc, '修改区（%s）：' % k['revision'], bold=True, indent=False)
        add_table(doc, ['书写区'], [[''] for _ in range(2)])
        add_para(doc, '避坑提示：%s' % '；'.join(k['pitfalls']))
        add_para(doc, '本关自评：☆☆☆')
    doc.add_heading('四把钥匙 · 方法汇总', level=1)
    add_table(doc, ['钥匙', '方法', '口诀'],
              [[k['name'], k['method'], k['mnemonic']] for k in keys])
    doc.add_heading('学生自评表', level=1)
    add_table(doc, ['自评内容', '自评'],
              [[x, '☆☆☆'] for x in data.get('self_eval', [k['name'] + '：我能用对方法。' for k in keys])]
              + [['我的四关合计得分（满分 %d 分）' % data.get('total_score', len(keys) * per), '　　　分']])
    doc.add_heading('课后延伸', level=1)
    for i, h in enumerate(data.get('homework', []), 1):
        add_para(doc, '%d. %s' % (i, h))
    doc.save(path)


def demo_content():
    """内置示例：与真实课件同一套术语，便于直接跑通验证。"""
    return {
        "scenario": "重建巨人的花园", "role": "花园小使者",
        "textbook": "统编版《巨人的花园》（王尔德）", "grade": "四年级",
        "duration_min": 40, "score_per_task": 5, "total_score": 20,
        "golden_rules": ["依托文本", "贴合主旨", "积极向上"],
        "worksheet_intro": "花园又被冬天困住了",
        "intro_letter": ["我的花园又被冬天困住了。",
                         "只有集齐「想象、提问、品读、梳理」四把钥匙，春天才会回来。",
                         "今天，你们化身——花园小使者！",
                         "依次完成四项挑战，把春天带回花园。"],
        "task_intro": {"1": "花园被冬天封锁的第一现场——小使者先要「看见」那座寒冷的花园，把文字变成脑海里的画面，第一把钥匙才会亮起来。",
                       "2": "冬天的谜团太多了：春天为什么不来？雪和霜为什么高兴？学会问，才找得到答案。",
                       "3": "桃花、小鸟、花儿都回来了，作家是怎么把这一幕写活的？品一品词句，第三把钥匙就在这里。",
                       "4": "故事从赶走到拆墙，中间发生了什么？把起因、经过、结果理清楚，最后一把钥匙到手。"},
        "keys": [
            {"name": "想象之钥", "method": "想象画面四步法", "mnemonic": "一圈词语，二放电影，三找感觉，四写清楚。",
             "element": "边读边想象画面，感受自然之美（四上第一单元）", "task": "挑战一 · 想象画面题",
             "question": "读下面的句子，边读边想象：你仿佛看到了怎样的画面？听到了什么声音？有什么感受？把想象到的画面写下来。",
             "source": "春天来了，乡下到处开着小花，到处有小鸟歌唱。单单在巨人的花园里，仍旧是冬天的景象。小鸟不肯在他的花园里唱歌，因为那里没有孩子们的踪迹；桃树也忘了开花；偶尔有一朵美丽的花从草丛中伸出头来，可是一看见那块布告牌，就马上缩回到地里睡觉去了。高兴的只有雪和霜两位。他们嚷道：“春天把这个花园忘记了，我们一年到头都可以住在这儿啦！”雪用他的白色大衣覆盖着青草，霜把所有的树枝涂成了银色。他们还请来北风同住。北风身上裹着皮衣，整天在花园里呼啸着。",
             "scaffold": ["我圈出的关键词", "我仿佛看到", "我仿佛听到", "我感到"],
             "method_steps": [{"name": "圈词语", "desc": "圈出描写颜色、声音、动作的关键词。"},
                              {"name": "放电影", "desc": "闭上眼睛，把文字变成脑海里的画面。"},
                              {"name": "找感觉", "desc": "想一想这幅画面带给你什么感受。"},
                              {"name": "写清楚", "desc": "用「我仿佛看到……听到……感到……」写下来。"}],
             "answer_steps": ["我圈出「雪用他的白色大衣覆盖着青草」「霜把所有的树枝涂成了银色」「整天在花园里呼啸着」。",
                              "我仿佛看到花园里铺着白茫茫的雪，树枝被霜涂成银色，北风裹着皮衣乱跑。",
                              "我仿佛听到北风「呼呼」的吼叫声，感到花园冷清、荒凉，没有一点生气。"],
             "answer_full": "读了这段话，我仿佛看到巨人的花园里铺着白茫茫的雪，青草被雪盖住了，树枝被霜涂成了银色，北风裹着皮衣在花园里整天呼啸。我仿佛听到北风「呼呼」的吼叫声。我感到花园冷清、荒凉，没有一点生气。",
             "rubric": ["① 圈出或用上关键词 1 分", "② 画面有视觉内容 1 分", "③ 有声音或动态 1 分", "④ 说出感受 1 分", "⑤ 感受与画面一致、语句通顺 1 分"],
             "pitfalls": ["× 只解释词语意思，没有画面", "× 只写「我看到雪」，没有颜色、声音", "× 感受只写「很美」，和画面对不上"],
             "revision": "对照四步口诀和评分要点，把答案改得有颜色、有声音、有感受。"},
            {"name": "提问之钥", "method": "多角度提问四步法", "mnemonic": "一找疑问，二换角度，三挑价值，四答有据。",
             "element": "阅读时尝试从不同角度去思考，提出自己的问题（四上第二单元）", "task": "挑战二 · 多角度提问题",
             "question": "读了《巨人的花园》，请你从不同角度提出两个问题，并选择其中一个问题试着解决。",
             "source": "「我不懂为什么春天来得这样迟，」巨人坐在窗前，望着窗外那凄凉的花园，「我盼望天气快点儿变好。」可是春天始终没有来，夏天也没有来。秋天给每个花园带来了金色果实，但巨人的花园却什么也没有。",
             "scaffold": ["问题一（角度：　　）", "问题二（角度：　　）", "我选择解决的问题是", "我的答案是", "我从文中找到的依据"],
             "method_steps": [{"name": "找疑问", "desc": "把不懂、好奇的地方标出来。"},
                              {"name": "换角度", "desc": "从内容、写法、启示三个角度各想一个问题。"},
                              {"name": "挑价值", "desc": "挑出最有价值的问题。"},
                              {"name": "答有据", "desc": "回到上下文找依据再回答。"}],
             "answer_steps": ["问题一（内容角度）：为什么巨人砌起高墙后，花园里就一直是冬天？",
                              "问题二（写法角度）：作者为什么要写雪、霜、北风、雹在花园里「高兴」？",
                              "我选择解决问题一：因为巨人赶走了孩子，花园里「没有孩子们的踪迹」，小鸟不肯唱歌，桃树也忘了开花——没有孩子，春天就不肯来。依据：「小鸟不肯在他的花园里唱歌，因为那里没有孩子们的踪迹；桃树也忘了开花。」"],
             "answer_full": "我的问题是「为什么巨人砌起高墙后，花园里就一直是冬天？」我联系上下文找到答案：因为巨人赶走了孩子、砌起高墙，花园里没有孩子们的踪迹，小鸟不肯唱歌，桃树也忘了开花；后来孩子们回来，春天也跟着一起回来了。",
             "rubric": ["① 两个问题且来自不同角度 2 分", "② 问题有价值 1 分", "③ 解决时联系原文 1 分", "④ 答案完整、有依据、语句通顺 1 分"],
             "pitfalls": ["× 两个问题同一角度", "× 问题太浅，比如「巨人是谁？」", "× 解答凭空猜，不引原文"],
             "revision": "把问题换个角度再问一次，答案里补上原文依据。"},
            {"name": "品读之钥", "method": "词句赏析四步法", "mnemonic": "一找手法，二抓词语，三说效果，四悟情感。",
             "element": "体会文章准确生动的表达（四上第三单元）", "task": "挑战三 · 词句赏析题",
             "question": "读下面的句子，想想作者是怎样把桃树写得这么可爱的？这样写有什么好处？",
             "source": "桃树看见孩子们回来十分高兴，纷纷用花朵把自己装饰起来，还在孩子们头上轻轻地舞动胳膊。小鸟们快乐地飞舞歌唱。花儿们也从绿草丛中伸出头来。",
             "scaffold": ["我找到的写法", "我圈出的关键词", "这样写的好处", "表达的感情"],
             "method_steps": [{"name": "找手法", "desc": "判断用了什么方法（拟人、比喻、对比）。"},
                              {"name": "抓词语", "desc": "圈出最生动的词。"},
                              {"name": "说效果", "desc": "说说这样写让内容怎么样。"},
                              {"name": "悟情感", "desc": "说说表达了怎样的感情。"}],
             "answer_steps": ["这句话用了拟人的写法。",
                              "我从「十分高兴」「把自己装饰起来」「轻轻地舞动胳膊」看出作者把桃树当成人来写。",
                              "写出了桃树见到孩子们回来时开心、热情的样子，特别生动、可爱。",
                              "表达了作者对孩子们到来的喜悦，也让我们感受到花园重新充满生机。"],
             "answer_full": "这句话用了拟人的写法，从「十分高兴」「把自己装饰起来」「轻轻地舞动胳膊」这些词可以看出，作者把桃树当成人来写，写出了桃树见到孩子们回来时开心、热情的样子，特别生动可爱，表达了作者对孩子们到来的喜悦。",
             "rubric": ["① 指出写法（拟人）1 分", "② 圈出或引用关键词 1 分", "③ 说出生动、形象的效果 2 分", "④ 点出情感或表达作用 1 分"],
             "pitfalls": ["× 只说「用了拟人，很生动」，不说具体词语", "× 只解释词义", "× 不写情感"],
             "revision": "按四步法检查：手法找了吗？词语圈了吗？效果说具体了吗？情感写了吗？"},
            {"name": "梳理之钥", "method": "梳理概括四步法", "mnemonic": "一分三段，二抓主要，三串成句，四查通顺。",
             "element": "了解故事的起因、经过、结果，学习把握文章的主要内容（四上第四单元）", "task": "挑战四 · 故事梳理与概括题",
             "question": "《巨人的花园》这个故事的起因、经过、结果分别是什么？再用几句话概括故事的主要内容。",
             "source": "起因：“你们在这儿做什么？”他叱责道。……于是他在花园的四周砌了一道高墙，挂出一块布告牌：禁止入内 违者重惩。经过：孩子们从墙上一个小洞爬进花园里来了……巨人看见这个情景，心也软了。他对自己说：“我多么自私啊！现在我明白为什么春天不肯到这儿来了。”结果：巨人悄悄地走到他后面，轻轻抱起他，放到树枝上。……他拿出一把大斧子，拆除了围墙。从那以后，巨人的花园又成了孩子们的乐园。",
             "scaffold": ["起因", "经过", "结果", "主要内容"],
             "method_steps": [{"name": "分三段", "desc": "按起因、经过、结果分成三段。"},
                              {"name": "抓主要", "desc": "每段只抓最主要的人和事。"},
                              {"name": "串成句", "desc": "用「谁＋做了什么＋结果怎样」连起来。"},
                              {"name": "查通顺", "desc": "读一遍，删掉重复和次要的。"}],
             "answer_steps": ["起因：巨人回到花园，赶走孩子，砌起高墙，不许任何人进来玩。",
                              "经过：花园里从此只有冬天；孩子们从墙洞钻进花园后，春天跟着回来了；只有最远的角落还是冬天，一个小男孩够不到树枝，急得大哭。",
                              "结果：巨人醒悟了，轻轻抱起小男孩放到树上，拆除了围墙，把花园还给孩子们，花园又成了乐园。"],
             "answer_full": "巨人回到花园后把孩子们赶走、砌起高墙，花园里从此只有寒冷的冬天；后来孩子们从墙洞钻进来，春天也跟着回来了。巨人看见角落里的小男孩够不到树枝在哭，明白了「没有孩子就没有春天」，于是轻轻抱起他放到树上，拆除了围墙。从此，巨人的花园又成了孩子们的乐园。",
             "rubric": ["① 起因概括准确 1 分", "② 经过抓住关键转折 2 分", "③ 结果概括准确 1 分", "④ 主要内容完整、通顺，不啰嗦 1 分"],
             "pitfalls": ["× 复述细节，不会概括", "× 漏掉「小男孩哭泣」这个转折", "× 结果只写「很幸福」，不写拆墙"],
             "revision": "把主要内容压成三四句话，删掉次要细节。"},
        ],
        "self_eval": [
            "1. 想象之钥：我能圈出关键词，写出有颜色、有声音的画面。",
            "2. 提问之钥：我能从不同角度提出有价值的问题，并用原文解答。",
            "3. 品读之钥：我能说出写法、圈出词语，讲清好处和情感。",
            "4. 梳理之钥：我能按起因、经过、结果概括主要内容。"
        ],
        "homework": ["★ 用四把钥匙再读《巨人的花园》，把没写完的答案补完整。",
                     "★★ 把四关的答案整理成一张「我的答题卡」，标上对应的钥匙。",
                     "★★★ 用「想象之钥」，给花园「回春」的画面写一段有颜色、有声音的话。"]
    }

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--content')
    ap.add_argument('--demo', action='store_true')
    ap.add_argument('--out-dir', required=True)
    ap.add_argument('--tokens')
    ap.add_argument('--prefix', default='《巨人的花园》阅读答题技巧')
    ap.add_argument('--no-images', action='store_true', help='不插入插画，用于降级验证')
    a = ap.parse_args()
    tokens = DEFAULT_TOKENS
    if a.tokens:
        tokens = json.load(open(a.tokens, encoding='utf-8'))
    if a.no_images:
        tokens = json.loads(json.dumps(tokens))
        tokens.setdefault('visual', {})['illustrations'] = {}
    data = demo_content() if a.demo else json.load(open(a.content, encoding='utf-8'))
    os.makedirs(a.out_dir, exist_ok=True)
    pptx = os.path.join(a.out_dir, a.prefix + '_学生课件.pptx')
    lesson = os.path.join(a.out_dir, a.prefix + '_教案逐字稿.docx')
    sheet = os.path.join(a.out_dir, a.prefix + '_学生任务单.docx')
    pages, warnings = build_pptx(data, tokens, pptx)
    build_lesson(data, lesson)
    build_worksheet(data, sheet)
    print(json.dumps({'pptx': pptx, 'lesson': lesson, 'worksheet': sheet, 'slides': pages,
                      'warnings': warnings,
                      'note': 'worksheet 已跳过 answer_steps / answer_full'},
                     ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
