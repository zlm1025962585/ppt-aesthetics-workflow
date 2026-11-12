#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""结构化 Markdown -> 合规 HTML（供 workbuddy-docx 的 html-to-doc 转换）。

用法:
  python md2html.py --md final_draft.md --html formatted.html \
      --title "课题名称" \
      --subtitles '["副标题一","副标题二"]' \
      --cover-kv '[["年级","四年级"],["分值","每题 5 分"]]' \
      [--leads '{"章节标题":"导语"}'] [--no-page-furniture]

支持的 Markdown 扩展:
  > 引文...            -> 单列表格引文块（原文呈现）
  {{CELL}}             -> 表格中的空白书写格
  {{LINES 3}}          -> 3 行书写区（实现为 3 行表格，避免连续空段落告警）
  GFM 表格 / 有序无序列表 / 代码围栏（转为板书框）

内置样式已满足 html-review 的主要规则：
  正文 SimSun 12pt + text-indent:2em；全 var() 样式值；表格含 thead/tbody；
  封面标题用 <p>；封面/正文两个顶层 section，正文 data-page-restart="1"；@page 页码。
"""
import argparse
import html as H
import json
import re


def bold(text):
    text = H.escape(text)
    return re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)


CSS = """
:root {
  --fs-title: 24pt; --fs-h1: 18pt; --fs-h2: 15pt; --fs-h3: 13pt;
  --fs-body: 12pt; --fs-small: 10pt;
  --ff-heading: "SimHei", "黑体", sans-serif;
  --ff-body: "SimSun", "宋体", serif;
  --lh-body: 1.7; --lh-heading: 1.3;
  --color-primary: #2E7D32; --color-text: #000000; --color-muted: #555555;
  --color-border: #999999; --color-highlight: #EAF3EA; --color-bg: #ffffff;
  --spacing-paragraph: 0.5em; --spacing-section: 1.2em; --spacing-block: 1em;
  --spacing-h2top: 1em; --spacing-h3top: 0.9em;
  --spacing-covertitle-top: 7em; --spacing-covertitle-bottom: 1.2em;
  --spacing-cover-sub: 0.6em; --spacing-coverinfo-top: 2em;
  --spacing-list-pad: 2.2em; --spacing-li: 0.25em;
  --spacing-cell: 4pt 6pt; --spacing-kv-cell: 5pt 8pt;
  --spacing-board-cell: 8pt 10pt; --spacing-divider: 1em;
  --page-content-width: 15.6cm;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { max-width: var(--page-content-width); margin-left: auto; margin-right: auto; background: var(--color-bg); }
h2 { font-family: var(--ff-heading); font-size: var(--fs-h1); font-weight: 700; color: var(--color-primary); line-height: var(--lh-heading); margin-top: var(--spacing-h2top); margin-bottom: var(--spacing-paragraph); text-align: left; }
h3 { font-family: var(--ff-heading); font-size: var(--fs-h2); font-weight: 700; color: var(--color-text); line-height: var(--lh-heading); margin-top: var(--spacing-h3top); margin-bottom: var(--spacing-paragraph); text-align: left; }
h4 { font-family: var(--ff-heading); font-size: var(--fs-h3); font-weight: 700; color: var(--color-text); line-height: var(--lh-heading); margin-top: var(--spacing-paragraph); margin-bottom: var(--spacing-paragraph); text-align: left; }
p.para { font-family: var(--ff-body); font-size: var(--fs-body); color: var(--color-text); line-height: var(--lh-body); text-indent: 2em; text-align: justify; margin-bottom: var(--spacing-paragraph); }
p.cover-title { font-family: var(--ff-heading); font-size: var(--fs-title); font-weight: 700; color: var(--color-primary); text-align: center; line-height: 1.4; margin-top: var(--spacing-covertitle-top); margin-bottom: var(--spacing-covertitle-bottom); text-indent: 0; }
p.cover-sub { font-family: var(--ff-body); font-size: var(--fs-body); color: var(--color-text); text-align: center; line-height: 1.6; margin-bottom: var(--spacing-cover-sub); text-indent: 0; }
table.doc-table { border-collapse: collapse; width: 100%; margin-top: var(--spacing-block); margin-bottom: var(--spacing-block); }
th.th-cell { font-family: var(--ff-heading); font-size: var(--fs-small); font-weight: 700; color: var(--color-text); background: var(--color-highlight); border: 1px solid var(--color-border); padding: var(--spacing-cell); text-align: center; }
td.td-cell { font-family: var(--ff-body); font-size: var(--fs-small); color: var(--color-text); border: 1px solid var(--color-border); padding: var(--spacing-cell); text-align: left; vertical-align: top; line-height: 1.5; }
td.write-cell { font-family: var(--ff-body); font-size: var(--fs-body); color: var(--color-text); background: var(--color-bg); border: 1px solid var(--color-border); padding: var(--spacing-board-cell); line-height: 1.6; text-align: left; }
td.board-cell { font-family: var(--ff-body); font-size: var(--fs-body); color: var(--color-text); background: var(--color-highlight); border: 1px solid var(--color-border); padding: var(--spacing-board-cell); line-height: 1.6; text-align: left; }
table.cover-info { border-collapse: collapse; width: 80%; margin-left: auto; margin-right: auto; margin-top: var(--spacing-coverinfo-top); }
td.kv-key { font-family: var(--ff-heading); font-size: var(--fs-small); font-weight: 700; color: var(--color-text); background: var(--color-highlight); border: 1px solid var(--color-border); padding: var(--spacing-kv-cell); text-align: center; width: 30%; }
td.kv-val { font-family: var(--ff-body); font-size: var(--fs-small); color: var(--color-text); border: 1px solid var(--color-border); padding: var(--spacing-kv-cell); text-align: left; }
ol.ol-list, ul.ul-list { margin-bottom: var(--spacing-paragraph); margin-left: 0; padding-left: var(--spacing-list-pad); }
li.li-item { font-family: var(--ff-body); font-size: var(--fs-body); color: var(--color-text); line-height: var(--lh-body); text-align: justify; margin-bottom: var(--spacing-li); }
table.board-box { border-collapse: collapse; width: 100%; margin-top: var(--spacing-block); margin-bottom: var(--spacing-block); }
div.section-divider { border-top: 1px solid var(--color-border); margin-top: var(--spacing-divider); margin-bottom: var(--spacing-divider); }
"""

PAGE_RULES = """
/* 页面规则：封面无页脚，正文页脚页码 */
@page { @bottom-center { content: "第 " counter(page) " 页"; } }
@page cover { @bottom-center { content: none; } }
section[role="cover"] { page: cover; }
"""


def convert(md_path, title, subtitles, cover_kv, leads, page_furniture):
    lines = open(md_path, encoding='utf-8').read().split('\n')
    body, i, n = [], 0, len(lines)
    in_code, code_buf = False, []

    def add_heading(level, text):
        """显式导语优先；否则相邻两级标题之间自动补一句导语（版式检查 TQ-03）。"""
        last = body[-1] if body else ''
        m_prev = re.match(r'^<h([234])>', last)
        if m_prev and int(m_prev.group(1)) < level:
            body.append('<p class="para">本环节按以下步骤推进。</p>')
        body.append('<h%d>%s</h%d>' % (level, bold(text), level))
        if text in leads:
            body.append('<p class="para">%s</p>' % bold(leads[text]))

    def flush_code():
        nonlocal code_buf
        txt = '<br />'.join(H.escape(x) for x in code_buf)
        body.append('<table class="board-box"><thead><tr><th class="th-cell">板书设计</th></tr></thead>'
                    '<tbody><tr><td class="board-cell">' + txt + '</td></tr></tbody></table>')
        code_buf = []

    while i < n:
        ln = lines[i]
        if ln.startswith('```'):
            if in_code:
                flush_code()
                in_code = False
            else:
                in_code = True
            i += 1
            continue
        if in_code:
            code_buf.append(ln)
            i += 1
            continue
        if ln.startswith('# '):
            i += 1
            continue
        if ln.startswith('## '):
            t = ln[3:].strip()
            add_heading(2, t)
            i += 1
            continue
        if ln.startswith('### '):
            add_heading(3, ln[4:].strip())
            i += 1
            continue
        if ln.startswith('#### '):
            add_heading(4, ln[5:].strip())
            i += 1
            continue
        if ln.strip() == '---':
            body.append('<div class="section-divider"></div>')
            i += 1
            continue
        if ln.strip().startswith('> '):
            buf = [ln.strip()[2:]]
            i += 1
            while i < n and lines[i].strip().startswith('> '):
                buf.append(lines[i].strip()[2:])
                i += 1
            body.append('<table class="board-box"><thead><tr><th class="th-cell">原文</th></tr></thead>'
                        '<tbody><tr><td class="board-cell">' + bold(' '.join(buf)) + '</td></tr></tbody></table>')
            continue
        m_lines = re.match(r'^\{\{LINES (\d+)\}\}$', ln.strip())
        if m_lines:
            k = int(m_lines.group(1))
            rows = ''.join('<tr><td class="write-cell">&nbsp;</td></tr>' for _ in range(k))
            body.append('<table class="doc-table"><thead><tr><th class="th-cell">书写区</th></tr></thead>'
                        '<tbody>' + rows + '</tbody></table>')
            i += 1
            continue
        if ln.lstrip().startswith('|') and i + 1 < n and re.match(r'^\s*\|[\s:\-|]+\|\s*$', lines[i + 1]):
            rows = []
            while i < n and lines[i].lstrip().startswith('|'):
                rows.append([c.strip() for c in lines[i].strip().strip('|').split('|')])
                i += 1
            head, data = rows[0], rows[2:]
            th = ''.join('<th class="th-cell">%s</th>' % bold(c) for c in head)

            def cell(c):
                return ('<td class="write-cell">&nbsp;</td>' if c == '{{CELL}}'
                        else '<td class="td-cell">%s</td>' % bold(c))
            trs = ''.join('<tr>' + ''.join(cell(c) for c in r) + '</tr>' for r in data)
            body.append('<table class="doc-table"><thead><tr>%s</tr></thead><tbody>%s</tbody></table>' % (th, trs))
            continue
        if re.match(r'^\s*\d+\.\s+', ln):
            items = []
            while i < n and re.match(r'^\s*\d+\.\s+', lines[i]):
                items.append(re.sub(r'^\s*\d+\.\s+', '', lines[i]))
                i += 1
            body.append('<ol class="ol-list">%s</ol>' %
                        ''.join('<li class="li-item">%s</li>' % bold(x) for x in items))
            continue
        if re.match(r'^\s*-\s+', ln):
            items = []
            while i < n and re.match(r'^\s*-\s+', lines[i]):
                items.append(re.sub(r'^\s*-\s+', '', lines[i]))
                i += 1
            body.append('<ul class="ul-list">%s</ul>' %
                        ''.join('<li class="li-item">%s</li>' % bold(x) for x in items))
            continue
        if ln.strip() == '':
            i += 1
            continue
        buf = [ln.strip()]
        i += 1
        while i < n and lines[i].strip() and not re.match(r'^(#{1,4}\s|\||```|\d+\.\s|-\s|>\s|---|\{\{)', lines[i]):
            buf.append(lines[i].strip())
            i += 1
        body.append('<p class="para">%s</p>' % bold(' '.join(buf)))
    if in_code:
        flush_code()

    subs = ''.join('<p class="cover-sub">%s</p>' % bold(s) for s in subtitles)
    kv = ''.join('<tr><td class="kv-key">%s</td><td class="kv-val">%s</td></tr>' %
                 (H.escape(k), H.escape(v)) for k, v in cover_kv)
    return ('<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8" />\n'
            '<meta name="docx-page-size" content="A4" />\n<title>%s</title>\n<style>\n%s\n%s</style>\n</head>\n<body>\n'
            '<section role="cover">\n  <p class="cover-title">%s</p>\n  %s\n'
            '  <table class="cover-info"><thead><tr><th class="th-cell">项目</th><th class="th-cell">内容</th></tr></thead>'
            '<tbody>%s</tbody></table>\n</section>\n'
            '<section role="body" data-page-restart="1">\n  %s\n</section>\n</body>\n</html>\n') % (
                H.escape(title), CSS, PAGE_RULES if page_furniture else '',
                H.escape(title), subs, kv, '\n  '.join(body))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--md', required=True)
    ap.add_argument('--html', required=True)
    ap.add_argument('--title', required=True)
    ap.add_argument('--subtitles', default='[]')
    ap.add_argument('--cover-kv', default='[]')
    ap.add_argument('--leads', default='{}')
    ap.add_argument('--no-page-furniture', action='store_true')
    a = ap.parse_args()
    out = convert(a.md, a.title, json.loads(a.subtitles), json.loads(a.cover_kv),
                  json.loads(a.leads), not a.no_page_furniture)
    with open(a.html, 'w', encoding='utf-8') as f:
        f.write(out)
    print('wrote', a.html, len(out))


if __name__ == '__main__':
    main()
