#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""三件套一致性断言。

用法:
  python consistency_check.py --config check.json
  python consistency_check.py --config check.json --files 课件.pptx 教案.docx 任务单.docx
                              [--labels 课件 教案 任务单]

推荐把文件清单也写进 config.json 的 "files" 字段（用户常会移动/重命名成品，
命令行中文长路径易失效；JSON 用 UTF-8 读取更稳）:
{
 "files": {"课件": "C:/.../课件.pptx", "教案": "C:/.../教案.docx", "任务单": "C:/.../任务单.docx"},
 "must_appear_in_all": ["四句口诀", "四个任务名", ...],
 "must_appear_in": {"教案": ["板书设计"], "任务单": ["修改区"]},
 "must_not_appear_in": {"任务单": ["完整答案：..."]},
 "forbidden_quotes": ["旧版教材特有句"],
 "key_quotes": ["教材原文关键句"]
}
"files" 也接受路径字符串数组（标签自动取文件名）。

输出 ALL-IN / MISS / clean 明细；全绿退出码 0，否则 1。
需要 python-pptx、python-docx（缺失时给明确安装提示）。
"""
import argparse
import json
import os
import sys


def load_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.pptx', '.ppt'):
        from pptx import Presentation
        prs = Presentation(path)
        chunks = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    chunks.append(shape.text_frame.text)
                if getattr(shape, 'has_table', False):
                    for row in shape.table.rows:
                        for cell in row.cells:
                            chunks.append(cell.text)
        return '\n'.join(chunks)
    if ext in ('.docx', '.doc'):
        from docx import Document
        doc = Document(path)
        chunks = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    chunks.append(cell.text)
        return '\n'.join(chunks)
    if ext in ('.md', '.txt', '.html', '.htm'):
        with open(path, encoding='utf-8', errors='ignore') as f:
            return f.read()
    raise ValueError('不支持的文件类型: %s' % path)


def hit(text, key):
    return key in text or key.replace(' ', '') in text.replace(' ', '')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--config', required=True)
    ap.add_argument('--files', nargs='*', default=None)
    ap.add_argument('--labels', nargs='*', default=None)
    args = ap.parse_args()

    with open(args.config, encoding='utf-8') as f:
        cfg = json.load(f)

    file_map = cfg.get('files')
    if isinstance(file_map, dict):
        labels = list(file_map.keys())
        paths = [file_map[k] for k in labels]
    elif isinstance(file_map, list):
        paths = file_map
        labels = [os.path.splitext(os.path.basename(p))[0] for p in paths]
    elif args.files:
        paths = args.files
        labels = args.labels or [os.path.splitext(os.path.basename(p))[0] for p in paths]
    else:
        print('错误：config 缺少 "files" 字段，且未传 --files。')
        return 2
    if len(labels) != len(paths):
        print('labels 数量与 files 不一致')
        return 2
    missing = [p for p in paths if not os.path.exists(p)]
    if missing:
        print('错误：以下文件不存在（成品可能被移动或重命名，请更新 config 的 "files"）：')
        for p in missing:
            print('  ', p)
        return 2

    try:
        texts = {label: load_text(path) for label, path in zip(labels, paths)}
    except ImportError as e:
        print('缺少依赖，请用当前解释器安装: python -m pip install python-pptx python-docx --no-cache-dir')
        print('原始错误:', e)
        return 2

    failures = []
    print('文件:', ', '.join('%s=%s' % (l, os.path.basename(p)) for l, p in zip(labels, paths)))

    for group in ('must_appear_in_all', 'key_quotes'):
        keys = cfg.get(group, [])
        misses = [(k, l) for k in keys for l, t in texts.items() if not hit(t, k)]
        print('%s: %d 键 × %d 文件 = %d 次比对 -> %s' %
              (group, len(keys), len(texts), len(keys) * len(texts),
               'ALL-IN' if not misses else 'MISS %s' % misses))
        failures.extend(misses)

    for label, keys in (cfg.get('must_appear_in') or {}).items():
        if label not in texts:
            failures.append(('must_appear_in 标签不存在', label))
            continue
        misses = [(k, label) for k in keys if not hit(texts[label], k)]
        print('must_appear_in[%s]: %d 键 -> %s' %
              (label, len(keys), 'ALL-IN' if not misses else 'MISS %s' % misses))
        failures.extend(misses)

    for label, keys in (cfg.get('must_not_appear_in') or {}).items():
        if label not in texts:
            failures.append(('must_not_appear_in 标签不存在', label))
            continue
        leaks = [(k, label) for k in keys if hit(texts[label], k)]
        print('must_not_appear_in[%s]: %d 键 -> %s' %
              (label, len(keys), 'clean' if not leaks else 'LEAK %s' % leaks))
        failures.extend(leaks)

    bad = [(k, l) for k in (cfg.get('forbidden_quotes') or []) for l, t in texts.items() if hit(t, k)]
    print('forbidden_quotes: %d 键 × %d 文件 -> %s' %
          (len(cfg.get('forbidden_quotes') or []), len(texts), 'clean' if not bad else 'LEAK %s' % bad))
    failures.extend(bad)

    print('结论:', 'PASS（全绿，可交付）' if not failures else 'FAIL（先补齐上述 MISS/LEAK 再复验）')
    return 0 if not failures else 1


if __name__ == '__main__':
    sys.exit(main())
