#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""一条命令完成：生成三件套 + 自动派生一致性断言 + 视觉底线校验。

用法:
  python run_all.py --demo --out-dir ./out
  python run_all.py --content content_lock.json --out-dir ./out [--tokens assets/visual_tokens.json] [--prefix 课题名]

断言不需要手写：从内容锁定表里自动提取四句口诀、四个任务名、四把钥匙名、
方法与分值，再检查任务单里没有答案字段。任何一项不过，退出码为 1。
"""
import argparse
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(HERE)


def make_config(data, files):
    keys = data['keys']
    per = data.get('score_per_task', 5)
    must = []
    for k in keys:
        must += [k['name'], k['task'], k['method'], k['mnemonic']]
    must += ['每题 %d 分' % per, data['scenario'], data['role']]
    must += list(data.get('golden_rules', []))
    unseen = [k['answer_full'][:14] for k in keys if len(k.get('answer_full', '')) >= 14]
    cfg = {
        'files': files,
        'must_appear_in_all': must,
        'must_appear_in': {'课件': ['示例答案', '评分要点']},
        'must_not_appear_in': {'任务单': ['完整答案：', '示例答案'] + unseen},
        'forbidden_quotes': list(data.get('forbidden_quotes', [])),
        'key_quotes': list(data.get('key_quotes', [])),
    }
    return cfg


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    return r.returncode, (r.stdout or '') + (r.stderr or '')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--content')
    ap.add_argument('--demo', action='store_true')
    ap.add_argument('--out-dir', required=True)
    ap.add_argument('--tokens', default=os.path.join(SKILL_ROOT, 'assets', 'visual_tokens.json'))
    ap.add_argument('--prefix', default='《巨人的花园》阅读答题技巧')
    ap.add_argument('--no-images', action='store_true')
    ap.add_argument('--min-slides', type=int, default=24)
    ap.add_argument('--max-slides', type=int, default=30)
    a = ap.parse_args()

    build = [sys.executable, os.path.join(HERE, 'build_kit.py'),
             '--out-dir', a.out_dir, '--tokens', a.tokens, '--prefix', a.prefix]
    build += ['--demo'] if a.demo else ['--content', a.content]
    if a.no_images:
        build.append('--no-images')
    rc, out = run(build)
    print(out.strip())
    if rc != 0:
        print('\n结论: FAIL（生成阶段失败）')
        return 1

    data = json.load(open(a.content, encoding='utf-8')) if a.content else None
    if data is None:
        import importlib.util
        spec = importlib.util.spec_from_file_location('build_kit', os.path.join(HERE, 'build_kit.py'))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        data = mod.demo_content()

    files = {
        '课件': os.path.join(a.out_dir, a.prefix + '_学生课件.pptx'),
        '教案': os.path.join(a.out_dir, a.prefix + '_教案逐字稿.docx'),
        '任务单': os.path.join(a.out_dir, a.prefix + '_学生任务单.docx'),
    }
    cfg_path = os.path.join(a.out_dir, 'check.generated.json')
    json.dump(make_config(data, files), open(cfg_path, 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)

    rc1, out1 = run([sys.executable, os.path.join(HERE, 'consistency_check.py'), '--config', cfg_path])
    print(out1.strip())
    rc2, out2 = run([sys.executable, os.path.join(HERE, 'visual_check.py'),
                     '--tokens', a.tokens, '--pptx', files['课件'],
                     '--min-slides', str(a.min_slides), '--max-slides', str(a.max_slides)])
    print(out2.strip())

    ok = (rc1 == 0 and rc2 == 0)
    print('\n===== 总检 =====')
    print('一致性: %s' % ('PASS' if rc1 == 0 else 'FAIL'))
    print('视觉  : %s' % ('PASS' if rc2 == 0 else 'FAIL'))
    for k, v in files.items():
        print('%s: %s%s' % (k, v, '' if os.path.exists(v) else '  <缺失>'))
    print('结论: %s' % ('PASS（三件套可直接使用）' if ok else 'FAIL（按上方明细修正后重跑）'))
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
