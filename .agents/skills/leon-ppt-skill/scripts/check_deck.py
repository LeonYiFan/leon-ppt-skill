# -*- coding: utf-8 -*-
"""Sanity-check a deck built from this template.

    python check_deck.py deck.pptx [--manifest deck.audit.json] [--allow 术语]

Checks that are cheap and catch real mistakes:
  * slide size and count
  * paragraphs that are too long for the 要点块 (they will overflow the card)
  * pages without speaker notes
  * text still showing the placeholder wording of the template
Exit code 1 when an error-severity item is found.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from zipfile import BadZipFile

from pptx import Presentation
from pptx.exc import PackageNotFoundError

sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
import ppt_template as T  # noqa: E402
from artifact_privacy import core_text, slide_text  # noqa: E402
from delivery_audit import audit_package  # noqa: E402

# The bundled example uses fictional content.
SAMPLE_TERMS = []
PLACEHOLDERS = ['页面标题：一句话给结论', '卡片标题', '一句话说明', '面板说明',
                '读图要点', 'XX 会议汇报', 'English title of the paper as published',
                '章节标题', '本页一句话结论', '三条并列的判断标准',
                '单位名称', '第一作者：XXX', '汇报人：XXX', 'DOI:10.xxxx']

# 中文 AI 腔的典型句式。命中不一定错，但每一处都值得再看一眼：
# 幻灯片上一句话占一屏，套话会立刻显得不是本人在讲。
AI_TELLS = [
    (r'不是[^。；]{1,24}而是', '「不是A而是B」套式'),
    (r'不仅[^。；]{0,16}(更是|而是|还)', '「不仅…更是」套式'),
    (r'至关重要|重中之重|举足轻重', '放大重要性'),
    (r'赋能|助力|抓手|闭环|生态|全方位|多维度|深层次', '空泛动词/名词'),
    (r'众所周知|毋庸置疑|不言而喻', '无信息开场'),
    (r'值得注意的是|需要强调的是|不难发现', '提示语，直接说事'),
    (r'综上所述|总而言之|总的来说', '总结套话（模板已有总结条）'),
    (r'本文将|本页将|接下来我们来看|下面让我们', '论文腔/主持腔'),
    (r'旨在|以期|进而为.{0,10}提供|奠定.{0,6}基础', '论文腔'),
    (r'显著提升|大幅改善|极大提高(?!\s*$)', '无数据的效果形容'),
    (r'——', '破折号：改成逗号或分句'),
    (r'!!!|！！', '多个感叹号'),
]

EXPECTED_SIZE = (12192000, 6858000)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--manifest', type=Path, help='本地交付清单，包含可编辑槽位与证据图来源')
    ap.add_argument('--allow', default='', help='样本词白名单，逗号分隔')
    ap.add_argument('--max-bullet', type=int, default=46)
    args = ap.parse_args()

    try:
        manifest = json.loads(args.manifest.read_text(encoding='utf-8-sig')) if args.manifest else None
    except (OSError, UnicodeError, json.JSONDecodeError):
        ap.error('无法读取交付清单')
    if args.manifest and manifest is None:
        ap.error('交付清单应为 JSON 对象')
    errors = list(audit_package(args.pptx, manifest))
    try:
        prs = Presentation(args.pptx)
    except (BadZipFile, OSError, ValueError, PackageNotFoundError):
        print('ERROR  无法打开 PPTX')
        return 1
    allow = {s.strip() for s in args.allow.split(',') if s.strip()}
    warnings = []
    sample_terms = [term for term in SAMPLE_TERMS if term not in allow]

    def check_sample(location, text):
        for term in sample_terms:
            if term.casefold() in text.casefold():
                errors.append('%s: 残留样本内容 %r' % (location, term))

    for field, value in core_text(prs):
        check_sample('文档属性 ' + field, value)

    size = (prs.slide_width, prs.slide_height)
    if size != EXPECTED_SIZE:
        errors.append('slide size %s != %s (template is 33.87 x 19.05 cm)'
                      % (size, EXPECTED_SIZE))

    for i, slide in enumerate(prs.slides, 1):
        roles = set()
        if not errors and manifest:
            by_id = {shape.shape_id: shape.name for shape in T.iter_shapes(slide.shapes)}
            roles = {by_id[item['shape_id']] for item in manifest['slides'][str(i)].get('shapes', [])
                     if item['role'] in {'title', 'body', 'summary', 'label', 'caption'}}
        texts = []
        for name, body in slide_text(slide):
            if not body.strip():
                continue
            flat = body.replace('\n', ' ')
            texts.append((name, flat))
            check_sample('P%d %s' % (i, name), flat)
            for ph in PLACEHOLDERS:
                if flat.strip().startswith(ph):
                    warnings.append('P%d %s: 仍是占位文字 %r' % (i, name, ph[:12]))
            if name == '要点1':
                for line in body.split('\n'):
                    if len(line) > args.max_bullet:
                        warnings.append('P%d 要点1: %d 字，可能换行溢出（建议 ≤%d）'
                                        % (i, len(line), args.max_bullet))
            if name in ('页标题', '要点1', '总结文字', '卡片标题', '卡片脚注') or name.startswith('组件 ') or name in roles:
                for pattern, why in AI_TELLS:
                    if re.search(pattern, flat):
                        warnings.append('P%d %s: %s → %r'
                                        % (i, name, why, flat[:24]))
                if re.search(r'\*\*', flat):
                    warnings.append('P%d %s: 残留 ** 标记（只会标红一次、也别漏删）'
                                    % (i, name))
        has_title = any(n == '页标题' for n, _ in texts)
        if has_title and not any(n == '讲稿备注' for n, _ in texts):
            warnings.append('P%d: 没有讲稿备注' % i)

    print('%s  %d slides  %.2f x %.2f cm'
          % (args.pptx, len(prs.slides),
             prs.slide_width / T.CM, prs.slide_height / T.CM))
    for line in errors:
        print('ERROR  ' + line)
    for line in warnings:
        print('warn   ' + line)
    print('%d error(s), %d warning(s)' % (len(errors), len(warnings)))
    return 1 if errors else 0


if __name__ == '__main__':
    sys.exit(main())
