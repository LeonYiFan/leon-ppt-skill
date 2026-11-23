# -*- coding: utf-8 -*-
"""Build a paper-report deck from the bundled template deck.

    python build_report.py --plan plan.json --output out.pptx [--template path]

The plan (see references/page-plan.md) lists the pages in order.  Every page is
produced by duplicating one archetype page of ``assets/template-deck.pptx`` and
filling its named shapes, so the result keeps the template's own title bar,
要点块, 卡片, 灰底总结条, 章节页照片, 封面与致谢.

Never overwrites an existing output file.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
from hashlib import sha256
import io
import json
from math import isfinite
import os
from pathlib import Path
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ppt_template as T   # noqa: E402
from artifact_privacy import clean_template, identity_issues  # noqa: E402
from delivery_audit import audit_package  # noqa: E402
from native_distillation import FORM_FIELDS, infer_intent  # noqa: E402
from native_layout_selection import choose_native_layout  # noqa: E402
from native_reference import NativeLibrary  # noqa: E402
import primary_figure_layout  # noqa: E402
from report_manifest import ReportManifest, finalize_slide, native_records  # noqa: E402
import reusable_components  # noqa: E402

ARCHETYPES = {'cover': 0, 'agenda': 1, 'divider': 2, 'bullets': 3,
              'table': 4, 'cards2': 5, 'triple': 6, 'thanks': 7}

ASSETS = Path(__file__).resolve().parents[1] / 'assets'
PORTABLE_LIBRARY = ASSETS / 'portable-library'
PRIVATE_LIBRARY = ASSETS / 'native-library'
DEFAULT_NATIVE_LIBRARY = PORTABLE_LIBRARY if PORTABLE_LIBRARY.is_dir() else PRIVATE_LIBRARY
DEFAULT_TEMPLATE = PORTABLE_LIBRARY / 'template-deck.pptx'
if not DEFAULT_TEMPLATE.is_file():
    DEFAULT_TEMPLATE = ASSETS / 'template-deck.pptx'
DEFAULT_TEMPLATE = str(DEFAULT_TEMPLATE)

COMPONENT_FOR_INTENT = {
    'figure_interpretation': 'gut_detail',
    'figure_metric_comparison': 'gut_metric',
    'parallel_workstreams': 'guide_tracks',
    'risk_response': 'guide_risks',
    'sequential_work': 'flow',
    'same_basis_comparison': 'compare',
    'independent_findings': 'matrix',
    'dated_observations': 'timeline',
    'evidence_to_claim': 'evidence',
    'deliverable_checkpoints': 'milestones',
    'stage_outputs': 'roadmap',
    'decision_thresholds': 'criteria',
}
PRIMARY_FIGURE_COMPONENTS = {'gut_detail', 'gut_metric'}


def need(page, key):
    if key not in page or page[key] in (None, '', []):
        raise ValueError('%s page is missing %r' % (page.get('type'), key))
    return page[key]


BASE = os.getcwd()          # directory of the plan file, set in main()


def fig_path(path):
    """Figure paths in the plan are relative to the plan file."""
    return path if os.path.isabs(path) else os.path.join(BASE, path)


def apply_native_reference(library, slide, spec, *, automatic):
    source = spec.get("source")
    # Existing direct donor references use original page numbers on the local installation.
    if (not automatic and library.root == PORTABLE_LIBRARY and PRIVATE_LIBRARY.is_dir()
            and isinstance(source, str) and not source.startswith("catalog-")):
        library = NativeLibrary(PRIVATE_LIBRARY)
    return library.apply(slide, spec, fig_path)


def native_page(prs, page, library, decision=None):
    """Use a reference page's native geometry without the default body frame."""
    if 'component' in page or page.get('figures'):
        raise ValueError('native pages use native.text, native.tables and native.images')
    spec = need(page, 'native')
    if not isinstance(spec, dict):
        raise ValueError('native must be an object')

    slide = prs.slides.add_slide(prs.slides[ARCHETYPES['bullets']].slide_layout)
    for shape in list(slide.shapes):
        T.drop(shape)
    slide._element.set('showMasterSp', '0')
    remap = apply_native_reference(library, slide, spec, automatic=decision is not None)
    if decision and decision.get("summary_box"):
        archetype = prs.slides[ARCHETYPES["bullets"]]
        summary = T.shape(archetype, "总结条")
        copied = deepcopy(summary._element)
        for node in copied.iter():
            for attr, value in tuple(node.attrib.items()):
                if attr.startswith(T.R) and value in archetype.part.rels:
                    relation = archetype.part.rels[value]
                    if relation.is_external:
                        raise ValueError("总结条原型须使用本地资源")
                    node.set(attr, slide.part.relate_to(relation.target_part, relation.reltype))
        NativeLibrary._adopt(slide, [copied])
        summary = T.shape(slide, "总结条")
        summary.left, summary.top, summary.width, summary.height = (
            T.cm(value) for value in decision["summary_box"]
        )
        T.fill_text(slide, "总结文字", page["summary"], base="111820")
        frame = T.shape(slide, "总结文字").text_frame
        frame.margin_top = frame.margin_bottom = 0
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    return slide, remap


def position_native_selection(shapes, region, canvas):
    """Fit a cross-page selection as one unit without changing its aspect."""
    if (not isinstance(region, list) or len(region) != 4
            or any(type(value) not in (int, float) or not isfinite(value) for value in region)):
        raise ValueError('native_components.box must be [x, y, w, h] in cm')
    x, y, w, h = (round(v * T.CM) for v in region)
    if min(x, y) < 0 or min(w, h) <= 0 or x + w > canvas[0] or y + h > canvas[1]:
        raise ValueError('native_components.box is outside the slide')
    if not shapes:
        raise ValueError('native component has no matching transferable shapes')
    left = min(shape.left for shape in shapes)
    top = min(shape.top for shape in shapes)
    width = max(shape.left + shape.width for shape in shapes) - left
    height = max(shape.top + shape.height for shape in shapes) - top
    if not width or not height:
        raise ValueError('selected shapes need nonzero bounds to fit a box')
    scale = min(w / width, h / height)
    origin_x = x + (w - width * scale) / 2
    origin_y = y + (h - height * scale) / 2
    for shape in shapes:
        shape.left = round(origin_x + (shape.left - left) * scale)
        shape.top = round(origin_y + (shape.top - top) * scale)
        shape.width = round(shape.width * scale)
        shape.height = round(shape.height * scale)
        for el in shape._element.iter():
            if el.tag in (qn('a:rPr'), qn('a:defRPr'), qn('a:endParaRPr')) and el.get('sz'):
                el.set('sz', str(max(100, min(400000, round(int(el.get('sz')) * scale)))))


def add_native_components(slide, page, library, canvas, manifest, record):
    selections = page.get('native_components') or []
    if not isinstance(selections, list):
        raise ValueError('native_components must be a list')
    for item in selections:
        if not isinstance(item, dict) or not item.get('include'):
            raise ValueError('each native component requires a nonempty include list')
        spec = dict(item)
        region = spec.pop('box', None)
        order = spec.pop('audit', {})
        before = len(slide.shapes)
        automatic = (record.get("selection") or {}).get("status") == "selected_component"
        remap = apply_native_reference(library, slide, spec, automatic=automatic)
        native_records(slide, order, manifest, record, spec, remap, page.get('source'))
        if region is not None:
            position_native_selection(list(slide.shapes)[before:], region, canvas)


def prepare_component_page(page):
    """Resolve the authored information relationship before allocating boxes."""
    if 'component' not in page:
        return page
    if not isinstance(page['component'], dict):
        raise ValueError('component must be an object')
    page = deepcopy(page)
    spec = page['component']
    if 'intent' in spec:
        if 'layout' in spec:
            raise ValueError('choose component.intent or component.layout')
        intent = spec.pop('intent')
        if not isinstance(intent, str) or intent not in COMPONENT_FOR_INTENT:
            raise ValueError('unknown component.intent; choose ' + ', '.join(COMPONENT_FOR_INTENT))
        spec['layout'] = COMPONENT_FOR_INTENT[intent]
    layout = spec.get('layout')
    if isinstance(layout, str) and layout in PRIMARY_FIGURE_COMPONENTS:
        primary_figure_layout.layout_figures(page, spec, fig_path)
    return page


def add_component(slide, page):
    """Add one native component inside the open body of a bullets slide."""
    if page['type'] != 'bullets':
        raise ValueError('component requires a bullets page')
    if len(page.get('bullets') or []) > 2:
        raise ValueError('component pages support at most two bullets')
    for key in ('summary', 'source', 'notes'):
        need(page, key)
    spec = page['component']
    reusable_components.apply(slide, spec)
    # The component renderer validates its box. Check figure boxes only on these pages.
    left, top, width, height = spec.get('box', reusable_components.BODY_BOX)
    body_x, body_y, body_w, body_h = reusable_components.BODY_BOX
    for figure in page.get('figures') or []:
        box = figure.get('box')
        if (not isinstance(box, list) or len(box) != 4
                or any(type(value) not in (int, float) or not isfinite(value) for value in box)):
            raise ValueError('a figure beside a component needs a numeric [x, y, w, h] box')
        x, y, w, h = box
        if (w <= 0 or h <= 0 or x < body_x or y < body_y
                or x + w > body_x + body_w + 1e-8 or y + h > body_y + body_h + 1e-8):
            raise ValueError('figure beside a component must stay inside the content area')
        overlap_width = min(x + w, left + width) - max(x, left)
        overlap_height = min(y + h, top + height) - max(y, top)
        if overlap_width > 1e-8 and overlap_height > 1e-8:
            raise ValueError('figure and component boxes overlap')


def place_card_figures(slide, cards, manifest, record, source):
    boxes = T.card_boxes(slide)
    if len(cards) != len(boxes):
        raise ValueError('%d cards in the plan but the archetype has %d'
                         % (len(cards), len(boxes)))
    for card, (x, y, w, h) in zip(cards, boxes):
        if card.get('foot'):
            h -= 1.55
        fig = card.get('figure')
        if fig:
            picture = T.place_picture(slide, fig_path(fig), x, y, w, h,
                                      name='正文图_%d' % (len(slide.shapes)))
            manifest.evidence(record, picture.shape_id, fig, source)


def fill_optional_shape(shape, value):
    if shape is None:
        return
    if value is None or not str(value).strip():
        T.drop(shape)
        return
    T.set_lines(shape.text_frame, [value])


def build_triple(slide, page, manifest, record):
    cards = need(page, 'cards')
    titles = sorted(T.shapes_named(slide, '卡片标题'), key=lambda sh: sh.left)
    if not 1 <= len(cards) <= len(titles):
        raise ValueError('triple page needs one to %d cards' % len(titles))
    backgrounds = [
        sorted(T.shapes_named(slide, name), key=lambda sh: sh.left)
        for name in ('卡片底板', '卡片标题栏')
    ]
    for i, title in enumerate(titles, 1):
        caption = T.shape(slide, '图注%d' % i, required=False)
        top = caption.top / T.CM + 0.75 if caption is not None else 9.8
        fields = (
            ('value', T.shape(slide, '卡片数值%d' % i, required=False)),
            ('note', T.shape(slide, '卡片说明%d' % i, required=False)),
            ('caption', caption),
        )
        if i > len(cards):
            for column in backgrounds:
                if i <= len(column):
                    T.drop(column[i - 1])
            T.drop(title)
            for _, shape in fields:
                if shape is not None:
                    T.drop(shape)
            continue

        card = cards[i - 1]
        T.set_lines(title.text_frame, [need(card, 'title')])
        for key, shape in fields:
            fill_optional_shape(shape, card.get(key))
        if card.get('figure'):
            x = title.left / T.CM - 0.2
            w = title.width / T.CM + 1.0
            picture = T.place_picture(slide, fig_path(card['figure']), x, top, w, 16.15 - top,
                                      name='正文图%d' % i)
            manifest.evidence(record, picture.shape_id, card['figure'], page.get('source'))

    strip = page.get('strip')
    fill_optional_shape(T.shape(slide, '前置标签', required=False), strip)
    if strip is None or not str(strip).strip():
        background = T.shape(slide, '前置标签底', required=False)
        if background is not None:
            T.drop(background)


def build(plan, template, output, native_library=DEFAULT_NATIVE_LIBRARY, *,
          audit_manifest=None, layout_options=0):
    if type(layout_options) is not int or layout_options not in {0, 2, 3}:
        raise ValueError("layout_options 只能为 0、2 或 3")
    prs = Presentation(template)
    meta = plan.get('meta') or {}
    clean_template(prs, str(meta.get('title_cn') or '学术文献汇报').strip() or '学术文献汇报')
    sections = plan.get('sections') or ['研究背景', '设计与方法', '研究结果', '结论与展望']
    if len(sections) != 4:
        raise ValueError('the agenda archetype needs exactly four sections')

    manifest = ReportManifest(plan.get('audit'), fig_path)
    prefer_native = plan.get("prefer_native", False)
    if not isinstance(prefer_native, bool):
        raise ValueError("prefer_native 必须是 true 或 false")
    library = None
    for n, page in enumerate(plan['slides']):
        record = manifest.new_slide(n + 1)
        decision = None
        try:
            if "evidence_cards" in page:
                if (page.keys() & (FORM_FIELDS | {"native", "native_profile", "native_components",
                                                 "component", "cards", "strip"})
                        or ("native_auto" in page and page["native_auto"] is not False)):
                    raise ValueError("evidence_cards 自按图文内容选版，请移除其它版式或结构字段")
                page, decision = primary_figure_layout.prepare_evidence_page(page, fig_path)
                record['selection'] = decision
            else:
                if page.get("native_auto") is False:
                    page = {key: value for key, value in page.items() if key != "native_auto"}
                elif prefer_native and "native_auto" not in page:
                    if inferred := infer_intent(page):
                        page = {**page, "native_auto": inferred}
                if "native_auto" in page:
                    page, decision = choose_native_layout(
                        page, Path(BASE), Path(native_library), option_limit=layout_options)
                    record['selection'] = decision
                elif FORM_FIELDS & page.keys():
                    raise ValueError("结构化内容需要指定 native_auto 或改用常规组件字段")
                page = prepare_component_page(page)
        except ValueError as error:
            raise ValueError(f'slide {n + 1}: {error}') from error
        kind = page.get('type')
        if (kind == 'native' or page.get('native_components')) and library is None:
            library = NativeLibrary(native_library)
        if kind == 'native':
            try:
                slide, remap = native_page(prs, page, library, decision)
                order = decision if decision and decision['status'] == 'selected' else page.get('native_audit', {})
                native_records(slide, order, manifest, record, page['native'], remap, page.get('source'))
            except ValueError as error:
                raise ValueError(f'slide {n + 1}: {error}') from error
        elif kind in ARCHETYPES:
            slide = T.duplicate_slide(prs, prs.slides[ARCHETYPES[kind]])
        else:
            raise ValueError('slide %d: unknown type %r' % (n + 1, kind))

        if kind == 'cover':
            T.fill_text(slide, 'T_Title', need(meta, 'title_cn'))
            T.fill_text(slide, 'T_EnglishTitle', need(meta, 'title_en'))
            T.fill_text(slide, 'T_Conference', '文献汇报', required=False)
            footer = meta.get('footer')
            if footer is not None and str(footer).strip():
                T.fill_text(slide, 'T_Footer', footer)
            else:
                shape = T.shape(slide, 'T_Footer', required=False)
                if shape is not None:
                    T.drop(shape)
        elif kind == 'agenda':
            for i, name in enumerate(sections, 1):
                T.fill_text(slide, 'S%d_Name' % i, name)
        elif kind == 'divider':
            T.fill_text(slide, 'D_Title', need(page, 'title'))

        if kind in ('bullets', 'table', 'cards2', 'triple'):
            T.fill_text(slide, '页标题', need(page, 'title'))
            bullets = page.get('bullets') or []
            if bullets:
                if len(bullets) > 3:
                    raise ValueError('slide %d: keep 要点 to at most three lines' % (n + 1))
                T.fill_text(slide, '要点1', bullets)
            else:
                T.drop(T.shape(slide, '要点1'))
            if page.get('summary'):
                T.fill_text(slide, '总结文字', page['summary'])
            else:
                T.drop(T.shape(slide, '总结条'))

        if kind == 'cards2':
            cards = page.get('cards') or []
            titles = sorted(T.shapes_named(slide, '卡片标题'), key=lambda s: s.left)
            if len(cards) != len(titles):
                raise ValueError('slide %d: %d cards but the archetype has %d columns'
                                 % (n + 1, len(cards), len(titles)))
            for sh, card in zip(titles, cards):
                T.set_lines(sh.text_frame, [card['title']])
            feet = sorted(T.shapes_named(slide, '卡片脚注'), key=lambda s: s.left)
            for sh, card in zip(feet, cards):
                if card.get('foot'):
                    T.set_lines(sh.text_frame, [card['foot']])
                else:
                    T.drop(sh)
            place_card_figures(slide, cards, manifest, record, page.get('source'))

        if kind == 'triple':
            build_triple(slide, page, manifest, record)

        if kind == 'table':
            spec = need(page, 'table')
            rows = spec['rows']
            tbl_shapes = [sh for sh in slide.shapes if sh.has_table]
            if not tbl_shapes:
                raise ValueError('table archetype has no table')
            table = tbl_shapes[0].table
            if (len(table.rows), len(table.columns)) != (len(rows), len(rows[0])):
                raise ValueError('slide %d: table is %dx%d in the template, plan has %dx%d'
                                 % (n + 1, len(table.rows), len(table.columns),
                                    len(rows), len(rows[0])))
            T.fill_table(table, rows, size=spec.get('size', 15),
                         widths=spec.get('widths'))

        if 'component' in page:
            try:
                add_component(slide, page)
            except ValueError as error:
                raise ValueError(f'slide {n + 1}: {error}') from error

        if kind in ('bullets', 'cards2', 'triple') and page.get('figures'):
            if "_evidence_panels" in page:
                try:
                    primary_figure_layout.draw_evidence_page(slide, page, record)
                except ValueError as error:
                    raise ValueError(f'slide {n + 1}: {error}') from error
            primary_figure_layout.panel_backgrounds(slide, page)
            for caption in page.get('_native_auto_captions', []):
                shape = slide.shapes.add_textbox(*(T.cm(value) for value in caption['box']))
                shape.name = '原图标题'
                frame = shape.text_frame
                frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
                run = frame.paragraphs[0].add_run()
                run.font.name = 'Microsoft YaHei'
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor.from_string(T.TABLE_HEAD)
                run.text = caption['text']
                record['shapes'].append({'shape_id': shape.shape_id, 'role': 'caption'})
            for i, fig in enumerate(page['figures']):
                picture = T.place_picture(slide, fig_path(fig['path']), *fig['box'], name='补充图%d' % i)
                manifest.evidence(record, picture.shape_id, fig['path'], page.get('source'))

        if page.get('native_components'):
            try:
                add_native_components(slide, page, library, (prs.slide_width, prs.slide_height),
                                      manifest, record)
            except ValueError as error:
                raise ValueError(f'slide {n + 1}: {error}') from error

        try:
            finalize_slide(slide, page, record, manifest.strict, cover=(n == 0))
        except ValueError as error:
            raise ValueError(f'slide {n + 1}: {error}') from error

    for _ in range(len(ARCHETYPES)):
        T.delete_slide(prs, 0)
    problems = list(identity_issues(prs, manifest.data['identity_terms']))
    if problems:
        raise ValueError('身份信息检查失败：\n' + '\n'.join(problems))
    with io.BytesIO() as package:
        prs.save(package)
        problems = list(audit_package(package, manifest.data))
        if problems:
            raise ValueError('成品验收失败：\n' + '\n'.join(problems))
        if hasattr(output, 'write'):
            output.write(package.getbuffer())
        else:
            with open(output, 'xb') as file:
                file.write(package.getbuffer())
    if audit_manifest is not None:
        audit_manifest.update(manifest.data)
    return len(prs.slides)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--plan', required=True)
    ap.add_argument('--output', required=True)
    ap.add_argument('--template', default=DEFAULT_TEMPLATE)
    ap.add_argument('--native-library', default=DEFAULT_NATIVE_LIBRARY)
    ap.add_argument('--manifest', type=Path, help='write the local-only delivery and text-fit checks')
    ap.add_argument('--layout-options', type=int, choices=(2, 3), default=0,
                    help='record up to this many compatible layouts per unpinned page')
    args = ap.parse_args()
    if args.layout_options and args.manifest is None:
        ap.error('--layout-options requires --manifest')
    if os.path.exists(args.output):
        sys.exit('refuse to overwrite %s' % args.output)
    if args.manifest is not None and (args.manifest.exists() or args.manifest.resolve() == Path(args.output).resolve()):
        sys.exit('the manifest must be a new file, separate from the PPTX')
    if args.manifest is not None and not args.manifest.parent.is_dir():
        sys.exit('the manifest directory does not exist')
    plan_bytes = Path(args.plan).read_bytes()
    plan = json.loads(plan_bytes.decode("utf-8"))
    global BASE
    BASE = os.path.dirname(os.path.abspath(args.plan))
    manifest = {}
    count = build(plan, args.template, args.output, args.native_library,
                  audit_manifest=manifest, layout_options=args.layout_options)
    if args.layout_options:
        manifest['plan_sha256'] = sha256(plan_bytes).hexdigest()
    if args.manifest is not None:
        with args.manifest.open('x', encoding='utf-8') as file:
            json.dump(manifest, file, ensure_ascii=False, indent=2)
            file.write('\n')
    for idx, slide in manifest['slides'].items():
        selection = slide.get('selection')
        if selection and selection['status'] == 'selected':
            print(f"  P{idx} 选用原生版式 {selection['source']} P{selection['page']}")
        elif selection and selection['status'] == 'selected_component':
            print(f"  P{idx} 选用补充组件 {selection['source']} P{selection['page']}")
        elif selection:
            print(f"  P{idx} 保留指定的常规页型：{selection['reason']}")
    print('wrote %s (%d slides)' % (args.output, count))
    for name, before, after in T.FIT_LOG:
        print('  自动缩字 %s: %g pt -> %g pt（原文太长，会溢出框）'
              % (name, before, after))


if __name__ == '__main__':
    main()
