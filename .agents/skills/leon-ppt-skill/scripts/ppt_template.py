# -*- coding: utf-8 -*-
"""Helpers for building a deck from the bundled blue-and-white template deck.

The template deck (``assets/template-deck.pptx``) keeps the real page furniture
of the user's own PPT — title bar, 要点块, 灰底总结条, 卡片, 章节页照片, 封面与
致谢 — with every paper-specific text replaced by a placeholder.  Pages are
produced by *duplicating an archetype page and filling its named shapes*, so a
new deck inherits the template instead of re-drawing it.

Conventions used by this module and by ``build_report.py``:

* Text markers: ``**文字**`` renders that fragment in the template red (E91824).
* Geometry is expressed in centimetres (the deck is 33.87 × 19.05 cm, 16:9).
* Shapes are addressed by name; a missing name is a hard error, never a silent skip.
"""
from __future__ import annotations

import copy
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

CM = 360000
SLIDE_W = 33.867
SLIDE_H = 19.05
RED = 'E91824'
DECK_SIZE = (12192000, 6858000)

R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

_FILL_TAGS = ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
              'a:pattFill', 'a:grpFill')


def cm(value):
    """Centimetres -> EMU."""
    return Emu(int(round(value * CM)))


# --------------------------------------------------------------------- shapes
def iter_shapes(shapes, include_groups=False):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            if include_groups:
                yield sh
            yield from iter_shapes(sh.shapes, include_groups=include_groups)
        else:
            yield sh


def shape(slide, name, required=True):
    for sh in iter_shapes(slide.shapes, include_groups=True):
        if sh.name == name:
            return sh
    if required:
        raise KeyError('shape %r not found (available: %s)'
                       % (name, ', '.join(sh.name for sh in iter_shapes(slide.shapes))))
    return None


def shapes_named(slide, name):
    return [sh for sh in iter_shapes(slide.shapes) if sh.name == name]


def drop(sh):
    el = sh._element
    el.getparent().remove(el)


# -------------------------------------------------------------------- slides
def delete_slide(prs, index):
    lst = prs.slides._sldIdLst
    item = list(lst)[index]
    prs.part.drop_rel(item.get(qn('r:id')))
    lst.remove(item)


def move_slide(prs, old_index, new_index):
    lst = prs.slides._sldIdLst
    item = list(lst)[old_index]
    lst.remove(item)
    lst.insert(new_index, item)


def duplicate_slide(prs, src):
    """Append a deep copy of *src* (same layout, pictures and background)."""
    dst = prs.slides.add_slide(src.slide_layout)
    for sh in list(dst.shapes):           # drop layout placeholders
        drop(sh)
    for el in src.shapes._spTree:
        if el.tag in (qn('p:nvGrpSpPr'), qn('p:grpSpPr')):
            continue
        dst.shapes._spTree.append(copy.deepcopy(el))
    src_csld = src._element.find(qn('p:cSld'))
    bg = src_csld.find(qn('p:bg'))
    if bg is not None:
        csld = dst._element.find(qn('p:cSld'))
        old = csld.find(qn('p:bg'))
        if old is not None:
            csld.remove(old)
        csld.insert(0, copy.deepcopy(bg))
    for el in dst._element.cSld.iter():
        for attr, val in list(el.attrib.items()):
            if not attr.startswith(R) or val not in src.part.rels:
                continue
            rel = src.part.rels[val]
            if rel.is_external:
                el.set(attr, dst.part.relate_to(rel.target_ref, rel.reltype, is_external=True))
            else:
                el.set(attr, dst.part.relate_to(rel.target_part, rel.reltype))
    return dst


# ---------------------------------------------------------------------- text
def _set_fill(rpr, color):
    for tag in _FILL_TAGS:
        for el in rpr.findall(qn(tag)):
            rpr.remove(el)
    fill = rpr.makeelement(qn('a:solidFill'), {})
    clr = rpr.makeelement(qn('a:srgbClr'), {'val': color})
    fill.append(clr)
    ln = rpr.find(qn('a:ln'))
    index = list(rpr).index(ln) + 1 if ln is not None else 0
    rpr.insert(index, fill)


def _donor_rpr(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            rpr = r._r.find(qn('a:rPr'))
            if rpr is not None:
                return copy.deepcopy(rpr)
    return None


def _make_run(text, rpr, color=None):
    if rpr is None:
        raise ValueError('placeholder paragraph has no run properties')
    run = rpr.makeelement(qn('a:r'), {})
    pr = copy.deepcopy(rpr)
    if color:
        _set_fill(pr, color)
    run.append(pr)
    t = run.makeelement(qn('a:t'), {})
    t.text = text
    if text != text.strip():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    run.append(t)
    return run


def segments(line):
    """'结论：**0.917** 很稳' -> [('结论：', False), ('0.917', True), (' 很稳', False)]"""
    out = []
    for i, part in enumerate(str(line).split('**')):
        if part:
            out.append((part, i % 2 == 1))
    return out


def text_em(line):
    """Rough width of *line* in em: CJK counts 1, ASCII about half."""
    line = str(line).replace('**', '')
    em = 0.0
    for ch in line:
        if ch.isascii():
            em += 0.28 if ch in ' il.,:;|!\'"()[]' else 0.56
        else:
            em += 1.0
    return em


def _fonts(rpr):
    """(latin, ea, bold) of the placeholder run, from its own rPr."""
    def face(tag):
        el = rpr.find(qn(tag))
        return (el.get('typeface') or '') if el is not None else ''
    return face('a:latin'), face('a:ea'), rpr.get('b') == '1'


_FONT_FILES = {
    ('msyh', True): 'msyhbd.ttc', ('msyh', False): 'msyh.ttc',
    ('arial', True): 'arialbd.ttf', ('arial', False): 'arial.ttf',
    ('times', True): 'timesbd.ttf', ('times', False): 'times.ttf',
}
_FONT_CACHE = {}


def _font_file(family, bold):
    key = family.strip().lower()
    kind = ('msyh' if ('yahei' in key or '雅黑' in key) else
            'arial' if 'arial' in key else
            'times' if 'times' in key else None)
    if kind is None:
        return None
    name = _FONT_FILES[(kind, bold)]
    path = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', name)
    return path if os.path.exists(path) else None


def line_width_pt(line, latin, ea, bold, size):
    """Exact advance width in points, measured with the real font files.

    Falls back to the em estimate when a font cannot be found, so the call
    still works on a machine without the template fonts."""
    try:
        from PIL import ImageFont
    except ImportError:
        return text_em(line) * size
    total = 0.0
    for ch in str(line).replace('**', ''):
        family = latin if ch.isascii() else ea
        path = _font_file(family, bold)
        if path is None:
            return text_em(line) * size
        key = (path, int(round(size * 10)))
        if key not in _FONT_CACHE:
            _FONT_CACHE[key] = ImageFont.truetype(path, key[1])
        total += _FONT_CACHE[key].getlength(ch) / 10.0
    return total


def _set_sz(rpr, pt):
    if rpr is not None:
        rpr.set('sz', str(int(round(pt * 100))))


FIT_LOG = []          # (shape name, original pt, fitted pt), one build at a time


def set_lines(text_frame, lines, red=RED, base=None, width_cm=None, name='',
              min_scale=0.70):
    """Replace every paragraph, keeping the placeholder's paragraph/run style.

    ``base`` forces the colour of the unmarked text (use it when the
    placeholder happens to start with a red keyword); ``None`` keeps whatever
    colour the placeholder run already has.

    ``width_cm`` turns on horizontal auto-fit: when the longest line would be
    wider than the box at the placeholder's own size, all runs shrink together
    (never below ``min_scale``) instead of spilling out of the slide.  Each
    change lands in ``FIT_LOG`` so the builder can report it."""
    if isinstance(lines, str):
        lines = [lines]
    expanded = []
    for ln in lines:
        for part in str(ln).replace('\x0b', '\n').split('\n'):
            if part.strip():
                expanded.append(part)
    lines = expanded
    if not lines:
        raise ValueError('empty text block')
    txbody = text_frame._txBody
    paras = txbody.findall(qn('a:p'))
    donor_p = copy.deepcopy(paras[0])
    rpr = _donor_rpr(text_frame)
    for p in paras:
        txbody.remove(p)
    size = None
    if rpr is not None and rpr.get('sz'):
        size = int(rpr.get('sz')) / 100.0
    if width_cm and size:
        latin, ea, bold = _fonts(rpr)
        need = max(line_width_pt(ln, latin, ea, bold, size) for ln in lines) / 28.3465
        if need > width_cm * 0.99:
            fitted = max(size * width_cm * 0.92 / need, size * min_scale)
            fitted = max(8.0, int(fitted * 2) / 2.0)      # round down, never up
            if fitted < size:
                FIT_LOG.append((name, size, fitted))
                size = fitted
                _set_sz(rpr, size)
    for line in lines:
        p = copy.deepcopy(donor_p)
        for tag in ('a:r', 'a:br', 'a:fld'):
            for el in p.findall(qn(tag)):
                p.remove(el)
        end = p.find(qn('a:endParaRPr'))       # must stay the last child of a:p
        if end is not None and size:
            _set_sz(end, size)
        for text, is_red in segments(line):
            run = _make_run(text, rpr, red if is_red else base)
            if end is None:
                p.append(run)
            else:
                p.insert(list(p).index(end), run)
        txbody.append(p)


def fill_text(slide, name, lines, red=RED, required=True, base=None,
              width_cm=None):
    """Fill the named shape; returns it, or None when optional and absent."""
    sh = shape(slide, name, required=required)
    if sh is None:
        return None
    if width_cm is None:
        tf = sh.text_frame
        width_cm = (sh.width - (tf.margin_left or 0) - (tf.margin_right or 0)) / CM
    set_lines(sh.text_frame, lines, red, base, width_cm=width_cm, name=name)
    return sh


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = str(text).strip()


# ------------------------------------------------------------------- pictures
def _img_size(path):
    from PIL import Image
    with Image.open(path) as im:
        return im.size


def place_picture(slide, path, left, top, width, height, name=None):
    """Insert *path* fitted (contain) inside the given cm box, centred."""
    iw, ih = _img_size(path)
    k = min(width / iw, height / ih)
    w, h = iw * k, ih * k
    pic = slide.shapes.add_picture(path, cm(left + (width - w) / 2),
                                  cm(top + (height - h) / 2), cm(w), cm(h))
    if name:
        pic.name = name
    # python-pptx sets the image source filename as description by default.
    properties = pic._element.nvPicPr.cNvPr
    properties.attrib.pop('descr', None)
    return pic


def card_boxes(slide, text_sizes=None, pad=0.30, top_pad=1.45):
    """Interior rectangles (cm) of the template cards, ordered as on the slide."""
    boxes = []
    for sh in shapes_named(slide, '卡片底板'):
        boxes.append((sh.left / CM, sh.top / CM, sh.width / CM, sh.height / CM))
    boxes.sort(key=lambda b: (round(b[1], 1), b[0]))
    return [(x + pad, y + top_pad, w - 2 * pad, h - top_pad - pad) for x, y, w, h in boxes]


# ---------------------------------------------------------------------- table
TABLE_HEAD = '2F5597'
TABLE_BAND = 'EAF1FA'
TABLE_INK = '111820'


def style_table(table, size=14.0):
    """Blue white-on-blue header + light banding, matching the template cards."""
    for r, row in enumerate(table.rows):
        for cell in row.cells:
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = RGBColor.from_string(TABLE_HEAD)
            else:
                cell.fill.fore_color.rgb = RGBColor.from_string(
                    TABLE_BAND if r % 2 == 1 else 'FFFFFF')
            cell.vertical_anchor = 3  # middle
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)
                    run.font.bold = r == 0
                    run.font.color.rgb = RGBColor.from_string('FFFFFF' if r == 0 else TABLE_INK)


def fill_table(table, rows, size=14.0, widths=None):
    """Write *rows* (list of lists, first is the header) into a table."""
    if len(rows) != len(table.rows) or any(len(r) != len(table.columns) for r in rows):
        raise ValueError('rows do not match the table template (%d x %d)'
                         % (len(table.rows), len(table.columns)))
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            table.cell(r, c).text = str(value)
    if widths:
        for c, w in enumerate(widths):
            table.columns[c].width = cm(w)
    style_table(table, size)
