"""Editable content components for the blue-and-white academic template."""
from __future__ import annotations

from functools import lru_cache
import math
import re

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

import ppt_template as template


BODY_BOX = (1.24, 5.80, 31.35, 10.45)
NAVY = "173467"
BLUE = template.TABLE_HEAD
PALE = template.TABLE_BAND
INK = template.TABLE_INK
WHITE = "FFFFFF"
LATIN = "Arial"
EAST_ASIAN = "Microsoft YaHei"
PT_PER_CM = 72 / 2.54
WORD = re.compile(r"[\w./%+-]+| +|.", re.ASCII)
CLOSING_PUNCTUATION = frozenset("，。！？、；：）】》〉」』’”％,!?;)")


def _columns(box, count, gap):
    x, y, width, height = box
    column_width = (width - gap * (count - 1)) / count
    return [(x + i * (column_width + gap), y, column_width, height)
            for i in range(count)]


def _rows(box, count, gap):
    x, y, width, height = box
    row_height = (height - gap * (count - 1)) / count
    return [(x, y + i * (row_height + gap), width, row_height)
            for i in range(count)]


@lru_cache(maxsize=8192)
def _advance(text, size, bold):
    return template.line_width_pt(text, LATIN, EAST_ASIAN, bold, size)


def _wrapped_lines(text, width, size, bold):
    """Keep the red runs through manual line breaks in a single editable shape."""
    paragraphs = [[]]
    for fragment, red in template.segments(text):
        for index, part in enumerate(fragment.split("\n")):
            if index:
                paragraphs.append([])
            words = [part] if red and part else WORD.findall(part)
            for word in words:
                styled_word = [(char, red) for char in word]
                if paragraphs[-1] and word[0] in CLOSING_PUNCTUATION:
                    paragraphs[-1][-1].extend(styled_word)
                else:
                    paragraphs[-1].append(styled_word)

    lines = []
    for words in paragraphs:
        line = []
        advance = 0
        for styled_word in words:
            word = "".join(char for char, _ in styled_word)
            if word.isspace():
                if line:
                    line.extend(styled_word)
                    advance += _advance(word, size, bold)
                continue
            width_pt = _advance(word, size, bold)
            if width_pt > width:
                return None
            if line and advance + width_pt > width:
                while line and line[-1][0].isspace():
                    line.pop()
                lines.append(line)
                line = []
                advance = 0
            line.extend(styled_word)
            advance += width_pt
        if line:
            lines.append(line)
    return lines


class _Painter:
    def __init__(self, slide, layout):
        self.shapes = slide.shapes
        self.prefix = f"组件 {layout}"

    def shape(self, kind, box, color, name, edge=None):
        shape = self.shapes.add_shape(kind, *(template.cm(value) for value in box))
        shape.name = f"{self.prefix} {name}"
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(color)
        if edge:
            shape.line.color.rgb = RGBColor.from_string(edge)
            shape.line.width = Pt(0.8)
        else:
            shape.line.fill.background()
        return shape

    def rect(self, box, color, name, edge=None):
        return self.shape(MSO_SHAPE.RECTANGLE, box, color, name, edge)

    def text(self, box, value, name, font, color=INK, bold=False,
             center=False, middle=False):
        if not value:
            return
        ideal, minimum = font
        x, y, width_cm, height_cm = box
        width = width_cm * PT_PER_CM * 0.94
        height = height_cm * PT_PER_CM
        for size in range(ideal, minimum - 1, -1):
            spacing = size * 1.18
            lines = _wrapped_lines(value, width, size, bold)
            if lines and len(lines) * spacing + 3 <= height:
                break
        else:
            raise ValueError(f"{self.prefix} {name}放不下，请缩短文字或扩大对应区域")

        textbox = self.shapes.add_textbox(*(template.cm(value) for value in (x, y, width_cm, height_cm)))
        textbox.name = f"{self.prefix} {name}"
        frame = textbox.text_frame
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.word_wrap = True
        frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE if middle else MSO_ANCHOR.TOP
        for index, line in enumerate(lines):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.alignment = PP_ALIGN.CENTER if center else PP_ALIGN.LEFT
            paragraph.line_spacing = Pt(spacing)
            paragraph.space_before = paragraph.space_after = Pt(0)
            for part, red in _coalesce(line):
                run = paragraph.add_run()
                run.text = part
                run.font.name = LATIN
                run.font.size = Pt(size)
                run.font.bold = bold
                run.font.color.rgb = RGBColor.from_string(template.RED if red else color)
                east_asian = OxmlElement("a:ea")
                east_asian.set("typeface", EAST_ASIAN)
                run._r.get_or_add_rPr().append(east_asian)
        return textbox


def _coalesce(styled_line):
    word, current = [], styled_line[0][1]
    for char, red in styled_line:
        if red != current:
            yield "".join(word), current
            word, current = [], red
        word.append(char)
    yield "".join(word), current


def _card(painter, box, item, label, *, center=False):
    x, y, width, height = box
    pad = 0.42
    title_height = min(1.80, height * 0.42)
    painter.rect(box, WHITE, f"{label} 底板", edge=PALE)
    painter.rect((x, y, width, title_height), PALE, f"{label} 标题底")
    painter.text((x + pad, y, width - 2 * pad, title_height),
                 item["title"], f"{label} 标题", (19, 15), NAVY, True, center, True)
    tag_height = 0.78 if item["tag"] else 0
    painter.text((x + pad, y + title_height + 0.24, width - 2 * pad, tag_height),
                 item["tag"], f"{label} 标签", (14, 12), BLUE, True, center, True)
    painter.text((x + pad, y + title_height + 0.30 + tag_height, width - 2 * pad,
                  height - title_height - tag_height - 0.62),
                 item["body"], f"{label} 正文", (17, 14), INK, False, center)


def _flow(painter, box, items):
    x, y, width, height = box
    cards = _columns((x, y + height * 0.16, width, height * 0.68), len(items), 1.20)
    for i, (card, item) in enumerate(zip(cards, items), 1):
        _card(painter, card, item, f"步骤 {i}", center=True)
        if i < len(cards):
            left, top, wide, high = card
            painter.shape(MSO_SHAPE.RIGHT_ARROW, (left + wide + 0.22, top + high / 2 - 0.23, 0.76, 0.46),
                          BLUE, f"箭头 {i}")


def _compare(painter, box, items):
    x, y, width, height = box
    cards = _columns((x, y + 0.32, width, height - 0.64), 2, 0.72)
    for i, (card, item) in enumerate(zip(cards, items), 1):
        _card(painter, card, item, f"对照 {i}")


def _matrix(painter, box, items):
    rows = _rows(box, 2, 0.50)
    cards = [card for row in rows for card in _columns(row, 2, 0.50)]
    for i, (card, item) in enumerate(zip(cards, items), 1):
        _card(painter, card, item, f"象限 {i}")


def _timeline(painter, box, items):
    x, y, width, height = box
    cols = _columns(box, len(items), 0.38)
    centers = [left + wide / 2 for left, _, wide, _ in cols]
    axis_y = y + max(1.60, height * 0.28)
    painter.rect((centers[0], axis_y - 0.035, centers[-1] - centers[0], 0.07),
                 BLUE, "时间轴")
    for i, ((left, _, wide, _), item, center) in enumerate(zip(cols, items, centers), 1):
        painter.shape(MSO_SHAPE.OVAL, (center - 0.16, axis_y - 0.16, 0.32, 0.32),
                      BLUE, f"节点 {i}")
        painter.text((left, y, wide, axis_y - y - 0.28),
                     item["tag"], f"节点 {i} 时间", (15, 12), BLUE, True, True, True)
        painter.text((left, axis_y + 0.32, wide, 1.56),
                     item["title"], f"节点 {i} 标题", (18, 15), NAVY, True, True, True)
        painter.text((left + 0.10, axis_y + 2.04, wide - 0.20, y + height - axis_y - 2.12),
                     item["body"], f"节点 {i} 正文", (16, 14), INK, False, True)


def _evidence(painter, box, items):
    x, _, width, _ = box
    cue_width = width * 0.32
    for i, ((_, y, _, height), item) in enumerate(zip(_rows(box, len(items), 0.22), items), 1):
        painter.rect((x, y, cue_width, height), PALE, f"证据 {i} 标题底")
        painter.rect((x + cue_width, y, width - cue_width, height), WHITE, f"证据 {i} 内容底", edge=PALE)
        tag_height = 0.70 if item["tag"] else 0
        painter.text((x + 0.38, y + 0.18, cue_width - 0.76, tag_height),
                     item["tag"], f"证据 {i} 标签", (14, 12), BLUE, True, middle=True)
        painter.text((x + 0.38, y + 0.18 + tag_height, cue_width - 0.76, height - tag_height - 0.36),
                     item["title"], f"证据 {i} 标题", (19, 15), NAVY, True, middle=True)
        painter.text((x + cue_width + 0.42, y + 0.18, width - cue_width - 0.84, height - 0.36),
                     item["body"], f"证据 {i} 正文", (17, 14), middle=True)


def _milestones(painter, box, items):
    x, y, width, height = box
    rows = _rows((x + 1.16, y, width - 1.16, height), len(items), 0.25)
    centers = [top + high / 2 for _, top, _, high in rows]
    painter.rect((x + 0.33, centers[0], 0.07, centers[-1] - centers[0]), BLUE, "里程碑轴")
    for i, ((left, top, wide, high), item, center) in enumerate(zip(rows, items, centers), 1):
        painter.shape(MSO_SHAPE.OVAL, (x + 0.20, center - 0.16, 0.32, 0.32), BLUE, f"检查节点 {i}")
        painter.rect((left, top, wide, high), WHITE, f"里程碑 {i} 底板", edge=PALE)
        tag_width = wide * 0.25 if item["tag"] else 0
        title_height = 0.94 if item["body"] else high - 0.22
        painter.text((left + 0.35, top + 0.11, wide - tag_width - 0.70, title_height),
                     item["title"], f"里程碑 {i} 标题", (18, 15), NAVY, True, middle=True)
        painter.text((left + wide - tag_width, top + 0.12, tag_width - 0.28, high - 0.24),
                     item["tag"], f"里程碑 {i} 标签", (15, 12), BLUE, True, True, True)
        painter.text((left + 0.35, top + 1.10, wide - tag_width - 0.70, high - 1.23),
                     item["body"], f"里程碑 {i} 正文", (16, 14))


def _roadmap(painter, box, items):
    for i, ((x, y, width, height), item) in enumerate(zip(_columns(box, len(items), 0.36), items), 1):
        painter.rect((x, y + 1.66, width, height - 1.66), WHITE, f"阶段 {i} 底板", edge=PALE)
        painter.shape(MSO_SHAPE.CHEVRON, (x, y, width, 1.72), PALE, f"阶段 {i} 标题底")
        painter.text((x + 1.0, y, width - 2.0, 1.72),
                     item["title"], f"阶段 {i} 标题", (19, 15), NAVY, True, True, True)
        tag_height = 0.80 if item["tag"] else 0
        painter.text((x + 0.45, y + 2.02, width - 0.90, tag_height),
                     item["tag"], f"阶段 {i} 标签", (15, 12), BLUE, True, middle=True)
        painter.text((x + 0.45, y + 2.11 + tag_height, width - 0.90, height - 2.50 - tag_height),
                     item["body"], f"阶段 {i} 正文", (17, 14))


def _criteria(painter, box, items):
    for i, ((x, y, width, height), item) in enumerate(zip(_rows(box, len(items), 0.22), items), 1):
        painter.rect((x, y, width, height), PALE if i % 2 else WHITE, f"判定 {i} 底板")
        painter.rect((x, y, 0.095, height), BLUE, f"判定 {i} 边线")
        tag_width = width * 0.28 if item["tag"] else 0
        title_height = 0.94 if item["body"] else height - 0.26
        painter.text((x + 0.40, y + 0.13, width - tag_width - 0.82, title_height),
                     item["title"], f"判定 {i} 标题", (18, 15), NAVY, True, middle=True)
        painter.text((x + width - tag_width, y + 0.13, tag_width - 0.36, height - 0.26),
                     item["tag"], f"判定 {i} 阈值", (15, 12), BLUE, True, True, True)
        painter.text((x + 0.40, y + 1.14, width - tag_width - 0.82, height - 1.27),
                     item["body"], f"判定 {i} 正文", (16, 14))


def _blue_header(painter, box, title, label):
    x, y, width, height = box
    painter.rect(box, BLUE, f"{label} 标题底")
    painter.text((x + 0.22, y, width - 0.44, height), title,
                 f"{label} 标题", (17, 14), WHITE, True, True, True)


def draw_evidence_panels(slide, panels, items, *, horizontal, record):
    """Draw GUT-style evidence panels with editable labels and inset measurements."""
    painter = _Painter(slide, "证据图文")

    def rounded(box, color, name, edge=None):
        shape = painter.shape(MSO_SHAPE.ROUNDED_RECTANGLE, box, color, name, edge)
        shape.adjustments[0] = 0.09
        return shape

    def bind(text, container, role, field):
        record["shapes"].append({
            "shape_id": text.shape_id, "role": role, "field": field,
            "container_required": True,
        })
        record["containers"].append({
            "shape_id": text.shape_id, "container_id": container.shape_id, "inset_pt": [2] * 4,
        })

    for box, title, role, field in panels:
        x, y, width, _ = box
        rounded(box, WHITE, f"{role}面板底", PALE)
        header = rounded((x + 0.12, y + 0.12, width - 0.24, 0.96), BLUE, f"{role}标题底")
        text = painter.text((x + 0.40, y + 0.17, width - 0.80, 0.86), title,
                            f"{role}标题", (18, 16), WHITE, True, True, True)
        bind(text, header, role, field)

    x, y, width, height = panels[0][0]
    content = (x + 0.38, y + 1.37, width - 0.76, height - 1.70)
    if horizontal:
        boxes = _columns(content, len(items), 0.36)
    else:
        # Keep short fact cards compact and take up residual space symmetrically.
        total = min(content[3], len(items) * 4.1 + (len(items) - 1) * 0.35)
        content = (content[0], content[1] + (content[3] - total) / 2, content[2], total)
        boxes = _rows(content, len(items), 0.35)
    for index, (box, item) in enumerate(zip(boxes, items, strict=True)):
        x, y, width, height = box
        name = f"条目 {index + 1}"
        inset = rounded(box, PALE, f"{name}底")
        title = painter.text((x + 0.30, y + 0.13, width - 0.60, 0.92),
                             item["title"], f"{name}标题", (17, 16), NAVY, True, middle=True)
        body = painter.text((x + 0.30, y + 1.13, width - 0.60, height - 1.27),
                            item["body"], f"{name}正文", (19, 17), INK, True, middle=True)
        bind(title, inset, "label", f"evidence_cards.items.{index}.title")
        bind(body, inset, "body", f"evidence_cards.items.{index}.body")


def _gut_detail(painter, box, items):
    for i, (row, item) in enumerate(zip(_rows(box, len(items), 0.26), items), 1):
        x, y, width, height = row
        label = f"解读 {i}"
        body = "\n".join(text for text in (item["tag"], item["body"]) if text)
        if not body:
            raise ValueError("gut_detail 每项须给出解读文字")
        header_height = 0.94
        painter.rect(row, PALE, f"{label} 底板", edge=PALE)
        _blue_header(painter, (x, y, width, header_height), item["title"], label)
        painter.text((x + 0.34, y + header_height + 0.14, width - 0.68,
                      height - header_height - 0.28),
                     body, f"{label} 正文", (17, 14), middle=True)


def _gut_metric(painter, box, items):
    x, y, width, height = box
    gap = 0.28
    group_height = min(height, len(items) * 4.3 + (len(items) - 1) * gap)
    group_box = (x, y + (height - group_height) / 2, width, group_height)
    for i, (row, item) in enumerate(zip(_rows(group_box, len(items), gap), items), 1):
        left, top, wide, high = row
        label = f"指标 {i}"
        if not item["tag"]:
            raise ValueError("gut_metric 每项须在 tag 填写指标数值")
        header_height = 0.94
        caption_height = high * 0.32 if item["body"] else 0
        painter.rect(row, PALE, f"{label} 底板", edge=PALE)
        _blue_header(painter, (left, top, wide, header_height), item["title"], label)
        painter.text((left + 0.24, top + header_height, wide - 0.48,
                      high - header_height - caption_height),
                     item["tag"], f"{label} 数值", (29, 18), template.RED, True, True, True)
        painter.text((left + 0.25, top + high - caption_height, wide - 0.50,
                      caption_height - 0.12),
                     item["body"], f"{label} 口径", (15, 13), NAVY, center=True, middle=True)


def _guide_tracks(painter, box, items):
    x, y, width, height = box
    header_height = 1.00
    gap = 0.76
    lane_width = width * 0.18
    stage_box = (x + lane_width + gap, y, width - lane_width - gap, header_height)
    headers = [(x, y, lane_width, header_height), *_columns(stage_box, 3, gap)]
    for header, label in zip(headers, ("研究任务", "输入", "关键工作", "预期产出")):
        _blue_header(painter, header, label, label)

    rows_box = (x, y + header_height + 0.24, width, height - header_height - 0.24)
    for i, ((_, top, _, high), item) in enumerate(zip(_rows(rows_box, len(items), 0.23), items), 1):
        steps = [step.strip() for step in item["body"].splitlines()]
        if len(steps) not in (2, 3) or not all(steps):
            raise ValueError("guide_tracks 的 body 按输入、关键工作、预期产出填写 2–3 行")
        lane_tag = item["tag"]
        if len(steps) == 2:
            steps.append(lane_tag)
            lane_tag = ""
        label = f"任务 {i}"
        painter.rect((x, top, lane_width, high), PALE, f"{label} 底板")
        lane_name = "\n".join(part for part in (item["title"], lane_tag) if part)
        painter.text((x + 0.24, top + 0.14, lane_width - 0.48, high - 0.28),
                     lane_name, f"{label} 名称", (17, 14), NAVY, True, True, True)
        for ((left, _, wide, _), step, stage) in zip(headers[1:], steps, ("输入", "工作", "产出")):
            cell = (left, top, wide, high)
            painter.rect(cell, WHITE, f"{label} {stage}底板", edge=PALE)
            painter.text((left + 0.23, top + 0.13, wide - 0.46, high - 0.26),
                         step, f"{label} {stage}正文", (16, 13), center=True, middle=True)
        for j, ((left, _, wide, _), next_step) in enumerate(zip(headers[1:-1], steps[1:]), 1):
            if next_step:
                painter.shape(MSO_SHAPE.RIGHT_ARROW, (left + wide + 0.13, top + high / 2 - 0.15,
                              gap - 0.26, 0.30), BLUE, f"{label} 连接 {j}")


def _guide_risks(painter, box, items, headers=("预期风险", "应对措施")):
    x, y, width, height = box
    header_height = 1.00
    risk_width = (width - 1.48) * 0.43
    response_x = x + risk_width + 1.48
    response_width = width - (response_x - x)
    _blue_header(painter, (x, y, risk_width, header_height), headers[0], "风险")
    _blue_header(painter, (response_x, y, response_width, header_height), headers[1], "应对")
    rows_box = (x, y + header_height + 0.24, width, height - header_height - 0.24)
    for i, ((_, top, _, high), item) in enumerate(zip(_rows(rows_box, len(items), 0.23), items), 1):
        if not item["body"]:
            raise ValueError("guide_risks 每项须在 body 填写对应的应对措施")
        label = f"风险应对 {i}"
        risk = "\n".join(part for part in (item["title"], item["tag"]) if part)
        painter.rect((x, top, risk_width, high), PALE, f"{label} 风险底板")
        painter.rect((response_x, top, response_width, high), WHITE,
                     f"{label} 应对底板", edge=PALE)
        painter.shape(MSO_SHAPE.RIGHT_ARROW,
                      (x + risk_width + 0.32, top + high / 2 - 0.19, 0.84, 0.38),
                      BLUE, f"{label} 连接")
        painter.text((x + 0.33, top + 0.12, risk_width - 0.66, high - 0.24),
                     risk, f"{label} 风险正文", (17, 14), NAVY, True, middle=True)
        painter.text((response_x + 0.33, top + 0.12, response_width - 0.66, high - 0.24),
                     item["body"], f"{label} 应对正文", (17, 14), middle=True)


_LAYOUTS = {
    "flow": (_flow, 2, 4),
    "compare": (_compare, 2, 2),
    "matrix": (_matrix, 4, 4),
    "timeline": (_timeline, 2, 5),
    "evidence": (_evidence, 2, 3),
    "milestones": (_milestones, 2, 4),
    "roadmap": (_roadmap, 2, 3),
    "criteria": (_criteria, 2, 4),
    "gut_detail": (_gut_detail, 2, 3),
    "gut_metric": (_gut_metric, 1, 2),
    "guide_tracks": (_guide_tracks, 2, 3),
    "guide_risks": (_guide_risks, 2, 3),
}
_FIELDS = {"title", "tag", "body"}
_NARROW_LAYOUT_WIDTHS = {"gut_detail": 8.8, "gut_metric": 5.8}


def _read_box(box, minimum_width=12):
    if not isinstance(box, (list, tuple)) or len(box) != 4:
        raise ValueError("component.box 需为 [x, y, w, h]，单位 cm")
    if any(type(value) not in (int, float) or not math.isfinite(value) for value in box):
        raise ValueError("component.box 仅接受有限坐标")
    x, y, width, height = box
    left, top, max_width, max_height = BODY_BOX
    if (x < left or y < top or width < minimum_width or height < 6
            or x + width > left + max_width + 1e-8 or y + height > top + max_height + 1e-8):
        raise ValueError(f"component.box 应位于正文区 x=1.24–32.59、y=5.80–16.25 cm，至少 {minimum_width:g} × 6 cm")
    return tuple(box)


def _read_item(item, number):
    if not isinstance(item, dict) or item.keys() - _FIELDS or "title" not in item:
        raise ValueError(f"component.items[{number}] 仅接受 title、tag、body，且须有 title")
    values = {}
    for key in ("title", "tag", "body"):
        value = item.get(key, "")
        if not isinstance(value, str):
            raise ValueError(f"component.items[{number}].{key} 应为字符串")
        value = value.strip()
        if key == "title" and not value:
            raise ValueError(f"component.items[{number}].title 不能为空")
        if key != "body" and "\n" in value:
            raise ValueError(f"component.items[{number}].{key} 请填写单段文字")
        if any(ord(char) < 32 and char != "\n" for char in value):
            raise ValueError(f"component.items[{number}].{key} 含控制字符")
        if value.count("**") not in (0, 2) or ("**" in value and not value.split("**")[1].strip()):
            raise ValueError(f"component.items[{number}].{key} 可用一对 **…** 标注重点")
        values[key] = value
    return values


def apply(slide, spec):
    """Add one native component to the content region of a template slide."""
    if not isinstance(spec, dict) or spec.keys() - {"layout", "items", "box", "headers"}:
        raise ValueError("component 仅接受 layout、items、box 和双栏 headers")
    layout = spec.get("layout")
    if not isinstance(layout, str) or layout not in _LAYOUTS:
        raise ValueError("component.layout 可选 " + ", ".join(_LAYOUTS))
    draw, minimum, maximum = _LAYOUTS[layout]
    items = spec.get("items")
    if not isinstance(items, list) or not minimum <= len(items) <= maximum:
        expected = str(minimum) if minimum == maximum else f"{minimum}–{maximum}"
        raise ValueError(f"component.layout {layout} 需要 {expected} 项")
    box = _read_box(spec.get("box", BODY_BOX), _NARROW_LAYOUT_WIDTHS.get(layout, 12))
    options = {}
    if "headers" in spec:
        headers = spec["headers"]
        if (layout != "guide_risks" or not isinstance(headers, list) or len(headers) != 2
                or any(not isinstance(value, str) or not value.strip() or "\n" in value for value in headers)):
            raise ValueError("双栏 headers 需要两个非空单行标题")
        options["headers"] = headers
    draw(_Painter(slide, layout), box, [_read_item(item, i) for i, item in enumerate(items)], **options)
