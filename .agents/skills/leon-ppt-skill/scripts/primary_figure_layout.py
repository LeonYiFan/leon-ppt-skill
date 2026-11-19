"""Place original figures alongside the native GUT-inspired components."""

from copy import deepcopy
from itertools import combinations
from math import isfinite

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt

import ppt_template as template
from reusable_components import BODY_BOX, draw_evidence_panels


_FIGURE_COUNTS = {"gut_detail": 1, "gut_metric": 2}
_GAP = 0.50
_INSET = 0.22
_EPS = 1e-8
_EVIDENCE_BULLET_SIZE = 20.5
_PT_PER_CM = 72 / 2.54


def _card_text(value, field, max_lines=1):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} 需要非空文字")
    lines = value.strip().splitlines()
    if len(lines) > max_lines or any(not line.strip() for line in lines):
        raise ValueError(f"{field} 最多 {max_lines} 行")
    for line in lines:
        if (any(ord(char) < 32 for char in line) or line.count("**") not in (0, 2)
                or ("**" in line and not line.split("**")[1].strip())):
            raise ValueError(f"{field} 每行可用一对 **…** 标重点")
    return "\n".join(lines)


def _evidence_panel_boxes(top, aspect, horizontal):
    x, y, width, height = BODY_BOX
    available = y + height - top
    if horizontal:
        cards_height = 4.45
        cards = (x, top, width, cards_height)
        figure = (x, top + cards_height + _GAP, width, available - cards_height - _GAP)
    else:
        figure_width = max(12.0, min(19.0, aspect * (available - 1.55) + 0.7))
        cards_width = min(15.3, width - figure_width - _GAP)
        left = x + (width - cards_width - figure_width - _GAP) / 2
        cards = (left, top, cards_width, available)
        figure = (left + cards_width + _GAP, top, figure_width, available)
    fx, fy, fw, fh = figure
    picture = [fx + 0.34, fy + 1.26, fw - 0.68, fh - 1.59]
    return cards, figure, picture


def prepare_evidence_page(page, fig_path):
    """Select a compact evidence arrangement from the authored facts and figure."""
    if page.get("type") != "bullets":
        raise ValueError("evidence_cards 只用于 bullets 页")
    page = deepcopy(page)
    spec = page["evidence_cards"]
    if not isinstance(spec, dict) or spec.keys() != {"title", "items"}:
        raise ValueError("evidence_cards 需要 title 和 items")
    group_title = _card_text(spec["title"], "evidence_cards.title")
    items = spec["items"]
    if not isinstance(items, list) or not 2 <= len(items) <= 3:
        raise ValueError("evidence_cards.items 需要 2–3 组并列证据")
    for index, item in enumerate(items):
        if not isinstance(item, dict) or item.keys() != {"title", "body"}:
            raise ValueError("evidence_cards.items 每项需要 title 和 body")
        item["title"] = _card_text(item["title"], f"evidence_cards.items[{index}].title")
        item["body"] = _card_text(item["body"], f"evidence_cards.items[{index}].body", 2)
    for field in ("summary", "source", "notes"):
        if not isinstance(page.get(field), str) or not page[field].strip():
            raise ValueError(f"evidence_cards 页需要 {field}")

    figures = page.get("figures")
    if (not isinstance(figures, list) or len(figures) != 1 or not isinstance(figures[0], dict)
            or not isinstance(figures[0].get("path"), str) or not figures[0]["path"].strip()):
        raise ValueError("evidence_cards 需要一张原图，填写 figures[0].path")
    figure = figures[0]
    if "box" in figure:
        raise ValueError("evidence_cards 会自动放图，请移除 figures[0].box")
    caption = _card_text(figure.get("caption", "论文原图"), "figures[0].caption")
    aspect = _aspect(figure, fig_path)
    minima = [figure.get(f"min_{axis}_cm", default) for axis, default in (("width", 5.5), ("height", 3.5))]
    if any(type(limit) not in (int, float) or not isfinite(limit) or limit <= 0 for limit in minima):
        raise ValueError("原图最小宽高须为大于零的有限厘米值")

    bullets = page.get("bullets") or []
    if not isinstance(bullets, list) or len(bullets) > 2:
        raise ValueError("evidence_cards 页允许至多两条单行要点")
    for index, bullet in enumerate(bullets):
        line = _card_text(bullet, f"bullets[{index}]")
        if template.line_width_pt(line, "Arial", "Microsoft YaHei", True, _EVIDENCE_BULLET_SIZE) > 30 * _PT_PER_CM:
            raise ValueError("请缩短证据图文页分点，保持单行可读")
    bullet_height = 0.33 + len(bullets)
    top = max(3.60, 2.58 + bullet_height + 0.15)
    prefer_horizontal = aspect >= 2.55
    for horizontal in (prefer_horizontal, not prefer_horizontal):
        cards_box, figure_box, picture = _evidence_panel_boxes(top, aspect, horizontal)
        fit_width, fit_height = min(picture[2], picture[3] * aspect), min(picture[3], picture[2] / aspect)
        if fit_width >= minima[0] and fit_height >= minima[1] and fit_width * fit_height >= 28:
            break
    else:
        raise ValueError("证据卡与当前原图同页时图幅过小，请将图独立成页")

    figure["box"] = picture
    page["_evidence_panels"] = (
        (cards_box, group_title, "label", "evidence_cards.title"),
        (figure_box, caption, "caption", "figures.0.caption"),
    )
    page["_evidence_horizontal"] = horizontal
    page["_evidence_bullet_height"] = bullet_height
    decision = {
        "status": "selected", "intent": "evidence_cards", "source": "gut", "page": 52,
        "profile": "gut-evidence-cards", "variant": "wide-figure-below" if horizontal else "cards-beside-figure",
        "components": [
            {"source": "gut", "page": 52, "role": "paired-panel-style"},
            {"source": "gut", "page": 41, "role": "pale-blue-measurement-card-style"},
        ],
        "figures_cm": [[round(fit_width, 3), round(fit_height, 3)]],
    }
    return page, decision


def draw_evidence_page(slide, page, record):
    """Render the selected evidence panels into the filled template canvas."""
    bullet = template.shape(slide, "要点1", required=False)
    if bullet is not None:
        bullet.height = template.cm(page["_evidence_bullet_height"])
        frame = bullet.text_frame
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.word_wrap = False
        for paragraph in frame.paragraphs:
            paragraph.line_spacing = Pt(_EVIDENCE_BULLET_SIZE * 1.34)
            for run in paragraph.runs:
                run.font.size = Pt(_EVIDENCE_BULLET_SIZE)
                run.font.bold = True
    draw_evidence_panels(slide, page["_evidence_panels"], page["evidence_cards"]["items"],
                         horizontal=page["_evidence_horizontal"], record=record)


def _inside_body(box):
    if (not isinstance(box, (list, tuple)) or len(box) != 4
            or any(type(value) not in (int, float) or not isfinite(value) for value in box)):
        raise ValueError("主参考图文版式的 box 应为四个有限坐标 [x, y, w, h]，单位 cm")
    x, y, w, h = box
    body_x, body_y, body_w, body_h = BODY_BOX
    if (w <= 0 or h <= 0 or x < body_x or y < body_y
            or x + w > body_x + body_w + _EPS or y + h > body_y + body_h + _EPS):
        raise ValueError("原图和解读组件须位于正文区")


def _overlap(a, b):
    x1, y1, w1, h1 = a
    x2, y2, w2, h2 = b
    return (min(x1 + w1, x2 + w2) - max(x1, x2) > _EPS
            and min(y1 + h1, y2 + h2) - max(y1, y2) > _EPS)


def _aspect(figure, fig_path):
    try:
        width, height = template._img_size(fig_path(figure["path"]))
    except OSError as error:
        raise ValueError("主参考图文版式无法读取原图尺寸，请核对图片路径和文件格式") from error
    return width / height


def _detail(aspect):
    x, y, width, height = BODY_BOX
    # Keep close to 2:1 for typical landscape panels. Give back space to the
    # reading notes when the figure would hit the height limit first.
    desired = aspect * (height - 2 * _INSET) + 2 * _INSET
    figure_width = max(18.5, min(width - _GAP - 10.2, desired))
    notes = [x + figure_width + _GAP, y, width - _GAP - figure_width, height]
    return [[x, y, figure_width, height]], notes


def _metric(aspects):
    x, y, width, height = BODY_BOX
    metric_width = 6.5
    available = width - metric_width - 2 * _GAP
    # Balance the two panels around the native metric card, while giving a
    # wider original extra horizontal space without crowding its neighbour.
    even = available / 2
    target = available * aspects[0] / sum(aspects)
    left_width = max(even - 1.6, min(even + 1.6, target))
    metric_x = x + left_width + _GAP
    frames = [[x, y, left_width, height],
              [metric_x + metric_width + _GAP, y, available - left_width, height]]
    return frames, [metric_x, y, metric_width, height]


def layout_figures(page, spec, fig_path):
    """Validate and assign body geometry on the builder's copied plan."""
    layout = spec.get("layout")
    if not isinstance(layout, str) or layout not in _FIGURE_COUNTS:
        return
    count = _FIGURE_COUNTS[layout]
    figures = page.get("figures")
    if (not isinstance(figures, list) or len(figures) != count
            or any(not isinstance(figure, dict) or not isinstance(figure.get("path"), str)
                   or not figure["path"].strip() for figure in figures)):
        raise ValueError(f"{layout} 需要 {count} 张原图，每张图填写 path")

    overrides = ["box" in spec, *("box" in figure for figure in figures)]
    if any(overrides):
        if not all(overrides):
            raise ValueError("自定义位置时请同时填写组件和每张原图的 box")
        frames = [figure["box"] for figure in figures]
        boxes = [spec["box"], *frames]
        for box in boxes:
            _inside_body(box)
        if any(_overlap(a, b) for a, b in combinations(boxes, 2)):
            raise ValueError("原图之间及原图与解读组件之间不能重叠")
    else:
        aspects = [_aspect(figure, fig_path) for figure in figures]
        frames, spec["box"] = _detail(aspects[0]) if count == 1 else _metric(aspects)
        for figure, (x, y, w, h) in zip(figures, frames):
            figure["box"] = [x + _INSET, y + _INSET, w - 2 * _INSET, h - 2 * _INSET]
    page["_primary_figure_frames"] = frames


def panel_backgrounds(slide, page):
    """Insert editable, quiet panel frames before the original pictures."""
    for index, frame in enumerate(page.get("_primary_figure_frames", ()), 1):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *(template.cm(value) for value in frame))
        shape.name = f"原图面板 {index}"
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
        shape.line.color.rgb = RGBColor.from_string(template.TABLE_BAND)
        shape.line.width = Pt(0.8)
