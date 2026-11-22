"""Choose a profiled native slide without discarding current-report content."""

from copy import deepcopy
from hashlib import sha256
from math import isfinite
from pathlib import Path
from zipfile import BadZipFile

from PIL import Image

from native_distillation import CandidateMismatch, FORM_FIELDS, load_profiles, slot_values
from native_figure_balance import (
    GUT_LEGEND, METRIC_PANELS, PAIRED_READOUTS, SINGLE_PANEL, STACKED_PICTURE,
    balance_metric_panels, balance_panels, body_geometry, coordinates_in_group,
)
from kw16_component_selection import SUPPLEMENT_INTENTS, choose_kw16_component, iter_kw16_components
from native_reference import A, P, CM, NativeLibrary, _node_map, _paragraph_style
from ppt_template import line_width_pt
from reusable_components import BODY_BOX


PT_PER_CM = 72 / 2.54
PRIMARY_SOURCES = {"gut", "guide"}

def _fallback(page, intent, base_dir):
    page = deepcopy(page)
    page.pop("native_auto")
    figures = page.get("figures")
    if (intent == "figure_evidence" and page.get("type") == "bullets"
            and not page.get("component") and isinstance(figures, list)
            and 1 <= len(figures) <= 4 and all(isinstance(figure, dict) for figure in figures)):
        x, y, width, height = BODY_BOX
        gap = 0.5
        columns = min(2, len(figures))
        rows = (len(figures) + columns - 1) // columns
        panel_w = (width - (columns - 1) * gap) / columns
        panel_h = (height - (rows - 1) * gap) / rows
        automatic_boxes = all("box" not in figure for figure in figures)
        captions = []
        for number, figure in enumerate(figures):
            caption = figure.get("caption", "")
            if not isinstance(caption, str):
                raise ValueError("figures.caption 应为文字")
            caption = caption.strip()
            if automatic_boxes:
                row, column = divmod(number, columns)
                count = min(columns, len(figures) - row * columns)
                offset = (columns - count) * (panel_w + gap) / 2
                start = x + offset + column * (panel_w + gap)
                figure["box"] = [start, y + row * (panel_h + gap), panel_w, panel_h]
            elif "box" not in figure:
                raise ValueError("原生选版回退时每张图须同时提供 box，或全部省略以自动布置")
            box = figure["box"]
            if (not isinstance(box, list) or len(box) != 4
                    or any(type(value) not in (int, float) or not isfinite(value) for value in box)):
                raise ValueError("回退版式的原图框须为有限的 [x, y, w, h] 坐标，单位厘米")
            left, top, fig_width, fig_height = box
            if caption:
                caption_height = 0.8
                captions.append({"text": caption, "box": [left, top, fig_width, caption_height]})
                fig_height -= caption_height
                figure["box"] = [left, top + caption_height, fig_width, fig_height]
            try:
                _figure_footprint(figure, base_dir, fig_width, fig_height)
            except CandidateMismatch as error:
                raise ValueError("当前原图在回退版式中仍过小，请给更大的 box 或拆成独占图页") from error
        if captions:
            page["_native_auto_captions"] = captions
    return page


def _lines(text, count, field):
    if not isinstance(text, str) or not text.strip():
        raise CandidateMismatch(f"{field} 需要非空文字")
    lines = text.strip().splitlines()
    if len(lines) > count or any(not line.strip() for line in lines):
        raise CandidateMismatch(f"{field} 允许 {count} 行")
    return lines


def _dimensions(figure, base_dir):
    path = figure.get("path")
    if not isinstance(path, str) or not path.strip():
        raise ValueError("figure_evidence 的每张原图须有 path")
    try:
        with Image.open(Path(base_dir) / path) as image:
            width, height = image.size
    except (OSError, ValueError) as error:
        raise ValueError("figure_evidence 无法读取原图尺寸") from error
    return width / height


def _minimum(figure, key, default):
    value = figure.get(key, default)
    if type(value) not in (int, float) or not isfinite(value) or value <= 0:
        raise ValueError(f"figures.{key} 应为大于零的有限厘米值")
    return value


def _figure_footprint(figure, base_dir, width, height):
    aspect = _dimensions(figure, base_dir)
    displayed_width, displayed_height = min(width, height * aspect), min(height, width / aspect)
    _validate_footprint(figure, displayed_width, displayed_height)
    return displayed_width, displayed_height


def _validate_footprint(figure, displayed_width, displayed_height):
    minimum_height = 2.25 if displayed_width >= 14 and displayed_width >= 4 * displayed_height else 3.5
    if (displayed_width < _minimum(figure, "min_width_cm", 5.5)
            or displayed_height < _minimum(figure, "min_height_cm", minimum_height)
            or displayed_width * displayed_height < 28):
        raise CandidateMismatch("按原图比例放入后图幅过小")


def _column_geometry(columns, boxes):
    """Use the visually verified Guide slide 30 card gutters and row anchors."""
    geometry = {}
    fixed_text = []
    for column in columns:
        x, _, w, _ = boxes[column["container"]]
        for positions, ids in zip(columns[0]["rows"], column["rows"]):
            badge, heading, body, rule = ids
            left = x + 0.65
            label = left + 1.10
            right = x + w - 0.65
            width, height = boxes[badge][2:]
            geometry[badge] = [left, boxes[positions[0]][1], width, height]
            geometry[heading] = [label, boxes[positions[1]][1], right - label, boxes[heading][3]]
            geometry[body] = [left, boxes[positions[2]][1], right - left, boxes[body][3]]
            geometry[rule] = [left, boxes[positions[3]][1] + 0.10, right - left, boxes[rule][3]]
            fixed_text.extend((heading, body))
    return geometry, fixed_text


def _font(style, name, default):
    item = style.find(A + name)
    return item.get("typeface", default) if item is not None else default


def _measure_text(page, node, box, value, fixed, font_size=None):
    body = node.find(P + "txBody")
    props = body.find(A + "bodyPr")
    paragraphs = body.findall(A + "p")
    fallback = page.theme.implicit_text_style(node)
    preferred = _paragraph_style(paragraphs[0], fallback)
    widest_ratio = 0.0
    required_height = 0.0
    text_height = box[3] * PT_PER_CM
    for index, line in enumerate(value.split("\n")):
        paragraph = paragraphs[min(index, len(paragraphs) - 1)]
        style = _paragraph_style(paragraph, preferred)
        page.theme.materialize(style)
        defaults = paragraph.find(A + "pPr/" + A + "defRPr")
        size = style.get("sz") or (defaults.get("sz") if defaults is not None else None)
        size = font_size if font_size is not None else int(size or "2400") / 100
        latin = _font(style, "latin", "Microsoft YaHei")
        cjk = _font(style, "ea", "Microsoft YaHei")
        bold = style.get("b") in ("1", "true")
        inner = 0 if fixed else sum(int(props.get(key, "91440")) for key in ("lIns", "rIns"))
        para_props = paragraph.find(A + "pPr")
        if para_props is not None:
            inner += sum(max(0, int(para_props.get(key, "0"))) for key in ("marL", "marR", "indent"))
        budget = box[2] * PT_PER_CM - inner / 12700
        if budget <= 0:
            raise CandidateMismatch("文字槽净宽不足")
        width = line_width_pt(line, latin, cjk, bold, size)
        widest_ratio = max(widest_ratio, width / budget)
        required_height += size * 1.2
    vertical_margin = sum(int(props.get(key, "45720")) for key in ("tIns", "bIns")) / 12700
    # COM measures the result after the real font/layout engine has run. This
    # conservative width estimate is for ranking and rejecting poor candidates.
    if widest_ratio > 0.94 or (len(value.splitlines()) > 1
                              and required_height > text_height - vertical_margin):
        raise CandidateMismatch("正文、标题或图注超出原生槽位容量")
    return widest_ratio


def _native_content(page, profile, base_dir, boxes, reflow_figures=False):
    intent = profile["intent"]
    if intent in {"risk_response", "parallel_cards"} and FORM_FIELDS & page.keys():
        raise CandidateMismatch("该版式需要本页对应的常规正文数据字段")
    marked_stage = profile.get("active_stage")
    if marked_stage is not None or "active_stage" in page:
        if type(page.get("active_stage")) is not int or page["active_stage"] != marked_stage:
            raise CandidateMismatch("当前阶段与源页高亮不匹配，请声明合适的 active_stage")
    title = _lines(page.get("title"), 1, "页标题")[0]
    summary = _lines(page.get("summary"), 1, "总结条")[0]
    if not isinstance(page.get("notes"), str) or not page["notes"].strip():
        raise CandidateMismatch("内容页需要讲稿")
    content = {profile["title"]: title}
    roles = {profile["title"]: "title"}
    if summary_id := profile.get("summary"):
        content[summary_id] = summary
        roles[summary_id] = "summary"
    images = {}
    containers = []
    figures_cm = []
    geometry = deepcopy(profile.get("geometry", {})) | {
        slot: [boxes[slot][0] + dx, boxes[slot][1] + dy, *boxes[slot][2:]]
        for slot, (dx, dy) in profile.get("shift_cm", {}).items()
    }
    fixed_text = []

    def put(slot, text, role, container=None, inset=None):
        content[slot] = text
        roles[slot] = role
        if container and text:
            containers.append({"shape_id": slot, "container_id": container,
                               "inset_pt": inset or [2, 2, 2, 2]})

    figures = page.get("figures") or []
    bullets = page.get("bullets") or []
    if not isinstance(figures, list) or not isinstance(bullets, list):
        raise ValueError("native_auto 的 figures 和 bullets 应为列表")
    if "body" in profile:
        if len(bullets) > profile["max_body_lines"]:
            raise CandidateMismatch("超出该版式的要点行数")
        body = "\n".join(_lines(bullet, 1, "要点")[0] for bullet in bullets)
        if body:
            put(profile["body"], body, "body")
    elif bullets:
        raise CandidateMismatch("该原生页未预留额外要点槽")

    if intent == "figure_evidence" or "slots" in profile:
        figure_slots = profile.get("figures", [])
        single_panel = profile["id"] == SINGLE_PANEL
        is_stack = single_panel and len(figures) == 2
        if page.get("type") != "bullets" or (len(figures) != len(figure_slots) and not is_stack):
            raise CandidateMismatch("该版式的图数或正文结构不匹配")
        component = page.get("component")
        if component and (not isinstance(component, dict) or set(component) - {"intent", "items"}):
            raise CandidateMismatch("该版式无法保留额外组件字段")
        if page.get("cards") or page.get("table") or page.get("strip"):
            raise CandidateMismatch("该版式无法保留额外卡片、表格或提示条")
        legend = page.get("figure_legend")
        if legend and (not single_panel or is_stack):
            raise CandidateMismatch("此版式没有单图独立图例槽")
        if is_stack:
            if any(not isinstance(f, dict) or _dimensions(f, base_dir) < 2.3 for f in figures):
                raise CandidateMismatch("上下拼版需要两幅横向原图")
            figure_slots = [figure_slots[0], {"id": STACKED_PICTURE, "label": figure_slots[0]["label"]}]
        caption_fields = {slot.get("field") for slot in profile.get("slots", {}).values()}
        for index, (figure, slot) in enumerate(zip(figures, figure_slots)):
            if not isinstance(figure, dict):
                raise ValueError("figures 每一项应为对象")
            if slot["id"] == STACKED_PICTURE:
                continue
            if not reflow_figures:
                x, y, width, height = slot["box_cm"] if "box_cm" in slot else boxes[slot["id"]]
                display_width, display_height = _figure_footprint(figure, base_dir, width, height)
                if "box_cm" in slot:
                    geometry[slot["id"]] = [
                        x + (width - display_width) / 2, y + (height - display_height) / 2,
                        display_width, display_height,
                    ]
                figures_cm.append([round(display_width, 2), round(display_height, 2)])
            if label := slot.get("label"):
                if is_stack:
                    captions = [_lines(f["caption"], 1, "原图短标题")[0]
                                for f in figures if f.get("caption")]
                    caption = "  ｜  ".join(dict.fromkeys(captions)) or "原图证据"
                else:
                    value = figure.get("caption") or ("原图证据" if intent == "figure_evidence" else None)
                    caption = _lines(value, 1, "原图短标题")[0]
                put(label, caption, "caption", slot.get("header", slot.get("container")), [2, 0, 2, 0])
            elif (not is_stack and figure.get("caption")
                  and f"figures.{index}.caption" not in caption_fields):
                raise CandidateMismatch("该版式没有独立图标题槽")
            images[slot["id"]] = figure["path"]
        if legend:
            band, text = GUT_LEGEND
            put(text, _lines(legend, 1, "原图图例")[0], "caption", band, [2, 0, 2, 0])
        slot_profile = {**profile, "figures": figure_slots} if is_stack else profile
        for shape_id, value, slot in slot_values(page, slot_profile):
            put(shape_id, value, slot["role"], slot.get("container"), slot.get("inset_pt"))
            if "box_cm" in slot:
                geometry[shape_id] = slot["box_cm"]
            if slot.get("fixed_text"):
                fixed_text.append(shape_id)
    elif intent == "parallel_cards":
        cards = page.get("cards") or []
        if (page.get("type") != "triple" or page.get("component") or len(cards) != 3
                or len(bullets) != 1 or figures or page.get("strip")):
            raise CandidateMismatch("指南三栏需要三项、一个总问题和无图的正文")
        for number, (card, slot) in enumerate(zip(cards, profile["cards"]), 1):
            if not isinstance(card, dict) or set(card) - {"title", "note"}:
                raise CandidateMismatch("指南三栏仅接收标题和说明，保留原有栏目内容")
            put(slot["title"], _lines(card.get("title"), 1, "栏标题")[0], "label", slot["container"])
            put(slot["badge"], str(number), "label", slot["container"])
            put(slot["body"], "\n".join(_lines(card.get("note"), 2, "栏内说明")),
                "body", slot["body_container"])
    else:
        component = page.get("component") or {}
        if not isinstance(component, dict):
            raise ValueError("component 应为对象")
        items = component.get("items") or []
        if (page.get("type") != "bullets" or component.get("intent") != "risk_response"
                or set(component) - {"intent", "items", "headers"} or len(items) != 3 or figures or bullets):
            raise CandidateMismatch("指南双栏需要三组风险及应对，且无另设要点、原图或外部框")
        columns = profile["columns"]
        geometry, fixed_text = _column_geometry(columns, boxes)
        headers = component.get("headers")
        if headers is not None and (not isinstance(headers, list) or len(headers) != 2):
            raise ValueError("双栏 headers 需要两个非空单行标题")
        for i, column in enumerate(columns):
            label = ("•  " + _lines(headers[i], 1, "栏标题")[0]) if headers else column["label"]
            put(column["header"], label, "label")
        for number, (item, left, right) in enumerate(zip(items, columns[0]["rows"], columns[1]["rows"]), 1):
            if not isinstance(item, dict) or set(item) - {"title", "tag", "body"}:
                raise CandidateMismatch("风险应对项仅接收 title、tag、body")
            response = _lines(item.get("body"), 2, "应对小标题及说明")
            if len(response) != 2:
                raise CandidateMismatch("应对内容填两行，第一行用作小标题")
            risk = _lines(item.get("title"), 1, "风险小标题")[0]
            explanation = _lines(item["tag"], 1, "风险说明")[0] if item.get("tag") else ""
            for column, ids, heading, body in zip(columns, (left, right), (risk, response[0]),
                                                   (explanation, response[1])):
                container = column["container"]
                put(ids[0], str(number), "label", container)
                put(ids[1], heading, "label", container)
                if body:
                    put(ids[2], body, "body", container)
    return content, images, roles, containers, figures_cm, geometry, fixed_text


def _badge_ids(profile):
    return {card["badge"] for card in profile.get("cards", ())} | {
        row[0] for column in profile.get("columns", ()) for row in column["rows"]}


def _summary_slots(profile, boxes, nodes):
    summary = profile["summary"]
    parent = nodes[summary].getparent()
    group = next((sid for sid, node in nodes.items() if node is parent), None)
    if group is not None:
        return [group]
    x, y, width, height = boxes[summary]
    centre = x + width / 2
    artwork = []
    for ident in profile["artwork"]:
        ax, ay, aw, ah = boxes[ident]
        if ax <= centre <= ax + aw and ay <= y and y + height <= ay + ah:
            artwork.append(ident)
    return [summary, *artwork]


def _select_profile(page, base_dir, library, profile):
    source, number = profile["source"], profile["page"]
    try:
        donor = library._open(source, number)
    except ValueError as error:
        raise CandidateMismatch("素材库缺少该候选源页") from error
    with donor.z:
        if sha256(donor.z.read(donor.part)).hexdigest() != profile["slide_sha256"]:
            raise CandidateMismatch("源页已经更新，须重新核对其槽位")
        boxes = {item["id"]: item["box_cm"] for item in
                 library._inspect_nodes(donor.tree, donor, None, (1, 1, 0, 0))}
        nodes = dict(_node_map(donor.tree))
        text_sizes = dict(profile.get("text_sizes", {}))
        added_pictures = []
        normalized_groups = {}
        automatic_boxes = all(isinstance(figure, dict) and "box" not in figure
                              for figure in page.get("figures") or ())
        reflow_boxes = profile["intent"] == "figure_evidence" or (
            profile["id"] == PAIRED_READOUTS and automatic_boxes)
        reflow_metrics = profile["id"] == METRIC_PANELS and automatic_boxes
        positioned = set()
        omit = list(profile["omit"])
        try:
            content, images, roles, containers, figures, geometry, fixed = _native_content(
                page, profile, base_dir, boxes, reflow_boxes or reflow_metrics)
            if reflow_boxes or reflow_metrics:
                authored = page["figures"]
                aspects = [_dimensions(figure, base_dir) for figure in authored]
                body = profile["body"]
                body_box, top = body_geometry(profile, boxes, len(page.get("bullets") or ()))
                legend = page.get("figure_legend")
                if legend:
                    omit = [sid for sid in omit if sid not in GUT_LEGEND]
                    text_sizes[GUT_LEGEND[1]] = 16.0
                summary_boxes = {sid: geometry.get(sid, boxes[sid])
                                 for sid in _summary_slots(profile, boxes, nodes)}
                summary_top = min(box[1] for box in summary_boxes.values())
                comparison = profile["id"] != SINGLE_PANEL
                if reflow_metrics:
                    layout = balance_metric_panels
                    figures.clear()
                else:
                    layout = balance_panels
                option = {} if reflow_metrics else {"legend": legend}
                for scale in ((.92, .94, 1) if comparison else (1,)):
                    adjusted, pictures, card_bottom = layout(
                        profile, boxes, aspects, top, summary_top - .62, scale=scale, **option)
                    try:
                        for figure, box in zip(authored, pictures, strict=True):
                            _validate_footprint(figure, *box[2:])
                        for slot in content.keys() & adjusted.keys():
                            _measure_text(donor, nodes[slot], adjusted[slot], content[slot],
                                          slot in fixed, text_sizes.get(slot))
                    except CandidateMismatch:
                        if scale == 1:
                            raise
                    else:
                        break
                # Close the former source-footer gap without lifting the bar out of the bottom band.
                summary_shift = min(0, max(profile.get("summary_floor_cm", 15.70), card_bottom + .62) - summary_top)
                for sid, (x, y, width, height) in summary_boxes.items():
                    adjusted[sid] = [x, y + summary_shift, width, height]
                geometry.update(adjusted)
                positioned.update(adjusted)
                if profile["id"] == PAIRED_READOUTS:
                    group = nodes[profile["figures"][0]["id"]].getparent()
                    group_id = next(sid for sid, node in nodes.items() if node is group)
                    normalized_groups[group_id] = group
                if body_box is not None:
                    geometry[body] = body_box
                    positioned.add(body)
                    bullet_sizes = (22.0, 21.5, 20.5, 19.0, 18.0) if comparison else (20.5, 19.0, 18.0)
                    for size in bullet_sizes:
                        try:
                            _measure_text(donor, nodes[body], body_box, content[body], False, size)
                        except CandidateMismatch:
                            continue
                        text_sizes[body] = size
                        break
                    else:
                        raise CandidateMismatch("要点在 18 磅下仍需缩短或拆页")
                for i, (figure, box) in enumerate(zip(authored, pictures, strict=True)):
                    figures.append([round(value, 2) for value in box[2:]])
                    if i >= len(profile["figures"]):
                        added_pictures.append({"path": figure["path"], "box_cm": box})
                    else:
                        sid = profile["figures"][i]["id"]
                        geometry[sid] = box
                        positioned.add(sid)
            text_utilization = 0.0
            # The narrow circular badges keep the original single-digit labels.
            for slot in content.keys() - _badge_ids(profile):
                try:
                    ratio = _measure_text(donor, nodes[slot], geometry.get(slot, boxes[slot]),
                                          content[slot], slot in fixed, text_sizes.get(slot))
                except CandidateMismatch as error:
                    raise CandidateMismatch(f"原生文字槽 {slot}：{error}") from error
                text_utilization = max(text_utilization, ratio)
            for slot in sorted(positioned):
                if nodes[slot].getparent() not in normalized_groups.values():
                    geometry[slot] = coordinates_in_group(donor, nodes[slot], geometry[slot], CM)
        except (KeyError, AttributeError, TypeError) as error:
            raise CandidateMismatch("素材库的源页和已核验槽位不匹配") from error
    native = dict(source=source, page=number, text=content, images=images,
                  artwork=profile["artwork"], omit=omit)
    if "include" in profile:
        native["include"] = profile["include"]
    detail = dict(source=source, page=number, profile=profile["id"], roles=roles,
                  containers=containers, figures_cm=figures, geometry=geometry, fixed_text=fixed)
    for option in ("text_alignments", "palette"):
        if option in profile:
            detail[option] = profile[option]
    if upright := profile.get("upright_figures"):
        if not set(upright) <= images.keys():
            raise CandidateMismatch("校正方向的槽位需要当前原图")
        detail["upright_figures"] = upright
    if text_sizes:
        detail["text_sizes"] = text_sizes
    if normalized_groups:
        detail["normalized_groups"] = list(normalized_groups)
    if added_pictures:
        detail["additional_evidence"] = added_pictures
    detail["source_fields"] = {
        slot: item["field"] for slot, item in profile.get("slots", {}).items()
        if item.get("field") and slot in content
    }
    if not profile.get("summary"):
        detail["summary_box"] = profile.get("summary_box", [1.1, 17.35, 31.5, 1.05])
    detail["text_colors"] = {
        slot_id: color for slot_id, color in profile.get("text_colors", {}).items()
        if slot_id in content
    } | {
        slot_id: item["text_color"] for slot_id, item in profile.get("slots", {}).items()
        if "text_color" in item and slot_id in content
    }
    if profile.get("body") in text_sizes:
        detail["text_colors"].setdefault(profile["body"], "111820")
    usable_areas = [w * h for w, h in figures]
    # Preserve the smallest original's readability, while rewarding total
    # area and remaining text capacity. Main references win over supplements.
    score = (source in PRIMARY_SOURCES, min(usable_areas, default=0),
             sum(usable_areas), -text_utilization)
    return score, native, detail


def _record_options(decision, native_matches, supplement_matches, limit):
    main = [dict(status="selected", source=detail["source"], page=detail["page"],
                 profile=detail["profile"], figures_cm=detail["figures_cm"])
            for _, _, detail in native_matches]
    supplements = [{key: match[key] for key in ("status", "source", "page", "profile", "structure")}
                   for _, match in supplement_matches]
    ranked = main + supplements
    ordered = ranked[:limit]
    if limit == 3 and len(ranked) > limit:
        # Prefer a different reference family for the third option when it fits.
        families = {entry["source"].removesuffix("-9") for entry in ordered[:2]}
        ordered[2] = next((entry for entry in ranked[2:]
                           if entry["source"].removesuffix("-9") not in families), ordered[2])
    if len(ordered) > 1:
        decision["options"] = [dict(option, intent=decision["intent"]) for option in ordered]


def choose_native_layout(page, base_dir, archive_root, *, option_limit=0):
    """Return the build-ready page and auditable native-selection decision."""
    if "native_auto" not in page:
        return page, None
    intent = page["native_auto"]
    profiles = load_profiles(archive_root)
    intents = {profile["intent"] for profile in profiles} | SUPPLEMENT_INTENTS
    if not isinstance(intent, str) or intent not in intents:
        raise ValueError("native_auto 可选 " + "、".join(sorted(intents)))
    if page.get("type") == "native" or "native" in page or page.get("native_components"):
        raise ValueError("显式原生源页或组件请保留原有写法，移除 native_auto")
    if intent == "figure_evidence" and any(isinstance(figure, dict) and "box" in figure
                                            for figure in page.get("figures") or ()):
        if page.get("native_profile") or page.get("figure_legend"):
            raise ValueError("自定原图 box 不能与固定原生配置或自动图例条一起使用")
        return _fallback(page, intent, base_dir), dict(status="fallback", intent=intent,
                reason="使用自定原图位置", candidates=[])
    if page.get("figure_legend") is not None and (
            intent != "figure_evidence" or not isinstance(page["figure_legend"], str)
            or not page["figure_legend"].strip()):
        raise ValueError("figure_legend 须为 figure_evidence 的非空单行文字")

    candidates = []
    matches = []
    library = NativeLibrary(archive_root)
    pinned = page.get("native_profile")
    for profile in profiles:
        if profile["intent"] != intent:
            continue
        if pinned is not None and profile["id"] != pinned:
            continue
        candidate = {"source": profile["source"], "page": profile["page"], "profile": profile["id"]}
        try:
            score, native, detail = _select_profile(page, base_dir, library, profile)
        except (CandidateMismatch, FileNotFoundError, BadZipFile) as error:
            candidate["reason"] = str(error) if isinstance(error, CandidateMismatch) else "本地原生参考暂不可用"
        else:
            candidate["figures_cm"] = detail["figures_cm"]
            matches.append((score, native, detail))
        candidates.append(candidate)
    supplements = []
    collect = option_limit > 1 and pinned is None
    if collect and intent in SUPPLEMENT_INTENTS:
        supplements = list(iter_kw16_components(page, archive_root))
    if not matches:
        if intent in SUPPLEMENT_INTENTS:
            if collect:
                effective, component = next(iter(supplements), (page, None))
            else:
                effective, component = choose_kw16_component(page, archive_root)
            if component is not None and (pinned is None or component["profile"] == pinned):
                component["candidates"] = candidates
                if collect:
                    component["intent"] = intent
                    _record_options(component, matches, supplements, option_limit)
                return effective, component
        if pinned is not None:
            detail = next((row["reason"] for row in candidates), "没找到同一意图下可用的配置")
            raise ValueError(f"指定原生版式 {pinned} 不可用：{detail}")
        if page.get("figure_legend"):
            raise ValueError("自动图例条需要可用的 GUT 单图页型")
        if FORM_FIELDS & page.keys():
            reasons = "；".join(dict.fromkeys(row.get("reason", "") for row in candidates))
            raise ValueError("原生候选无法完整承载内容。请拆页或指定合适的原生槽位。"
                             + (f" {reasons}" if reasons else ""))
        return _fallback(page, intent, base_dir), dict(status="fallback", intent=intent,
            reason="已核验的原生候选无法完整承载当前内容", candidates=candidates)

    matches.sort(key=lambda entry: entry[0], reverse=True)
    _, native, detail = matches[0]
    effective = {k: deepcopy(v) for k, v in page.items()
                 if k not in {"native_auto", "native_profile", "component", "figures", "cards",
                              "strip", "bullets", "figure_legend"}
                 | FORM_FIELDS}
    effective.update(type="native", native=native)
    decision = dict(status="selected", intent=intent, candidates=candidates, **detail)
    if collect:
        _record_options(decision, matches, supplements, option_limit)
    return effective, decision
