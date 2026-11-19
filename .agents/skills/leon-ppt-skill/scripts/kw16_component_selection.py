"""Select vetted editable KW16 structures inside the report's blue-white body."""

from copy import deepcopy
from hashlib import sha256
import json
from pathlib import Path
from zipfile import BadZipFile

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

from native_reference import NativeLibrary
from native_distillation import FORM_FIELDS
from ppt_template import RED, cm, line_width_pt, segments


RECIPE_FILE = Path(__file__).resolve().parents[1] / "references" / "kw16-distilled-components.json"
SUPPLEMENT_INTENTS = {
    "milestones", "parallel_cards", "paired_comparison", "evidence_cluster", "stage_matrix",
}
COLORS = {
    "none": None, "navy": "173467", "blue": "2F5597", "white": "FFFFFF",
    "pale": "EAF1FA", "mid": "D3E4F5", "outline": "AAC5E2",
    "text": "111820", "gray": "53677E",
}
STYLE = {
    "heading": ("navy", 18, True, "center", "middle"),
    "body": ("text", 16, False, "left", "top"),
    "cluster_body": ("text", 18, False, "left", "middle"),
    "cluster_label": ("white", 18, True, "center", "middle"),
    "center_body": ("text", 16, False, "center", "middle"),
    "left_body": ("text", 16, False, "left", "middle"),
    "right_body": ("text", 16, False, "right", "middle"),
    "right_label": ("navy", 17, True, "right", "middle"),
    "label": ("navy", 17, True, "left", "middle"),
    "white_label": ("white", 17, True, "center", "middle"),
    "mark": ("white", 20, True, "center", "middle"),
    "tag": ("gray", 14, False, "center", "middle"),
}
FONT = "Microsoft YaHei"
HORIZONTAL_PAD = 3
VERTICAL_PAD = 2
PT_PER_CM = 72 / 2.54


class ComponentMismatch(ValueError):
    """Current content cannot fit a verified supplement."""


def _payload(page, intent):
    """Return only the content fields a supplement is allowed to consume."""
    if (FORM_FIELDS & page.keys()) - {"items", "headers", "hub"}:
        raise ComponentMismatch("补充组件无法承载额外结构字段")
    if (page.get("type") not in {"bullets", "triple"} or page.get("native")
            or page.get("native_components") or page.get("figures") or page.get("strip")):
        raise ComponentMismatch("需要无图的常规正文页")
    if page["type"] == "triple":
        if intent != "parallel_cards" or "items" in page or "component" in page:
            raise ComponentMismatch("三栏页须使用 cards 填充")
        cards = page.get("cards", [])
        if not isinstance(cards, list) or any(not isinstance(card, dict) or
                set(card) != {"title", "note"} for card in cards):
            raise ComponentMismatch("三栏仅接收 title 和 note")
        return {"items": [{"title": card["title"], "body": card["note"]} for card in cards]}
    component = page.get("component")
    if component is not None:
        aliases = {"deliverable_checkpoints": "milestones", "independent_findings": "parallel_cards",
                   "same_basis_comparison": "paired_comparison"}
        if (not isinstance(component, dict) or set(component) != {"intent", "items"}
                or "items" in page or aliases.get(component.get("intent"), component.get("intent")) != intent):
            raise ComponentMismatch("保留用户指定的正文组件")
    items = page["items"] if "items" in page else (component or {}).get("items")
    payload = {"items": items}
    for key in ("headers", "hub"):
        if key in page:
            payload[key] = page[key]
    return payload


def _flatten(value, prefix=""):
    if isinstance(value, dict):
        pairs = value.items()
    elif isinstance(value, list):
        pairs = enumerate(value)
    else:
        return {prefix: value}
    result = {}
    for key, child in pairs:
        result.update(_flatten(child, f"{prefix}.{key}" if prefix else str(key)))
    return result


def _records(donor, library, include, omit):
    records = {entry["id"]: entry for entry in
               library._inspect_nodes(donor.tree, donor, None, (1, 1, 0, 0))}
    retained = set()
    # Inspection is pre-order, so a group is known before any of its children.
    for identifier, entry in records.items():
        if identifier in omit or (identifier not in include and entry["parent"] not in retained):
            continue
        if entry["type"] not in {"text", "shape", "group", "connector"}:
            raise ComponentMismatch("源组件含无法安全复制的非原生形状")
        retained.add(identifier)
    if not set(include) <= retained or any(records[i]["parent"] is not None for i in include):
        raise ComponentMismatch("源组件的顶层对象已变化")
    return records, retained


def _scale_to_box(records, roots, box):
    raw = [records[root]["box_cm"] for root in roots]
    if any(frame is None for frame in raw):
        raise ComponentMismatch("源组件边界缺失")
    width = max(x + width for x, _, width, _ in raw) - min(x for x, *_ in raw)
    height = max(y + height for _, y, _, height in raw) - min(y for _, y, _, _ in raw)
    if min(width, height) <= 0:
        raise ComponentMismatch("源组件边界无效")
    return min(box[2] / width, box[3] / height)


def _check_text(value, slot, frame, scale):
    if not isinstance(value, str) or not value.strip():
        raise ComponentMismatch(f"{slot.get('field', '标号')} 需要文字")
    lines = value.strip().splitlines()
    if len(lines) > slot.get("max_lines", 1) or any(not line.strip() for line in lines):
        raise ComponentMismatch(f"{slot.get('field', '标号')} 超出所选布局允许的行数")
    _, size, bold, _, _ = STYLE[slot["style"]]
    actual_width = frame[2] * PT_PER_CM * scale - 2 * HORIZONTAL_PAD
    actual_height = frame[3] * PT_PER_CM * scale - 2 * VERTICAL_PAD
    actual_size = size * scale
    widest = max(line_width_pt(line, FONT, FONT, bold, actual_size) for line in lines)
    if (actual_size < 12 or actual_height < actual_size * 1.20 * len(lines)
            or widest > actual_width * 0.96):
        raise ComponentMismatch(f"{slot.get('field', '标号')} 超出原生槽的可读空间")
    return "\n".join(lines)


def _compile(page, profile, donor, library):
    payload = _payload(page, profile["intent"])
    if not isinstance(payload.get("items"), list) or len(payload["items"]) != profile["item_count"]:
        raise ComponentMismatch(f"该组件需要 {profile['item_count']} 项内容")
    fields = _flatten(payload)
    expected = {slot["field"] for slot in profile["slots"].values() if "field" in slot}
    if set(fields) != expected:
        raise ComponentMismatch("当前内容字段与组件槽位不匹配")
    include = profile["include"]
    omit = profile.get("omit", [])
    records, retained = _records(donor, library, include, omit)
    styling = deepcopy(profile["palette"])
    for shape_id, style in styling.items():
        if "box_cm" in style:
            if shape_id not in retained or records[shape_id]["parent"] is not None:
                raise ComponentMismatch("仅支持调整已选用的顶层对象边界")
            records[shape_id]["box_cm"] = style["box_cm"]
    scale = _scale_to_box(records, include, profile["box"])
    text, roles, containers = {}, {}, []
    for shape_id, slot in profile["slots"].items():
        if shape_id not in retained or records[shape_id]["type"] != "text":
            raise ComponentMismatch("源文字槽已变化")
        entry = records[shape_id]
        if entry["rotation"] or not entry["box_cm"]:
            raise ComponentMismatch("文字槽需要可核验的水平边界")
        value = fields[slot["field"]] if "field" in slot else slot["literal"]
        text[shape_id] = _check_text(value, slot, entry["box_cm"], scale)
        styling.setdefault(shape_id, {}).update(style=slot["style"], text=text[shape_id])
        roles[shape_id] = slot["role"]
        if "container" in slot:
            if slot["container"] not in retained:
                raise ComponentMismatch("源文字槽的所属容器已变化")
            containers.append({"shape_id": shape_id, "container_id": slot["container"],
                               "inset_pt": slot.get("inset_pt", [1, 1, 1, 1])})
    if not styling.keys() <= retained:
        raise ComponentMismatch("源样式对象已变化")
    source_fields = {shape_id: slot["field"] for shape_id, slot in profile["slots"].items()
                     if "field" in slot}
    return {"source": profile["source"], "page": profile["page"], "include": include, "omit": omit,
            "text": text, "box": profile["box"],
            "audit": {"roles": roles, "containers": containers, "palette": styling,
                      "source_fields": source_fields}}


def choose_kw16_component(page, archive_root):
    """Try a source-hash-bound editable supplement after the main candidates."""
    intent = page.get("native_auto")
    if not isinstance(intent, str) or intent not in SUPPLEMENT_INTENTS:
        return page, None
    requested_profile = page.get("native_profile")
    profiles = json.loads(RECIPE_FILE.read_text(encoding="utf-8"))["components"]
    library = NativeLibrary(archive_root)
    for profile in profiles:
        if profile["intent"] != intent or (requested_profile and profile["id"] != requested_profile):
            continue
        try:
            donor = library._open(profile["source"], profile["page"])
        except (FileNotFoundError, BadZipFile, ValueError):
            continue
        with donor.z:
            if sha256(donor.z.read(donor.part)).hexdigest() != profile["slide_sha256"]:
                continue
            try:
                component = _compile(page, profile, donor, library)
            except ComponentMismatch:
                continue
        result = {key: value for key, value in page.items()
                  if key not in {"native_auto", "native_profile", "items", "cards", "component", "headers", "hub"}}
        result.update(type="bullets", native_components=[component])
        decision = {"status": "selected_component", "source": profile["source"],
                    "page": profile["page"], "profile": profile["id"], "recipe": profile["id"],
                    "panel_cm": profile["box"], "structure": profile["structure"]}
        return result, decision
    return page, None


def _color(fill, token):
    value = COLORS[token]
    if value is None:
        fill.background()
    else:
        fill.solid()
        color = fill.fore_color
        color.rgb = RGBColor.from_string(value)
        # Replacing an RGB value leaves donor tint and alpha in place.
        rgb = color._xFill.find(qn("a:srgbClr"))
        for inherited in list(rgb):
            rgb.remove(inherited)


def _font(font, size, bold, color):
    font.name = FONT
    font.size = Pt(size)
    font.bold = bold
    font.color.rgb = RGBColor.from_string(color)
    east_asian = OxmlElement("a:ea")
    east_asian.set("typeface", FONT)
    font._element.append(east_asian)


def _format_text(shape, style, contents):
    color, size, bold, align, anchor = STYLE[style]
    frame = shape.text_frame
    frame.clear()
    donor_defaults = frame._txBody.find(qn("a:lstStyle"))
    if donor_defaults is not None:
        donor_defaults.clear()
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.word_wrap = False
    frame.margin_left = frame.margin_right = Pt(HORIZONTAL_PAD)
    frame.margin_top = frame.margin_bottom = Pt(VERTICAL_PAD)
    frame.vertical_anchor = MSO_ANCHOR.TOP if anchor == "top" else MSO_ANCHOR.MIDDLE
    for index, line in enumerate(contents.splitlines()):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        for inherited in list(paragraph._p):
            paragraph._p.remove(inherited)
        paragraph.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                               "right": PP_ALIGN.RIGHT}[align]
        paragraph.space_before = paragraph.space_after = Pt(0)
        paragraph.line_spacing = 1
        paragraph._p.get_or_add_pPr().append(OxmlElement("a:buNone"))
        _font(paragraph.font, size, bold, COLORS[color])
        for string, emphasized in segments(line):
            run = paragraph.add_run()
            run.text = string
            _font(run.font, size, bold, RED if emphasized else COLORS[color])


def apply_kw16_palette(shapes_by_final_id, remap, styling):
    """Keep cloned shapes editable, with the main reference's color hierarchy."""
    for source_id, style in styling.items():
        shape = shapes_by_final_id[remap[source_id]]
        if "box_cm" in style:
            shape.left, shape.top, shape.width, shape.height = (cm(v) for v in style["box_cm"])
        if "fill" in style:
            _color(shape.fill, style["fill"])
        if "line" in style:
            _color(shape.line.fill, style["line"])
            if style["line"] != "none":
                shape.line.width = Pt(0.85)
        if "style" in style:
            _format_text(shape, style["style"], style["text"])
