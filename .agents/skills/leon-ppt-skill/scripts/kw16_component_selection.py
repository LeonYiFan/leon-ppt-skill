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

from native_reference import A, CM, NativeLibrary, _geometry, _node_map
from native_distillation import FORM_FIELDS
from ppt_template import RED, cm, line_width_pt, segments


RECIPE_FILE = Path(__file__).resolve().parents[1] / "references" / "kw16-distilled-components.json"
EXTRA_RECIPE_FILE = RECIPE_FILE.with_name("supplemental-components.json")
SUPPLEMENT_INTENTS = {
    "hypothesis_branches",
    "milestones", "parallel_cards", "paired_comparison", "evidence_cluster", "stage_matrix",
    "paired_schedule", "convergent_validation", "stage_handoffs",
    "comparison_table", "gantt_schedule",
}
STRUCTURED_FIELDS = {
    "hypothesis_branches": {"items", "lead"},
    "paired_schedule": {"items", "stages"},
    "convergent_validation": {"items", "lead", "outcome"},
    "stage_handoffs": {"items"},
    "comparison_table": {"detail", "table"},
    "gantt_schedule": {"timeline"},
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
    allowed = STRUCTURED_FIELDS.get(intent, {"items", "headers", "hub"})
    if (FORM_FIELDS & page.keys()) - allowed:
        raise ComponentMismatch("补充组件无法承载额外结构字段")
    if (page.get("type") not in {"bullets", "triple"} or page.get("native")
            or page.get("native_components") or page.get("figures") or page.get("strip")):
        raise ComponentMismatch("需要无图的常规正文页")
    if intent in STRUCTURED_FIELDS:
        if page["type"] != "bullets" or page.keys() & {"component", "cards"} or (
                "table" in page and intent != "comparison_table"):
            raise ComponentMismatch("结构化补充布局使用独立正文页")
        return {key: page[key] for key in allowed & page.keys()}
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


def _records(donor, library, include, omit, table_id):
    records = {entry["id"]: entry for entry in
               library._inspect_nodes(donor.tree, donor, None, (1, 1, 0, 0))}
    retained = set()
    # Inspection is pre-order, so a group is known before any of its children.
    for identifier, entry in records.items():
        if identifier in omit or (identifier not in include and entry["parent"] not in retained):
            continue
        if entry["type"] not in {"text", "shape", "group", "connector"} and (
                identifier != table_id or entry["type"] != "table"):
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


def _group_box(donor, nodes, shape_id, box):
    """Convert a slide-relative rectangle to an editable child's group coordinates."""
    mapping = (1, 1, 0, 0)
    for ancestor in reversed(tuple(nodes[shape_id].iterancestors())):
        if ancestor.tag.endswith("}grpSp"):
            _, mapping = _geometry(ancestor, donor, mapping)
    sx, sy, dx, dy = mapping
    x, y, w, h = box
    return [(x - dx / CM) / sx, (y - dy / CM) / sy, w / sx, h / sy]


def _table_data(payload, profile, donor, row_box):
    """Populate only authored cells and return grid decoration + real donor dimensions."""
    intent = profile["intent"]
    if intent == "comparison_table":
        spec = payload.pop("table", None)
        if not isinstance(spec, dict) or set(spec) != {"rows"}:
            raise ComponentMismatch("学术对照表需要 table.rows")
        rows = spec["rows"]
        if (not isinstance(rows, list) or not 3 <= len(rows) <= 6 or not isinstance(rows[0], list)
                or not 3 <= len(rows[0]) <= 6 or any(not isinstance(row, list) or
                len(row) != len(rows[0]) for row in rows)):
            raise ComponentMismatch("学术对照表需要 3–6 行（含表头）、3–6 列")
        rows_width = [row_box[2] / len(rows[0])] * len(rows[0])
        schedule = None
    else:
        record = payload.pop("timeline", None)
        if not isinstance(record, dict) or set(record) != {"periods", "items"}:
            raise ComponentMismatch("甘特日程需要 timeline.periods 和 timeline.items")
        periods, items = record["periods"], record["items"]
        if (not isinstance(periods, list) or not 4 <= len(periods) <= 6 or
                not isinstance(items, list) or not 3 <= len(items) <= 5):
            raise ComponentMismatch("甘特日程需要 4–6 个等长时段和 3–5 项任务")
        rows, schedule = [["任务", *periods]], []
        for item in items:
            if (not isinstance(item, dict) or item.keys() != {"title", "start", "end"} or
                    any(type(item[k]) is not int for k in ("start", "end")) or
                    not 1 <= item["start"] <= item["end"] <= len(periods)):
                raise ComponentMismatch("任务区间须为合法的 1 起算时段索引，含起止时段")
            rows.append([item["title"], *["" for _ in periods]])
            schedule.append([item["start"], item["end"]])
        task_width = 7.0
        rows_width = [task_width, *[(row_box[2] - task_width) / len(periods)] * len(periods)]
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if not isinstance(value, str) or ((r == 0 or c == 0) and not value.strip()):
                raise ComponentMismatch("表头、任务名和首列需用非空文字，其它表格值需用文字")
            if value.strip():
                frame = [0, 0, rows_width[c], row_box[3] / len(rows)]
                style = "heading" if r == 0 or c == 0 else "center_body"
                _check_text(value, {"style": style, "max_lines": 1, "field": f"第 {r+1} 行第 {c+1} 列"},
                            frame, 1)
    table_id = profile["table"]
    node = dict(_node_map(donor.tree))[table_id].find(f".//{A}tbl")
    donor_cols = len(node.find(A + "tblGrid"))
    donor_rows = len(node.findall(A + "tr"))
    # The existing native clone accepts a donor-sized array. Trim the empty
    # cells afterwards, then restore the authored viewport on the clone.
    padded = [row + [""] * (donor_cols - len(row)) for row in rows]
    padded += [[""] * donor_cols for _ in range(donor_rows - len(rows))]
    return padded, {"rows": len(rows), "widths": rows_width, "height": row_box[3],
                    "schedule": schedule}


def _compile(page, profile, donor, library):
    payload = _payload(page, profile["intent"])
    if "item_count" in profile and (
            not isinstance(payload.get("items"), list) or len(payload["items"]) != profile["item_count"]):
        raise ComponentMismatch(f"该组件需要 {profile['item_count']} 项内容")
    table_id = profile.get("table")
    table_values = {}
    styling = deepcopy(profile["palette"])
    if table_id:
        rows, table_style = _table_data(payload, profile, donor, styling[table_id]["box_cm"])
        table_values[table_id] = rows
        styling[table_id]["table"] = table_style
    fields = _flatten(payload)
    expected = {slot["field"] for slot in profile["slots"].values() if "field" in slot}
    if set(fields) != expected:
        raise ComponentMismatch("当前内容字段与组件槽位不匹配")
    include = profile["include"]
    omit = profile.get("omit", [])
    records, retained = _records(donor, library, include, omit, table_id)
    for shape_id, style in styling.items():
        if "box_cm" in style:
            if shape_id not in retained:
                raise ComponentMismatch("仅支持调整已选用的对象边界")
            records[shape_id]["box_cm"] = style["box_cm"]
        elif "points_cm" in style:
            if shape_id not in retained or records[shape_id]["type"] != "connector":
                raise ComponentMismatch("端点须绑定已选用的原生连接线")
            x1, y1, x2, y2 = style["points_cm"]
            records[shape_id]["box_cm"] = [min(x1, x2), min(y1, y2), abs(x2-x1), abs(y2-y1)]
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
    child_boxes = {sid: style for sid, style in styling.items()
                   if "box_cm" in style and records[sid]["parent"] is not None}
    if child_boxes:
        nodes = dict(_node_map(donor.tree))
        for sid, style in child_boxes.items():
            style["box_cm"] = _group_box(donor, nodes, sid, style["box_cm"])
    source_fields = {shape_id: slot["field"] for shape_id, slot in profile["slots"].items()
                     if "field" in slot}
    return {"source": profile["source"], "page": profile["page"], "include": include, "omit": omit,
            "text": text, "tables": table_values, "box": profile["box"],
            "audit": {"roles": roles, "containers": containers, "palette": styling,
                      "source_fields": source_fields}}


def load_component_profiles(archive_root=None):
    directory = Path(archive_root) / "profiles" if archive_root is not None else RECIPE_FILE.parent
    if not directory.is_dir():
        directory = RECIPE_FILE.parent
    paths = [directory / path.name for path in (RECIPE_FILE, EXTRA_RECIPE_FILE)
             if (directory / path.name).is_file()]
    return [row for path in paths for row in json.loads(path.read_text(encoding="utf-8"))["components"]]


def iter_kw16_components(page, archive_root):
    """Yield compatible source-hash-bound supplements in configured priority order."""
    intent = page.get("native_auto")
    if not isinstance(intent, str) or intent not in SUPPLEMENT_INTENTS:
        return
    requested_profile = page.get("native_profile")
    profiles = load_component_profiles(archive_root)
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
        consumed = STRUCTURED_FIELDS.get(intent, {"items", "headers", "hub"})
        result = {key: value for key, value in page.items()
                  if key not in consumed | {"native_auto", "native_profile", "cards", "component"}}
        result.update(type="bullets", native_components=[component])
        decision = {"status": "selected_component", "source": profile["source"],
                    "page": profile["page"], "profile": profile["id"], "recipe": profile["id"],
                    "panel_cm": profile["box"], "structure": profile["structure"]}
        yield result, decision


def choose_kw16_component(page, archive_root):
    """Choose the first compatible editable supplement."""
    return next(iter_kw16_components(page, archive_root), (page, None))


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


def _format_table(shape, style):
    table = shape.table
    widths = style["widths"]
    rows = list(table._tbl.findall(qn("a:tr")))
    for row in rows[style["rows"]:]:
        table._tbl.remove(row)
    grid = table._tbl.find(qn("a:tblGrid"))
    for column in list(grid)[len(widths):]:
        grid.remove(column)
    for index, width in enumerate(widths):
        table.columns[index].width = cm(width)
    for r, row in enumerate(table.rows):
        cells = row._tr.findall(qn("a:tc"))
        for cell in cells[len(widths):]:
            row._tr.remove(cell)
        row.height = cm(style["height"] / style["rows"])
        for c, cell in enumerate(row.cells):
            schedule = style["schedule"]
            active = bool(schedule and r and schedule[r-1][0] <= c <= schedule[r-1][1])
            fill = "blue" if r == 0 or active else "pale" if r % 2 else "white"
            _color(cell.fill, fill)
            cell.margin_left = cell.margin_right = Pt(HORIZONTAL_PAD)
            cell.margin_top = cell.margin_bottom = Pt(VERTICAL_PAD)
            text_style = "white_label" if r == 0 else "label" if c == 0 else "center_body"
            _format_text(cell, text_style, cell.text)
            properties = cell._tc.get_or_add_tcPr()
            for name in ("lnB", "lnT", "lnR", "lnL"):
                old = properties.find(qn(f"a:{name}"))
                if old is not None:
                    properties.remove(old)
                edge = OxmlElement(f"a:{name}")
                edge.set("w", "6350")
                solid = OxmlElement("a:solidFill")
                rgb = OxmlElement("a:srgbClr")
                rgb.set("val", COLORS["outline"])
                solid.append(rgb)
                edge.append(solid)
                properties.insert(0, edge)


def _format_connector(shape, points):
    begin_x, begin_y, end_x, end_y = map(cm, points)
    shape.rotation = 0
    shape.begin_x, shape.begin_y = begin_x, begin_y
    shape.end_x, shape.end_y = end_x, end_y
    properties = shape._element.spPr
    geometry = properties.find(qn("a:prstGeom"))
    geometry.set("prst", "line")
    for child in list(geometry):
        geometry.remove(child)
    geometry.append(OxmlElement("a:avLst"))
    binding = shape._element.find(qn("p:nvCxnSpPr")).find(qn("p:cNvCxnSpPr"))
    for name in ("stCxn", "endCxn"):
        old = binding.find(qn(f"a:{name}"))
        if old is not None:
            binding.remove(old)
    line = properties.get_or_add_ln()
    for name in ("headEnd", "tailEnd"):
        old = line.find(qn(f"a:{name}"))
        if old is not None:
            line.remove(old)
        arrow = OxmlElement(f"a:{name}")
        arrow.set("type", "none")
        line.append(arrow)


def apply_kw16_palette(shapes_by_final_id, remap, styling):
    """Keep cloned shapes editable, with the main reference's color hierarchy."""
    for source_id, style in styling.items():
        shape = shapes_by_final_id[remap[source_id]]
        if "box_cm" in style:
            shape.left, shape.top, shape.width, shape.height = (cm(v) for v in style["box_cm"])
        if "points_cm" in style:
            _format_connector(shape, style["points_cm"])
        if "table" in style:
            _format_table(shape, style["table"])
        if style.get("flat"):
            properties = shape._element.find(qn("p:spPr"))
            if properties is not None:
                for child in list(properties):
                    if child.tag in {qn("a:effectLst"), qn("a:effectDag"), qn("a:scene3d"), qn("a:sp3d")}:
                        properties.remove(child)
            inherited_style = shape._element.find(qn("p:style"))
            if inherited_style is not None:
                if "fill" in style and "line" in style:
                    shape._element.remove(inherited_style)
                else:
                    effect = inherited_style.find(qn("a:effectRef"))
                    if effect is not None:
                        effect.set("idx", "0")
        if style.get("upright"):
            transform = shape._element.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
            if transform is None:
                transform = shape._element.find(f"{qn('p:grpSpPr')}/{qn('a:xfrm')}")
            if transform is not None:
                for attribute in ("rot", "flipH", "flipV"):
                    transform.attrib.pop(attribute, None)
            if shape.has_text_frame:
                body = shape.text_frame._txBody.bodyPr
                body.set("vert", "horz")
                body.attrib.pop("rot", None)
        if style.get("front"):
            element = shape._element
            parent = element.getparent()
            extension = parent.find(qn("p:extLst"))
            if extension is None:
                parent.append(element)
            else:
                extension.addprevious(element)
        if "fill" in style:
            _color(shape.fill, style["fill"])
        if "line" in style:
            _color(shape.line.fill, style["line"])
            if style["line"] != "none":
                shape.line.width = Pt(0.85)
        if "style" in style:
            # Some donor drawing shapes have degenerate text bounds. Keep
            # their graphic path and use the resized rectangle for the text.
            text_box = shape._element.find(
                f"{qn('p:spPr')}/{qn('a:custGeom')}/{qn('a:rect')}"
            )
            if text_box is not None and set(text_box.attrib.values()) == {"0"}:
                text_box.attrib.update({"l": "l", "t": "t", "r": "r", "b": "b"})
            _format_text(shape, style["style"], style["text"])
