"""Map authored checks onto the actual object IDs in a generated deck."""
from hashlib import file_digest
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn

from kw16_component_selection import apply_kw16_palette
import ppt_template as T
from source_notes import SOURCE_SHAPE_NAME, append_note_sources


UNSOURCED_FIGURE = "未标注（旧版计划）"


class ReportManifest:
    def __init__(self, audit, relative_to):
        self.strict = audit is not None
        audit = {} if audit is None else audit
        if not isinstance(audit, dict):
            raise ValueError("audit must be an object")
        terms = audit.get("identity_terms", [])
        if not isinstance(terms, list) or any(not isinstance(t, str) or not t.strip() for t in terms):
            raise ValueError("audit.identity_terms must be a list of nonempty strings")
        figures = audit.get("figure_sources", {})
        if not isinstance(figures, dict):
            raise ValueError("audit.figure_sources must map figure paths to sources")
        self.data = {"identity_terms": terms, "slides": {}}
        self.figures = {
            Path(relative_to(path)).resolve(): source for path, source in figures.items()
        }
        self.relative_to = relative_to

    def new_slide(self, index):
        return self.data["slides"].setdefault(str(index), {
            "shapes": [], "evidence": [], "containers": [],
        })

    def evidence(self, record, shape_id, path, page_source):
        path = Path(self.relative_to(path)).resolve()
        provenance = self.figures.get(path)
        if provenance is None:
            if self.strict:
                raise ValueError(f"missing audit.figure_sources for one image on this slide ({path.name})")
            provenance = {"source": page_source or ""}
        if not isinstance(provenance, dict) or not str(provenance.get("source", "")).strip():
            # Older example plans place figures without a source caption.
            if self.strict or provenance != {"source": ""}:
                raise ValueError("figure provenance needs a nonempty source")
            provenance = {"source": UNSOURCED_FIGURE}
        with path.open("rb") as file:
            digest = file_digest(file, "sha256").hexdigest()
        if provenance.get("sha256", digest) != digest:
            raise ValueError("figure SHA-256 differs from the source declared in the plan")
        fields = {key: provenance[key] for key in ("source", "pdf_page", "crop", "panel")
                  if key in provenance}
        record["evidence"].append({"shape_id": shape_id, "sha256": digest, **fields})


def native_records(slide, order, manifest, record, spec, remap, page_source):
    """Bind declared source IDs to editable objects and inserted evidence."""
    roles = order.get("roles", {})
    contained = {str(item["shape_id"]) for item in order.get("containers", [])}
    for source_id, role in roles.items():
        if source_id not in remap:
            raise ValueError(f"audit role references missing native shape {source_id}")
        binding = {"shape_id": remap[source_id], "role": role}
        if source_id in order.get("source_fields", {}):
            binding["field"] = order["source_fields"][source_id]
        if source_id in contained:
            binding["container_required"] = True
        record["shapes"].append(binding)
    for item in order.get("containers", []):
        child, parent = str(item["shape_id"]), str(item["container_id"])
        if child not in remap or parent not in remap:
            raise ValueError(f"card binding references a missing native shape ({child}, {parent})")
        record["containers"].append({
            "shape_id": remap[child],
            "container_id": remap[parent],
            "inset_pt": item.get("inset_pt", [0, 0, 0, 0]),
        })
    for source_id, path in spec.get("images", {}).items():
        manifest.evidence(record, remap[source_id], path, page_source)
    if order.get("status") == "selected":
        for extra in order.get("additional_evidence", ()):
            box = (T.cm(value) for value in extra["box_cm"])
            picture = slide.shapes.add_picture(str(manifest.relative_to(extra["path"])), *box)
            picture.name = "原图面板"
            manifest.evidence(record, picture.shape_id, extra["path"], page_source)
    for source_id in spec.get("tables", {}):
        if source_id not in roles:
            record["shapes"].append({"shape_id": remap[source_id], "role": "table"})
    geometry = order.get("geometry", {})
    fixed_text = order.get("fixed_text", [])
    text_sizes = order.get("text_sizes", {})
    palette = order.get("palette", {})
    text_colors = dict(order.get("text_colors", {}))
    auto_selected = order.get("status") == "selected"
    if not (geometry or fixed_text or auto_selected or palette or text_colors or text_sizes):
        return
    shapes = {shape.shape_id: shape for shape in T.iter_shapes(slide.shapes, include_groups=True)}
    if palette:
        apply_kw16_palette(shapes, remap, palette)
    for source_id in fixed_text:
        frame = shapes[remap[source_id]].text_frame
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.word_wrap = True
        frame.margin_left = frame.margin_right = 0
    for source_id in order.get("normalized_groups", ()) if auto_selected else ():
        # Children supplied in slide coordinates can retain their literal aspect ratio.
        transform = shapes[remap[source_id]]._element.grpSpPr.xfrm
        transform.chOff.x, transform.chOff.y = transform.off.x, transform.off.y
        transform.chExt.cx, transform.chExt.cy = transform.ext.cx, transform.ext.cy
    for source_id in order.get("upright_figures", ()) if auto_selected else ():
        transform = shapes[remap[source_id]]._element.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
        for attribute in ("rot", "flipH", "flipV"):
            transform.attrib.pop(attribute, None)
    for source_id, box in geometry.items():
        shape = shapes[remap[source_id]]
        shape.left, shape.top, shape.width, shape.height = (T.cm(value) for value in box)
    alignment_values = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for source_id, alignment in order.get("text_alignments", {}).items():
        frame = shapes[remap[source_id]].text_frame
        for paragraph in frame.paragraphs:
            paragraph.alignment = alignment_values[alignment]
    for source_id, size in text_sizes.items():
        frame = shapes[remap[source_id]].text_frame
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.word_wrap = False
        for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
            for properties in frame._txBody.iter(qn(tag)):
                T._set_sz(properties, size)
    if auto_selected:
        # Some source summaries start with a red run. Keep the red emphasis tied
        # to the authored string so that the inherited style cannot turn it all red.
        for source_id, role in roles.items():
            if role == "summary":
                text_colors.setdefault(source_id, "111820")
    for source_id, base_color in text_colors.items():
        text = spec["text"][source_id]
        paragraphs = shapes[remap[source_id]].text_frame.paragraphs
        for line, paragraph in zip(text.split("\n"), paragraphs, strict=True):
            for (_, emphasized), run in zip(T.segments(line), paragraph.runs, strict=True):
                run.font.color.rgb = RGBColor.from_string(T.RED if emphasized else base_color.lstrip("#"))


def _place_note_sources(slide, page, record, cover):
    sources = [page.get("source")]
    source_slots = {item["shape_id"]: item for item in record["shapes"] if item["role"] == "source"}
    cover_credit = (cover and page.get("type") == "native" and not page.get("source")
                    and not record["evidence"]
                    and all(item["role"] not in {"body", "summary"} for item in record["shapes"]))
    dropped = set()
    for shape in list(T.iter_shapes(slide.shapes, include_groups=True)):
        item = source_slots.get(shape.shape_id)
        if item is not None and (not shape.has_text_frame or not shape.text.strip()):
            raise ValueError("原生图源槽位缺少可编辑文字")
        if item is not None and cover_credit and shape.name != SOURCE_SHAPE_NAME:
            item["role"] = "label"
        elif item is not None or shape.name == SOURCE_SHAPE_NAME:
            if item is not None:
                sources.append(shape.text)
            dropped.add(shape.shape_id)
            T.drop(shape)
    record["shapes"][:] = [item for item in record["shapes"] if item["shape_id"] not in dropped]
    record["containers"][:] = [
        item for item in record["containers"]
        if not dropped.intersection((item["shape_id"], item["container_id"]))
    ]
    if not any(value and (not isinstance(value, str) or value.strip()) for value in sources):
        sources = [item["source"] for item in record["evidence"] if item["source"] != UNSOURCED_FIGURE]
    notes, record["notes_sources"] = append_note_sources(page.get("notes"), sources)
    if notes:
        T.set_notes(slide, notes)


def finalize_slide(slide, page, record, strict, *, cover=False):
    """Record standard template roles and requirements independent of shape names."""
    _place_note_sources(slide, page, record, cover)
    names = {
        "T_Title": "title", "D_Title": "title", "页标题": "title",
        "要点1": "body", "总结文字": "summary",
        "卡片标题": "label", "卡片脚注": "caption",
    }
    for shape in T.iter_shapes(slide.shapes):
        if shape.has_table and page.get("type") != "native":
            record["shapes"].append({"shape_id": shape.shape_id, "role": "table"})
        if shape.name in names and shape.has_text_frame and shape.text.strip():
            record["shapes"].append({"shape_id": shape.shape_id, "role": names[shape.name]})

    roles = {shape["role"] for shape in record["shapes"]}
    content = (page.get("type") in ("bullets", "table", "cards2", "triple")
               or page.get("type") == "native" and (page.get("source") or "summary" in roles))
    record["notes_required"] = bool(strict and (content or record["evidence"]))
    record["source_required"] = bool(record["notes_sources"] or (strict and record["evidence"]))
    if strict and (content or record["evidence"]):
        if not {"title", "summary"} <= roles:
            raise ValueError("audited content pages need editable title and summary roles")
