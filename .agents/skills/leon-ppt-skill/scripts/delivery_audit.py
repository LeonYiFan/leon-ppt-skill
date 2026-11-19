"""Audit saved PPTX parts and the locally authored delivery manifest."""
from __future__ import annotations

import io
import posixpath
import re
from hashlib import sha256
from math import isclose
from zipfile import BadZipFile, ZipFile

from lxml import etree
from PIL import Image

from artifact_privacy import BRANDED_IMAGES, LOGO_LABEL, has_identity
from source_notes import PREFIX, SOURCE_SHAPE_NAME, source_key

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
XML_SUFFIXES = (".xml", ".rels", ".svg", ".vml")
RASTER_SUFFIXES = (".bmp", ".gif", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".webp")
BINARY_VECTOR_SUFFIXES = (".emf", ".wmf")
IDENTITY_TAGS = {
    "author", "authors", "creator", "lastmodifiedby", "lastauthor", "company", "manager",
    "contributor", "contact", "email", "affiliation", "artist", "owner",
    "createdby", "modifiedby", "personname",
}
IDENTITY_ATTRS = IDENTITY_TAGS | {"username", "userid", "authorid"}
MEDIA_PROFILE_KEYS = {
    "iptc", "xmp", "rawprofiletypeiptc", "rawprofiletypexmp", "xmlcomadobexmp",
    "photoshop", "comment", "usercomment", "description", "imagedescription",
    "copyright", "rights", "gpsinfo", "gps", "serialnumber", "bodyserialnumber",
    "lensserialnumber",
}
# Free text, equipment IDs, GPS, proprietary or people-related metadata.
# Resolution, orientation, ICC color profile and other technical tags remain usable.
PRIVATE_EXIF = {
    269, 270, 285, 315, 700, 33432, 33723, 34377, 34853, 37500, 37510,
    40091, 40092, 40093, 40094, 40095, 42016, 42032, 42033, 42037,
}
TEXT_ROLES = {"title", "body", "summary", "source", "label", "caption"}
EDITABLE_TAGS = {P + "sp", P + "cxnSp"}
EXTENSION_SHAPES = EDITABLE_TAGS | {P + "grpSp", P + "graphicFrame", P + "pic"}
SHA256 = re.compile(r"[0-9a-f]{64}")
PERSON_PART_ROOTS = {"cmAuthorLst", "personLst", "authors", "authorList", "people"}
PRIVATE_TARGET = re.compile(r"^(?:file:|mailto:|tel:|[a-z]:[\\/]|\\\\)", re.IGNORECASE)


def _xml(data):
    return etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def _local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _key(value):
    return re.sub(r"[^a-z0-9]", "", value.casefold())


def _meaningful(value):
    return value is not None and (not isinstance(value, (str, bytes)) or bool(value.strip()))


def _metadata_in_image(image):
    if any(_meaningful(value) and _key(key) in IDENTITY_TAGS | MEDIA_PROFILE_KEYS
           for key, value in image.info.items()):
        return True
    exif = image.getexif()
    if any(_meaningful(exif.get(tag)) for tag in PRIVATE_EXIF):
        return True
    if 34665 in exif:
        return any(_meaningful(value) for tag, value in exif.get_ifd(34665).items()
                   if tag in PRIVATE_EXIF)
    return False


def _identity_in_xml(root, terms):
    for element in root.iter():
        local = _local(element.tag)
        if not local:
            continue
        key = local.casefold()
        if key in IDENTITY_TAGS and _meaningful(element.text):
            return "署名、机构或人员属性"
        if local in PERSON_PART_ROOTS and any(isinstance(item.tag, str) for item in element):
            return "评论署名或人员信息"
        if local == "property" and _key(element.get("name", "")) in IDENTITY_TAGS | IDENTITY_ATTRS:
            if _meaningful("".join(element.itertext())):
                return "自定义文档属性中的人员信息"
        if key in {"paragraph", "si", "is"} or element.tag == A + "p":
            combined = "".join(element.itertext())
            if has_identity(combined, terms):
                return "文本（含拆分的文字段）"
        if element.text and has_identity(element.text, terms):
            return "文本或属性"
        if local == "cNvPr":
            if any(LOGO_LABEL.search(element.get(name, "")) for name in ("name", "title", "descr")):
                return "徽标形状或图片说明"
        for name, value in element.attrib.items():
            attr = _local(name).casefold()
            if attr in IDENTITY_ATTRS and _meaningful(value):
                return "署名、机构或人员属性"
            if attr in {"name", "title", "descr", "description", "alt", "target", "val", "tooltip", "href", "text", "value"}:
                if has_identity(value, terms):
                    return "形状名称、替代文字或其他属性"
        if local == "Relationship" and element.get("TargetMode") == "External":
            relation_type = element.get("Type", "").rsplit("/", 1)[-1]
            if PRIVATE_TARGET.search(element.get("Target", "")) or relation_type in {"image", "oleObject", "package", "audio", "video"}:
                return "含外部文件、媒体或私人联系方式的链接"
    return None


def _audit_parts(deck, terms):
    for item in deck.infolist():
        if item.is_dir():
            continue
        part = item.filename
        suffix = posixpath.splitext(part)[1].casefold()
        if part.casefold().endswith(XML_SUFFIXES):
            try:
                reason = _identity_in_xml(_xml(deck.read(item)), terms)
            except etree.XMLSyntaxError:
                yield f"{part}: 无法读取 XML，不能完成匿名检查"
                continue
            if reason:
                yield f"{part}: {reason}含潜在身份信息，请检查并匿名化"
        is_media = "/media/" in "/" + part.casefold()
        if is_media or suffix in RASTER_SUFFIXES or suffix in BINARY_VECTOR_SUFFIXES:
            blob = deck.read(item)
            if sha256(blob).hexdigest() in BRANDED_IMAGES:
                yield f"{part}: 仍含原模板的徽标或具名机构图片"
            if suffix in RASTER_SUFFIXES:
                try:
                    with Image.open(io.BytesIO(blob)) as image:
                        # Some PNG text chunks can occur after IDAT. Load the image
                        # before reading metadata, keeping the stored bytes intact.
                        image.load()
                        if _metadata_in_image(image):
                            yield (f"{part}: 图片含署名、说明、版权、位置或自定义元数据；"
                                   "请无损清理隐私元数据后重新登记原图")
                except (OSError, ValueError, TypeError, SyntaxError, Image.DecompressionBombError):
                    yield f"{part}: 无法读取媒体或其元数据，不能完成匿名检查"
            elif suffix in BINARY_VECTOR_SUFFIXES:
                yield f"{part}: 无法审查二进制矢量图片，请改用可编辑图形或经检查的图片"
            elif suffix != ".svg":
                yield f"{part}: 未受支持的媒体类型，请先验证或移除其中的身份信息"
        if "/embeddings/" in "/" + part.casefold():
            # A native chart can have its own workbook. Audit all its parts too.
            try:
                with ZipFile(io.BytesIO(deck.read(item))) as embedded:
                    for problem in _audit_parts(embedded, terms):
                        yield f"{part} / {problem}"
            except BadZipFile:
                yield f"{part}: 无法核查嵌入附件，请移除或用可编辑原生对象重建"


def _relationships(deck, part):
    folder, filename = posixpath.split(part)
    path = posixpath.join(folder, "_rels", filename + ".rels")
    try:
        root = _xml(deck.read(path))
    except KeyError:
        return {}
    return {rel.get("Id"): rel for rel in root}


def _target(part, rel):
    if rel is None or rel.get("TargetMode") == "External":
        return None
    value = rel.get("Target", "")
    if not value:
        return None
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), value)).lstrip("/")


def _ordered_slides(deck):
    part = "ppt/presentation.xml"
    main = _xml(deck.read(part))
    rels = _relationships(deck, part)
    return [_target(part, rels.get(row.get(R + "id")))
            for row in main.findall(f"{P}sldIdLst/{P}sldId")]


def _shapes(root):
    result = {}
    for properties in root.iter(P + "cNvPr"):
        shape = properties.getparent().getparent()
        if shape.tag not in EXTENSION_SHAPES:
            continue
        number = int(properties.get("id", "0"))
        if number in result:
            raise ValueError(f"重复的形状 ID {number}")
        result[number] = shape
    return result


def _has_text(shape):
    text_body = shape.find(P + "txBody")
    return text_body is not None and any((text.text or "").strip() for text in text_body.iter(A + "t"))


def _editable_for_role(shape, role):
    if shape is None:
        return False
    if role in TEXT_ROLES:
        return shape.tag == P + "sp" and _has_text(shape)
    if role == "table":
        return shape.tag == P + "graphicFrame" and shape.find(f".//{A}tbl") is not None
    if role in {"shape", "component"}:
        if shape.tag == P + "grpSp":
            return any(child.tag in EDITABLE_TAGS for child in shape.iterdescendants())
        return shape.tag in EDITABLE_TAGS or (
            shape.tag == P + "graphicFrame"
            and (shape.find(f".//{A}tbl") is not None or shape.find(f".//{C}chart") is not None))
    return False


def _shape_text_lines(shape):
    paragraphs = shape.findall(f"{P}txBody/{A}p")
    for paragraph in paragraphs:
        parts = []
        for node in paragraph.iter():
            if node.tag == A + "t":
                parts.append(node.text or "")
            elif node.tag == A + "br":
                parts.append("\n")
        yield from (line.strip() for line in "".join(parts).splitlines() if line.strip())


def _notes_body_lines(deck, part, rels):
    relation = next((rel for rel in rels.values()
                     if rel.get("Type", "").endswith("/notesSlide")), None)
    notes_part = _target(part, relation)
    if not notes_part:
        return ()
    for shape in _shapes(_xml(deck.read(notes_part))).values():
        placeholder = shape.find(f"{P}nvSpPr/{P}nvPr/{P}ph")
        if placeholder is not None and placeholder.get("type") == "body":
            return tuple(_shape_text_lines(shape))
    return ()


def _has_visible_source(shapes, notes_sources):
    keys = {source_key(line) for line in notes_sources}
    for shape in shapes.values():
        properties = shape.find(f"*/{P}cNvPr")
        if properties is not None and properties.get("name") == SOURCE_SHAPE_NAME:
            return True
        if any(PREFIX.match(line) and source_key(line) in keys for line in _shape_text_lines(shape)):
            return True
    return False


def _provenance_valid(item):
    if not isinstance(item.get("source"), str) or not item["source"].strip():
        return False
    if "pdf_page" in item and (type(item["pdf_page"]) is not int or item["pdf_page"] < 1):
        return False
    if "panel" in item and (not isinstance(item["panel"], str) or not item["panel"].strip()):
        return False
    if "crop" in item:
        box = item["crop"]
        return (isinstance(box, list) and len(box) == 4
                and all(type(x) in (int, float) for x in box)
                and 0 <= box[0] < box[2] <= 1 and 0 <= box[1] < box[3] <= 1)
    return True


def _picture_problem(deck, slide_part, rels, shape, wanted):
    if shape is None or shape.tag != P + "pic":
        return "要求独立嵌入可选中的论文原图"
    blip = shape.find(f"{P}blipFill/{A}blip")
    embed = blip.get(R + "embed") if blip is not None else None
    relation = rels.get(embed)
    if relation is None or not relation.get("Type", "").endswith("/image"):
        return "缺少原图的内部图片关系"
    part = _target(slide_part, relation)
    if not part or part not in deck.namelist():
        return "缺少原图资源"
    blob = deck.read(part)
    if sha256(blob).hexdigest() != wanted:
        return "嵌入图片与登记的论文原图不一致"
    crop = shape.find(f"{P}blipFill/{A}srcRect")
    if crop is not None and any(int(crop.get(side, "0")) != 0 for side in ("l", "t", "r", "b")):
        return "PPTX 又裁切了已登记的原图，需核对图例、面板号和坐标轴后登记裁切文件"
    xf = shape.find(f"{P}spPr/{A}xfrm")
    if xf is None:
        return "缺少原图的放置尺寸"
    if any(xf.get(name) in {"1", "true"} for name in ("flipH", "flipV")):
        return "原图遭到镜像翻转"
    scale = xf.find(A + "ext")
    if scale is None:
        return "缺少原图的放置尺寸"
    with Image.open(io.BytesIO(blob)) as image:
        ratio = image.width / image.height
    cx, cy = (int(scale.get(dim, "0")) for dim in ("cx", "cy"))
    if cx <= 0 or cy <= 0 or not isclose(ratio, cx / cy, rel_tol=0.015):
        return "原图缩放改变了宽高比"
    return None


def _slide_issues(deck, part, number, entry):
    place = f"P{number}"
    declared = entry.get("shapes", [])
    evidence = entry.get("evidence", [])
    if not isinstance(declared, list) or not isinstance(evidence, list):
        yield f"{place}: shapes 和 evidence 应为列表"
        return
    shapes = _shapes(_xml(deck.read(part)))
    rels = _relationships(deck, part)
    has_source = False
    for item in declared:
        if not isinstance(item, dict) or type(item.get("shape_id")) is not int or not isinstance(item.get("role"), str):
            yield f"{place}: 缺少有效的可编辑槽位 ID 或角色"
            continue
        role, number_id = item["role"], item["shape_id"]
        if not _editable_for_role(shapes.get(number_id), role):
            yield f"{place} 对象 {number_id}: {role} 槽位缺失、无文字或已转成图片"
        elif role == "source":
            has_source = True
    if "notes_sources" in entry:
        expected = entry["notes_sources"]
        if not isinstance(expected, list) or any(
                not isinstance(line, str) or not line.strip() or len(line.splitlines()) != 1
                for line in expected):
            yield f"{place}: notes_sources 应为单行图源文字的列表"
            expected = []
        actual = set(_notes_body_lines(deck, part, rels))
        if (entry.get("source_required") and not expected) or set(expected) - actual:
            yield f"{place}: 演讲者备注缺少登记的图源说明"
        if has_source or _has_visible_source(shapes, expected):
            yield f"{place}: 幻灯片上仍有图源脚注"
        if entry.get("notes_required") and not actual.difference(expected):
            yield f"{place}: 缺少讲稿备注"
    else:
        # Older decks declare visible slide sources instead of citations in notes.
        if entry.get("source_required") and not has_source:
            yield f"{place}: 缺少可编辑图源说明"
        if entry.get("notes_required") and not _notes_body_lines(deck, part, rels):
            yield f"{place}: 缺少讲稿备注"
    for item in evidence:
        if not isinstance(item, dict) or type(item.get("shape_id")) is not int:
            yield f"{place}: 论文原图缺少有效的对象 ID"
            continue
        number_id = item["shape_id"]
        if not _provenance_valid(item):
            yield f"{place} 对象 {number_id}: 缺少原图出处，或页码、裁切范围、面板号无效"
        wanted = item.get("sha256")
        if not isinstance(wanted, str) or not SHA256.fullmatch(wanted):
            yield f"{place} 对象 {number_id}: 缺少原图 SHA-256"
            continue
        problem = _picture_problem(deck, part, rels, shapes.get(number_id), wanted)
        if problem:
            yield f"{place} 对象 {number_id}: {problem}"


def _manifest_identity_terms(manifest):
    if manifest is None:
        return ()
    if not isinstance(manifest, dict):
        raise ValueError("交付清单应为 JSON 对象")
    terms = manifest.get("identity_terms", [])
    if not isinstance(terms, list) or any(not isinstance(term, str) or not term.strip() for term in terms):
        raise ValueError("identity_terms 应为本地维护的非空身份词列表")
    return tuple(terms)


def audit_package(pptx, manifest=None):
    """Yield findings from the serialized package, including unreferenced parts.

    Manifest identity terms are local inputs. Diagnostics report locations without
    echoing those terms. Media hashes refer to the reviewed, cropped source bytes.
    """
    try:
        terms = _manifest_identity_terms(manifest)
        with ZipFile(pptx) as deck:
            if len(deck.namelist()) != len(set(deck.namelist())):
                yield "PPTX 含有重名部件，无法可靠验收"
                return
            yield from _audit_parts(deck, terms)
            if manifest is not None:
                records = manifest.get("slides")
                slides = _ordered_slides(deck)
                if not isinstance(records, dict) or set(records) != {str(i) for i in range(1, len(slides) + 1)}:
                    yield "交付清单与 PPTX 的页数或页序不一致"
                    return
                for i, part in enumerate(slides, 1):
                    entry = records[str(i)]
                    if not part or not isinstance(entry, dict):
                        yield f"P{i}: 幻灯片关系或交付检查记录无效"
                        continue
                    yield from _slide_issues(deck, part, i, entry)
    except (BadZipFile, etree.XMLSyntaxError, KeyError, OSError, ValueError, TypeError):
        yield "无法完整读取 PPTX 或其交付清单，请核对文件结构与清单字段"
