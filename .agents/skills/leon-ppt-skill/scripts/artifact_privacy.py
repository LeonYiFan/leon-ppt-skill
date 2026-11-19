"""Keep identity details and branded template artwork out of generated PPTX files."""
from __future__ import annotations

import re
from hashlib import sha256
from io import BytesIO
from unicodedata import normalize

from PIL import Image
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

import ppt_template as T

# These are the bundled template's branded images. Detect them by content so
# copying or renumbering PowerPoint relationships does not restore them.
BADGES = {
    "37c4491befe5a93f6bfff3472d2e7c121aba015e62d52db54f73e3220d8384b2",
    "f3df9907a59485495a365cdebb62c5fb8f6b8d142b6d72f331b9f74335c3b543",
}
BRANDED_PHOTOS = {
    "45f2f1d672f90cdc51755b27f4cc693d15bf71d1c3f2b98a12c91e1721d1b9ea",
    "4c714f73268242536b711d6886179835be56f40a066c6539e02b0410e003d8a2",
}
SECTION_PHOTO = "bb90bb2bb16d4036eccf68bf7e87edb2c7639ea2690602be196b28adb1e9499e"
THANKS_ART = "6c63f50508d6f2496d31a2d3648c6c5f1cea98e2283b508e38d47cf23171613c"
BRANDED_IMAGES = BADGES | BRANDED_PHOTOS | {SECTION_PHOTO, THANKS_ART}

TEMPLATE_IDENTITIES = ()
IDENTITY_LABEL = re.compile(
    r"(?:^|[\s，。；;（(])(?:汇报人|演讲者|第一作者|通讯作者|作者姓名|姓名|"
    r"作者|所属单位|作者单位|汇报单位|工作单位|机构名称|联系方式|电子邮箱|手机号码|"
    r"presented by|speaker|corresponding author|affiliation|full name|"
    r"contact email|e-?mail)\s*[:：]\s*\S",
    re.IGNORECASE,
)
EMAIL = re.compile(r"\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b", re.IGNORECASE)
PHONE = re.compile(r"(?<![\d.])(?:\+86[-\s]+)?1[3-9]\d{9}(?![\d.])")

IDENTITY_SHAPES = {"T_Meta", "T_Authors", "K_Meta"}
LOGO_LABEL = re.compile(r"\blogo\b|徽标|院徽|校徽|单位标识", re.IGNORECASE)
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PACKAGE_PROPERTIES = {"thumbnail", "extended-properties", "custom-properties"}
CORE_PROPERTIES = (
    "title", "subject", "author", "last_modified_by", "keywords", "comments",
    "category", "identifier", "language", "content_status", "version",
)


def _section_artwork(shapes, picture):
    """Use the original divider picture's box for an anonymous blue chevron."""
    for name, color, points in (
        ("章节浅蓝边", "B9D6F4", [(480, 0), (925, 435), (975, 500), (925, 565), (420, 1000), (0, 1000)]),
        ("章节深蓝底", "386497", [(435, 0), (870, 439), (920, 500), (870, 563), (365, 1000), (0, 1000)]),
        ("章节蓝色叠影", "285284", [(663, 221), (870, 439), (920, 500), (870, 563), (365, 1000), (0, 1000),
                                   (0, 650), (350, 470)]),
    ):
        freeform = shapes.build_freeform(0, 0, scale=(picture.width / 1000, picture.height / 1000))
        freeform.add_line_segments(points, close=True)
        shape = freeform.convert_to_shape(picture.left, picture.top)
        shape.name = name
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(color)
        shape.line.fill.background()
        picture._element.addprevious(shape._element)


def _crop_thanks_art(picture):
    """Embed only the anonymous visible ribbon, leaving the original art on disk."""
    with Image.open(BytesIO(picture.image.blob)) as image:
        width, height = image.size
        # The two seals and institution name sit below the original visible crop.
        right = 1 - picture.crop_right
        bottom = min(1 - picture.crop_bottom, 0.70)
        box = (round(width * picture.crop_left), round(height * picture.crop_top),
               round(width * right), round(height * bottom))
        payload = BytesIO()
        image.convert("RGB").crop(box).save(payload, format="PNG")
    payload.seek(0)
    _, rid = picture.part.get_or_add_image_part(payload)
    picture._element.blipFill.blip.set(qn("r:embed"), rid)
    picture.crop_left = picture.crop_top = picture.crop_right = picture.crop_bottom = 0


def _clean_shapes(shapes):
    for shape in list(shapes):
        details = next(shape._element.iter(qn("p:cNvPr")), None)
        labels = "" if details is None else " ".join(
            details.get(attr, "") for attr in ("name", "title", "descr"))
        if shape.name in IDENTITY_SHAPES or LOGO_LABEL.search(labels):
            T.drop(shape)
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            _clean_shapes(shape.shapes)
            if not len(shape.shapes):
                T.drop(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            if shape._element.blipFill.blip.get(qn("r:embed")) is None:
                continue
            digest = sha256(shape.image.blob).hexdigest()
            if digest in BADGES or digest in BRANDED_PHOTOS:
                T.drop(shape)
            elif digest == SECTION_PHOTO:
                _section_artwork(shapes, shape)
                T.drop(shape)
            elif digest == THANKS_ART:
                _crop_thanks_art(shape)


def _drop_unused_images(part):
    refs = {value for element in part._element.iter()
            for name, value in element.attrib.items() if name.startswith(REL_NS)}
    for rid, rel in list(part.rels.items()):
        if rel.reltype == RT.IMAGE and rid not in refs:
            part.drop_rel(rid)


def clean_template(presentation, title):
    """Apply privacy changes to the in-memory template before cloning pages."""
    canvases = list(presentation.slides)
    for master in presentation.slide_masters:
        canvases.append(master)
        canvases.extend(master.slide_layouts)
    for canvas in canvases:
        _clean_shapes(canvas.shapes)
        _drop_unused_images(canvas.part)

    package = presentation.part.package
    for rid, rel in list(package._rels.items()):
        if rel.reltype.rsplit("/", 1)[-1] in PACKAGE_PROPERTIES:
            package.drop_rel(rid)
    core = presentation.core_properties
    for element in list(core._element):
        core._element.remove(element)
    core.title = title
    core.language = "zh-CN"


def slide_text(slide):
    """Yield text with locations for editable shapes, table cells and notes."""
    for shape in T.iter_shapes(slide.shapes):
        if shape.has_text_frame:
            yield shape.name, shape.text_frame.text
        if shape.has_table:
            for row, cells in enumerate(shape.table.rows, 1):
                for col, cell in enumerate(cells.cells, 1):
                    if not cell.is_spanned:
                        yield f"{shape.name} R{row}C{col}", cell.text
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame
        if notes is not None:
            yield "讲稿备注", notes.text


def core_text(presentation):
    for field in CORE_PROPERTIES:
        value = getattr(presentation.core_properties, field)
        if value.strip():
            yield field, value


def has_identity(text, identity_terms=()):
    text = normalize("NFKC", text)
    folded = " ".join(text.casefold().split())
    for term in (*TEMPLATE_IDENTITIES, *identity_terms):
        key = " ".join(normalize("NFKC", term).casefold().split())
        if key.isascii():
            if re.search(r"(?<![a-z0-9])" + re.escape(key) + r"(?![a-z0-9])", folded):
                return True
        elif key in folded:
            return True
    return bool(IDENTITY_LABEL.search(text) or EMAIL.search(text) or PHONE.search(text))


def identity_issues(presentation, identity_terms=()):
    for index, slide in enumerate(presentation.slides, 1):
        for slot, text in slide_text(slide):
            if has_identity(text, identity_terms):
                yield f"P{index} {slot}: 包含姓名、单位或联系方式，请匿名化"
    for field, value in core_text(presentation):
        if field in ("author", "last_modified_by") or has_identity(value, identity_terms):
            yield f"文档属性 {field}: 含身份信息"
    for part in presentation.part.package.iter_parts():
        if str(part.partname).startswith("/ppt/media/"):
            if sha256(part.blob).hexdigest() in BRANDED_IMAGES:
                yield f"{part.partname}: 包含原模板的徽标或具名机构图片"
