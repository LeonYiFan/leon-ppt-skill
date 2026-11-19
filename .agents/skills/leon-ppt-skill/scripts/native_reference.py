"""Import the editable geometry of a local PPTX slide without its content."""

from copy import deepcopy
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from posixpath import dirname, join, normpath
from zipfile import ZipFile
import json
import re

from lxml import etree
from PIL import Image
from pptx.oxml.ns import qn

from ppt_template import CM, RED, _make_run, segments


SOURCES = ("gut", "guide", "kw16", "kw16-9")
_WASH_SHA256 = "17e63fc3fe805c05d48f0c6442cbce22f12033671c773b42b1506ce761058511"
_LAYOUT_CHROME = {
    ("gut", "ppt/slideLayouts/slideLayout1.xml"): ("14", "17", "15", "30"),
    ("gut", "ppt/slideLayouts/slideLayout2.xml"): ("14", "17"),
    ("guide", "ppt/slideLayouts/slideLayout1.xml"): ("14", "17", "15", "30"),
    ("guide", "ppt/slideLayouts/slideLayout2.xml"): ("14", "17"),
}
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
RENDER_NS = (P, A)
STRUCTURAL = {"sp", "pic", "grpSp", "cxnSp", "graphicFrame"}
FILL_NAMES = ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")
PROP_RANK = {name: i for i, names in enumerate((
    ("xfrm",), ("prstGeom", "custGeom"), FILL_NAMES, ("ln",),
    ("effectLst", "effectDag"), ("scene3d",), ("sp3d",),
)) for name in names}
BRANDING = re.compile(r"\blogo\b|\bwatermark\b|商标|徽标|水印|标识|标志", re.IGNORECASE)
REMOVE_TAGS = {
    A+"hlinkClick", A+"hlinkHover", A+"blipFill", A+"blip", A+"buBlip",
    A+"audioFile", A+"videoFile", A+"wavAudioFile", P+"blipFill", P+"oleObj", P+"controls",
    P+"custDataLst", P+"tags", A+"extLst", P+"extLst",
}
_TEXT_DEFAULT = etree.fromstring(f'<a:rPr xmlns:a="{A[1:-1]}"/>'.encode())


def _tag(el):
    return etree.QName(el).localname


def _xml(z, part):
    return etree.fromstring(z.read(part), parser=etree.XMLParser(resolve_entities=False, no_network=True))


def _rels(z, part):
    file = join(dirname(part), "_rels", part.rsplit("/", 1)[-1] + ".rels")
    try:
        root = _xml(z, file)
    except KeyError:
        return {}
    return {
        e.get("Id"): (e.get("Type").rsplit("/", 1)[-1],
                      normpath(e.get("Target").lstrip("/") if e.get("Target").startswith("/")
                               else join(dirname(part), e.get("Target"))))
        for e in root if e.get("TargetMode") != "External"
    }


def _related(z, part, kind):
    return next((target for rel, target in _rels(z, part).values() if rel == kind), None) if part else None


def _first(parent, nsname):
    return parent.find(nsname) if parent is not None else None


def _insert_property(parent, el):
    rank = PROP_RANK.get(_tag(el), float("inf"))
    index = next((i for i, c in enumerate(parent)
                  if PROP_RANK.get(_tag(c), float("inf")) > rank), len(parent))
    parent.insert(index, el)


def _color_as_rgb(scheme):
    values = {}
    if scheme is None:
        return values
    for color in scheme:
        el = next(iter(color), None)
        if el is not None:
            val = el.get("lastClr") if _tag(el) == "sysClr" else el.get("val")
            if val and re.fullmatch(r"[0-9A-Fa-f]{6}", val):
                values[_tag(color)] = val
    return values


class _Theme:
    def __init__(self, slide, layout, master, theme):
        te = _first(theme, A+"themeElements")
        self.palette = _color_as_rgb(_first(te, A+"clrScheme"))
        self.color_map = dict(bg1="lt1", tx1="dk1", bg2="lt2", tx2="dk2")
        mapping = _first(master, P+"clrMap")
        if mapping is not None:
            self.color_map.update(mapping.attrib)
        for layer in (layout, slide):
            mapping = _first(_first(layer, P+"clrMapOvr"), A+"overrideClrMapping")
            if mapping is not None:
                self.color_map.update(mapping.attrib)
        self.fmt = _first(te, A+"fmtScheme")
        font = _first(te, A+"fontScheme")
        self.fonts = {}
        for prefix, branch in (("+mj", "majorFont"), ("+mn", "minorFont")):
            familial = _first(font, A+branch)
            for language in ("latin", "ea", "cs"):
                entry = _first(familial, A+language)
                typeface = entry.get("typeface") if entry is not None else ""
                if not typeface and language == "ea" and familial is not None:
                    hans = next((e for e in familial.findall(A+"font") if e.get("script") == "Hans"), None)
                    typeface = hans.get("typeface") if hans is not None else ""
                suffix = {"latin":"lt", "ea":"ea", "cs":"cs"}[language]
                if typeface:
                    self.fonts[prefix + "-" + suffix] = typeface

    def materialize(self, node, ref_color=None):
        for el in list(node.iter(A+"schemeClr")):
            value = el.get("val")
            if value == "phClr" and ref_color is not None:
                copied = deepcopy(ref_color)
                copied.extend(deepcopy(list(el)))
                el.getparent().replace(el, copied)
            elif value != "phClr":
                rgb = self.palette.get(self.color_map.get(value, value))
                if rgb:
                    el.tag = A+"srgbClr"
                    el.set("val", rgb)
        for el in node.iter():
            typ = el.get("typeface")
            if typ in self.fonts:
                el.set("typeface", self.fonts[typ])

    def style_element(self, kind, ref):
        if ref is None or self.fmt is None:
            return None
        lists = {"fill": "fillStyleLst", "line": "lnStyleLst", "effect": "effectStyleLst"}
        index = int(ref.get("idx", 0))
        sequence = _first(self.fmt, A+("bgFillStyleLst" if kind == "fill" and index >= 1001 else lists[kind]))
        if index >= 1001:
            index -= 1000
        if sequence is None or index < 1 or index > len(sequence):
            return None
        style = deepcopy(sequence[index-1])
        ref_color = deepcopy(next(iter(ref), None))
        if ref_color is not None:
            self.materialize(ref_color)
            value = ref_color.get("val")
            if ref_color.tag == A+"schemeClr":
                rgb = self.palette.get(self.color_map.get(value, value))
                if rgb:
                    ref_color.tag = A+"srgbClr"
                    ref_color.set("val", rgb)
        if kind == "effect":
            child = next((x for x in style if _tag(x) in ("effectLst", "effectDag")), None)
            if child is None:
                return None
            style = deepcopy(child)
        self.materialize(style, ref_color)
        return style

    def implicit_text_style(self, node):
        """Source fontRef controls the default for an unstyled run."""
        source = _first(_first(node, P+"style"), A+"fontRef")
        run = deepcopy(_TEXT_DEFAULT)
        if source is not None:
            prefix = "+mj" if source.get("idx") == "major" else "+mn"
            color = next(iter(source), None)
            if color is not None:
                fill = etree.SubElement(run, A+"solidFill")
                fill.append(deepcopy(color))
        else:
            prefix = "+mn"
        for name, suffix in (("latin", "lt"), ("ea", "ea"), ("cs", "cs")):
            font = self.fonts.get(prefix+"-"+suffix)
            if font:
                etree.SubElement(run, A+name, typeface=font)
        self.materialize(run)
        return run

    def shape_properties(self, node):
        style = _first(node, P+"style")
        sppr = _first(node, P+"spPr")
        if style is None or sppr is None:
            return
        for field, ref, names in (
            ("fill", "fillRef", FILL_NAMES), ("line", "lnRef", ("ln",)),
            ("effect", "effectRef", ("effectLst", "effectDag")),
        ):
            if any(_tag(e) in names for e in sppr):
                continue
            value = self.style_element(field, _first(style, A+ref))
            if value is not None and _tag(value) != "blipFill":
                _insert_property(sppr, value)
        # Keep the style as fallback for any omitted materialization, but resolve values.
        self.materialize(style)


def _cnv(node):
    return next((n.find(P+"cNvPr") for n in node if n.tag.startswith(P) and _tag(n).startswith("nv")), None)


def _shape_id(node):
    props = _cnv(node)
    return props.get("id") if props is not None else None


def _shape_type(el):
    kind = _tag(el)
    if kind == "sp":
        return "text" if el.find(P+"txBody") is not None else "shape"
    if kind == "graphicFrame":
        if el.find(f".//{A}tbl") is not None:
            return "table"
        data = el.find(f".//{A}graphicData")
        return "chart" if data is not None and data.get("uri", "").endswith("/chart") else "unsupported"
    return {"pic":"picture", "grpSp":"group", "cxnSp":"connector"}.get(kind, "unsupported")


def _xfrm(node):
    for path in (P+"spPr/"+A+"xfrm", P+"grpSpPr/"+A+"xfrm", P+"xfrm"):
        xf = node.find(path)
        if xf is not None:
            return xf
    return None


def _xy(xfrm, tag, x="x", y="y"):
    pt = _first(xfrm, A+tag)
    return None if pt is None else (int(pt.get(x, 0)), int(pt.get(y, 0)))


def _placeholder(el):
    return next(el.iter(P+"ph"), None)


def _matching_placeholder(el, layer):
    source = _placeholder(el)
    if source is None or layer is None:
        return None
    sid = source.get("idx", "0")
    kind = source.get("type", "obj")
    fallbacks = []
    tree = layer.find(P+"cSld/"+P+"spTree")
    if tree is None:
        return None
    for other in tree.findall(P+"sp"):
        ph = _placeholder(other)
        if ph is not None:
            if ph.get("idx", "0") == sid and ph.get("type", "obj") == kind:
                return other
            if ph.get("idx", "0") == sid or ph.get("type", "obj") == kind:
                fallbacks.append(other)
    return next(iter(fallbacks), None)


def _effective_xfrm(el, page):
    direct = _xfrm(el)
    if direct is not None:
        return direct
    for layer in (page.layout, page.master):
        parent = _matching_placeholder(el, layer)
        if parent is not None:
            inherited = _xfrm(parent)
            if inherited is not None:
                return inherited
    return None


def _geometry(el, page, mapping):
    xf = _effective_xfrm(el, page)
    corner, size = _xy(xf, "off"), _xy(xf, "ext", "cx", "cy")
    if corner is None or size is None:
        return None, mapping
    x, y = corner
    w, h = size
    sx, sy, dx, dy = mapping
    box = [round(v/CM, 4) for v in (x*sx+dx, y*sy+dy, w*sx, h*sy)]
    if _tag(el) == "grpSp":
        child_corner = _xy(xf, "chOff")
        child_size = _xy(xf, "chExt", "cx", "cy")
        if child_corner and child_size and all(child_size):
            tx, ty = w / child_size[0] * sx, h / child_size[1] * sy
            mapping = (tx, ty, x*sx+dx-child_corner[0]*tx, y*sy+dy-child_corner[1]*ty)
    return box, mapping


def _slot_art(rels, node):
    blip = node.find(".//"+A+"blip")
    rid = blip.get(R+"embed") if blip is not None else None
    kind, member = rels.get(rid, (None, None))
    return member if kind == "image" else None


class _SourcePage:
    def __init__(self, z, source, page_num, order, size):
        self.z = z
        self.source = source
        self.number = page_num
        self.part = order[page_num-1]
        self.width, self.height = size
        self.slide = _xml(z, self.part)
        self.rels = _rels(z, self.part)
        self.tree = self.slide.find(P+"cSld/"+P+"spTree")
        self.layout_part = next((target for rel, target in self.rels.values() if rel == "slideLayout"), None)
        self.layout_rels = _rels(z, self.layout_part) if self.layout_part else {}
        master = next((target for rel, target in self.layout_rels.values() if rel == "slideMaster"), None)
        theme = _related(z, master, "theme")
        self.layout = _xml(z, self.layout_part) if self.layout_part else None
        self.master = _xml(z, master) if master else None
        self.theme = _Theme(self.slide, self.layout, self.master, _xml(z, theme) if theme else None)

    def background(self):
        for layer in (self.slide, self.layout, self.master):
            bg = _first(_first(layer, P+"cSld"), P+"bg")
            if bg is None:
                continue
            bg = deepcopy(bg)
            ref = bg.find(P+"bgRef")
            if ref is not None:
                fill = self.theme.style_element("fill", ref)
                if fill is None or fill.tag == A+"blipFill":
                    break
                bg.remove(ref)
                etree.SubElement(bg, P+"bgPr").append(fill)
            if any(x in (A+"blipFill", P+"blipFill") for x in (e.tag for e in bg.iter())):
                break
            self.theme.materialize(bg)
            _scrub(bg)
            return bg
        tag = etree.Element(P+"bg")
        etree.SubElement(etree.SubElement(etree.SubElement(tag, P+"bgPr"), A+"solidFill"),
                         A+"srgbClr", val="FFFFFF")
        return tag


def _scrub(node):
    for el in list(node.iter()):
        parent = el.getparent()
        if el is not node and (el.tag in REMOVE_TAGS or not el.tag.startswith(RENDER_NS)):
            if parent is not None:
                parent.remove(el)
            continue
        for name in list(el.attrib):
            if name.startswith(R):
                del el.attrib[name]
        if el.tag == P+"cNvPr":
            id_ = el.get("id", "0")
            el.attrib.clear()
            el.set("id", id_)
            el.set("name", "Reference object " + id_)
        if el.tag == P+"nvPr":
            for child in list(el):
                el.remove(child)


def _paragraph_style(par, fallback):
    for run in par.findall(A+"r"):
        candidate = run.find(A+"rPr")
        if candidate is not None and (run.findtext(A+"t") or "").strip():
            return candidate
    prop = par.find(A+"pPr/"+A+"defRPr")
    if prop is not None:
        return prop
    return fallback


def _replace_text(body, value, fallback):
    source = body.findall(A+"p")
    source = source or [etree.Element(A+"p")]
    for par in body.findall(A+"p"):
        body.remove(par)
    for i, line in enumerate(value.split("\n")):
        donor = source[min(i, len(source)-1)]
        par = etree.SubElement(body, A+"p")
        properties = donor.find(A+"pPr")
        if properties is not None:
            par.append(deepcopy(properties))
        if line:
            inherited = _paragraph_style(donor, _paragraph_style(source[0], fallback))
            for text, red in segments(line):
                par.append(_make_run(text, inherited, RED if red else None))
        ending = donor.find(A+"endParaRPr")
        if ending is not None:
            par.append(deepcopy(ending))


def _set_xfrm(new, xf):
    props = new.find(P+"spPr")
    old = props.find(A+"xfrm")
    if old is not None:
        props.replace(old, deepcopy(xf))
    else:
        props.insert(0, deepcopy(xf))


def _set_bg(slide, page):
    csld = slide._element.find(P+"cSld")
    prior = csld.find(P+"bg")
    if prior is not None:
        csld.remove(prior)
    csld.insert(0, page.background())
    slide._element.set("showMasterSp", "0")


def _node_map(tree):
    for el in tree:
        if _tag(el) in STRUCTURAL and el.tag.startswith(P):
            ident = _shape_id(el)
            yield ident, el
            if el.tag == P+"grpSp":
                yield from _node_map(el)


class NativeLibrary:
    def __init__(self, root):
        self.root = Path(root)
        self._index = {}
        self._approved = None
        self._hashes = {}

    def _open(self, source, page):
        if source not in SOURCES or type(page) is not int or page < 1:
            raise ValueError("native reference requires a source alias and one-based page")
        z = ZipFile(self.root / "native" / (source+".pptx"))
        try:
            if source not in self._index:
                pres = _xml(z, "ppt/presentation.xml")
                rel = _rels(z, "ppt/presentation.xml")
                order = [rel[e.get(R+"id")][1] for e in pres.find(P+"sldIdLst")]
                size = pres.find(P+"sldSz")
                self._index[source] = order, (int(size.get("cx")), int(size.get("cy")))
            order, size = self._index[source]
            if page > len(order):
                raise ValueError(f"native reference {source} has {len(order)} pages, got {page}")
            return _SourcePage(z, source, page, order, size)
        except BaseException:
            z.close()
            raise

    def inspect(self, source, page):
        item = self._open(source, page)
        with item.z:
            return {"source":source, "page":page, "width_emu":item.width, "height_emu":item.height,
                    "shapes":list(self._inspect_nodes(item.tree, item, None, (1,1,0,0)))}

    def _inspect_nodes(self, tree, page, parent, mapping):
        for node in tree:
            if not node.tag.startswith(P) or _tag(node) not in STRUCTURAL:
                continue
            cnv = _cnv(node)
            ident = _shape_id(node)
            typ = _shape_type(node)
            box, child = _geometry(node, page, mapping)
            xf = _effective_xfrm(node, page)
            entry = {"id":ident, "type":typ, "box_cm":box,
                     "rotation":int(xf.get("rot", "0")) / 60000 if xf is not None else 0,
                     "name":cnv.get("name",""), "parent":parent}
            if typ == "text":
                entry["text"] = "\n".join("".join(el.text or "" for el in p.iter(A+"t"))
                                            for p in node.find(P+"txBody").findall(A+"p"))
            if typ == "table":
                table = node.find(".//"+A+"tbl")
                entry["rows"] = len(table.findall(A+"tr"))
                entry["columns"] = len(table.find(A+"tblGrid"))
            if typ == "picture":
                member = _slot_art(page.rels, node)
                if member:
                    entry["member"] = member
                    key = page.z.filename, member
                    if key not in self._hashes:
                        self._hashes[key] = sha256(page.z.read(member)).hexdigest()
                    entry["sha256"] = self._hashes[key]
            yield entry
            if typ == "group":
                yield from self._inspect_nodes(node, page, ident, child)

    def _approvals(self):
        if self._approved is None:
            allow = self.root / "approved-artwork.json"
            self._approved = set(json.loads(allow.read_text(encoding="utf-8-sig"))["sha256"]) if allow.exists() else set()
        return self._approved

    def _approved_art(self, page, node, rels=None):
        member = _slot_art(page.rels if rels is None else rels, node)
        if member is None:
            raise ValueError(f"reference image {_shape_id(node)} has no permitted embedded raster")
        blob = page.z.read(member)
        if sha256(blob).hexdigest() not in self._approvals():
            raise ValueError(f"reference image {_shape_id(node)} is absent from the local artwork approvals")
        return blob

    def apply(self, slide, spec, figure_path):
        """Fill source geometry with anonymous task content. All IDs in the spec are source IDs."""
        permit = {"source", "page", "text", "tables", "images", "omit", "include", "artwork"}
        if not isinstance(spec, dict) or spec.keys() - permit:
            raise ValueError("unknown native reference field")
        page = self._open(spec.get("source"), spec.get("page"))
        with page.z:
            return self._apply_to_slide(slide, spec, figure_path, page)

    @staticmethod
    def _field(spec, key, container):
        value = spec.get(key, container())
        ids = value.keys() if isinstance(value, dict) else value
        if not isinstance(value, container) or any(not isinstance(i, str) for i in ids):
            raise ValueError(f"native reference {key} needs source shape IDs as strings")
        if isinstance(value, list) and len(set(value)) != len(value):
            raise ValueError(f"native reference {key} contains duplicate shape IDs")
        return value

    def _apply_to_slide(self, slide, spec, figure_path, page):
        tree = page.tree
        entries = dict(_node_map(tree))
        roots = dict((i, node) for i, node in _node_map(tree) if node.getparent() is tree)
        fields = {k:self._field(spec, k, dict) for k in ("text","tables","images")}
        omit, artwork = (set(self._field(spec,k,list)) for k in ("omit","artwork"))
        include = set(self._field(spec, "include", list)) if "include" in spec else set(roots)
        if not include <= roots.keys() or not include:
            raise ValueError("native include needs one or more top-level source shape IDs")
        chosen = {key for key, node in entries.items()
                  if key in include or any(_shape_id(e) in include for e in node.iterancestors())}
        for key, values in (*fields.items(), ("omit", omit), ("artwork", artwork)):
            unavailable = values.keys() - chosen if isinstance(values, dict) else values - chosen
            if unavailable:
                raise ValueError(f"native {key} unavailable shape IDs: {', '.join(sorted(unavailable))}")
        specified = set().union(*fields.values(), artwork)
        if specified & omit or artwork & fields["images"].keys():
            raise ValueError("cannot fill, embed or approve the same ID twice or while omitting it")
        for key in specified:
            if any(_shape_id(e) in omit for e in entries[key].iterancestors()):
                raise ValueError("cannot fill a shape in an omitted group")
        for key, allowed in (("text", {"text"}), ("tables", {"table"}), ("images", {"picture", "chart", "unsupported"})):
            if any(_shape_type(entries[i]) not in allowed for i in fields[key]):
                raise ValueError(f"native {key} supplied for a different shape type")
        if any(_shape_type(entries[i])!="picture" for i in artwork):
            raise ValueError("only a reviewed embedded picture can be used as artwork")
        if any(not isinstance(text, str) for text in (*fields["text"].values(), *fields["images"].values())):
            raise ValueError("native text and image paths must be strings")
        if "include" not in spec:
            if len(slide.shapes):
                raise ValueError("full native reference requires a blank slide")
            _set_bg(slide, page)
            self._adopt(slide, self._inherited_chrome(slide, page))

        cloned = []
        for ident, node in roots.items():
            if ident in include:
                result = self._clone(slide, node, fields, omit, artwork, figure_path, page)
                if result is not None:
                    cloned.append(result)
        return self._adopt(slide, cloned)

    @staticmethod
    def _adopt(slide, cloned):
        target = slide.shapes._spTree
        next_id = max((int(n.get("id")) for n in target.iter(P+"cNvPr")), default=1)+1
        remap = {}
        for block in cloned:
            for nv in block.iter(P+"cNvPr"):
                remap[nv.get("id")] = str(next_id)
                nv.set("id", str(next_id))
                next_id += 1
        for block in cloned:
            for el in list(block.iter()):
                if el.tag in (A+"stCxn", A+"endCxn"):
                    new = remap.get(el.get("id"))
                    if new is not None:
                        el.set("id", new)
                    else:
                        el.getparent().remove(el)
            target.insert_element_before(block, "p:extLst")
        # The python-pptx index cache can have been used for transient pictures.
        if hasattr(slide.shapes, "_cached_max_shape_id"):
            slide.shapes._cached_max_shape_id = None
        return {original: int(current) for original, current in remap.items()}

    def _inherited_chrome(self, slide, page):
        if page.slide.get("showMasterSp") == "0":
            return []
        ids = _LAYOUT_CHROME.get((page.source, page.layout_part), ())
        if not ids:
            return []
        tree = page.layout.find(P+"cSld/"+P+"spTree")
        output = []
        for node in tree:
            if _shape_id(node) not in ids:
                continue
            typ = _shape_type(node)
            if typ == "picture" and _WASH_SHA256 in self._approvals():
                part = _slot_art(page.layout_rels, node)
                if part and sha256(page.z.read(part)).hexdigest() == _WASH_SHA256:
                    output.append(self._copy_art(slide, page, node, page.layout_rels))
            if typ in ("shape", "text", "connector") and not any(t.text for t in node.iter(A+"t")):
                copied = deepcopy(node)
                page.theme.shape_properties(copied)
                page.theme.materialize(copied)
                _scrub(copied)
                output.append(copied)
        return output

    def _clone(self, slide, node, fields, omit, artwork, figure_path, page):
        ident = _shape_id(node)
        typ = _shape_type(node)
        if ident in omit:
            return None
        caption = " ".join(_cnv(node).get(key,"") for key in ("name","title","descr"))
        if BRANDING.search(caption) and ident not in artwork and ident not in fields["images"]:
            return None
        if typ in ("picture","chart","unsupported"):
            if ident in fields["images"]:
                return self._new_picture(slide, page, node, figure_path(fields["images"][ident]))
            if ident in artwork:
                return self._copy_art(slide, page, node)
            return None
        clone = deepcopy(node)
        if typ == "group":
            for old in list(clone):
                if _tag(old) in STRUCTURAL and old.tag.startswith(P):
                    clone.remove(old)
            for nested in node:
                if _tag(nested) in STRUCTURAL and nested.tag.startswith(P):
                    item = self._clone(slide, nested, fields, omit, artwork, figure_path, page)
                    if item is not None:
                        clone.append(item)
            if not any(e.tag.startswith(P) and _tag(e) in STRUCTURAL for e in clone):
                return None
            # Children already scrubbed; scrubbing the whole group would remove the new image relations.
            for prop in clone[:2]:
                _scrub(prop)
            page.theme.materialize(clone[:2][0])
            page.theme.materialize(clone[:2][1])
            return clone

        if _xfrm(clone) is None and typ in ("text","shape","connector"):
            xf = _effective_xfrm(node, page)
            if xf is not None:
                _set_xfrm(clone, xf)
        fallback = page.theme.implicit_text_style(node)
        if typ == "text":
            _replace_text(clone.find(P+"txBody"), fields["text"].get(ident,""), fallback)
        if typ == "table":
            table = clone.find(".//"+A+"tbl")
            table_rows = table.findall(A+"tr")
            values = fields["tables"].get(ident)
            if values is not None and (
                not isinstance(values, list) or len(values) != len(table_rows)
                or any(not isinstance(row, list) or len(row)!=len(el.findall(A+"tc"))
                       or not all(isinstance(cell,str) for cell in row) for row,el in zip(values,table_rows))
            ):
                raise ValueError(f"native table {ident} needs a matching rectangular list of text rows")
            for r, row in enumerate(table_rows):
                for c, cell in enumerate(row.findall(A+"tc")):
                    value = values[r][c] if values is not None else ""
                    body = cell.find(A+"txBody")
                    if body is None and value:
                        body = etree.Element(A+"txBody")
                        etree.SubElement(body, A+"bodyPr")
                        etree.SubElement(body, A+"lstStyle")
                        cell.insert(0, body)
                    if body is not None:
                        _replace_text(body, value, fallback)
        page.theme.shape_properties(clone)
        page.theme.materialize(clone)
        _scrub(clone)
        return clone

    def _new_picture(self, slide, page, node, path):
        xf = _effective_xfrm(node, page)
        offset, extent = _xy(xf, "off"), _xy(xf, "ext", "cx", "cy")
        if offset is None or extent is None or min(extent) <= 0:
            raise ValueError(f"reference image slot {_shape_id(node)} has no usable rectangle")
        with Image.open(path) as im:
            iw, ih = im.size
        sx = min(extent[0]/iw, extent[1]/ih)
        nw, nh = round(iw*sx), round(ih*sx)
        x = offset[0] + (extent[0]-nw)//2
        y = offset[1] + (extent[1]-nh)//2
        picture = slide.shapes.add_picture(path, x, y, nw, nh)
        el = picture._element
        el.getparent().remove(el)
        nv = _cnv(el)
        nv.attrib.clear()
        nv.set("id", _shape_id(node))
        nv.set("name", "Reference image " + _shape_id(node))
        current = _xfrm(el)
        for name in ("rot", "flipH", "flipV"):
            if xf.get(name) is not None:
                current.set(name, xf.get(name))
        return el

    def _copy_art(self, slide, page, node, rels=None):
        blob = self._approved_art(page, node, rels)
        _, rid = slide.part.get_or_add_image_part(BytesIO(blob))
        clone = deepcopy(node)
        picture_fill = clone.find(P+"blipFill")
        clone.remove(picture_fill)
        _scrub(clone)
        blip = picture_fill.find(A+"blip")
        for child in list(blip):
            if child.tag in REMOVE_TAGS or not child.tag.startswith(A):
                blip.remove(child)
            else:
                _scrub(child)
        blip.attrib.clear()
        blip.set(R+"embed", rid)
        clone.insert(1, picture_fill)
        page.theme.materialize(clone)
        return clone
