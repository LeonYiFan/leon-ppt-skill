"""Build and search local PowerPoint slot maps."""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from contextlib import closing
import json
import os
from pathlib import Path
import sqlite3
from tempfile import NamedTemporaryFile


ARCHIVE = Path(__file__).resolve().parent.parent / "assets" / "native-library"
CATALOG = "native-pages.jsonl"
SOURCE_ORDER = {"gut": 0, "guide": 1, "kw16": 2, "kw16-9": 3}
FILLABLE = {"text", "picture", "table"}
TYPE_LABELS = {
    "text": "文本框 文字 text",
    "picture": "图片 原图 图框 picture",
    "table": "原生表格 表格 table",
    "shape": "原生形状 几何 shape",
    "connector": "连接线 箭头 connector",
    "group": "组合形状 group",
    "chart": "图表 chart",
    "unsupported": "其他对象",
}
SLOT_FIELDS = ("id", "type", "parent", "box_cm", "rows", "columns")
PAGE_FIELDS = (
    "source", "page", "label", "role", "priority", "recipe", "profile",
    "member", "native_path", "preview_path", "shape_count", "geometry_sha256", "hidden",
)


def iter_jsonl(path):
    with path.open(encoding="utf-8-sig") as stream:
        yield from (json.loads(line) for line in stream if line.strip())


def load_material_links(root):
    links = defaultdict(list)
    slots = defaultdict(list)
    uri = (root / "素材索引.sqlite3").resolve().as_uri() + "?mode=ro"
    with closing(sqlite3.connect(uri, uri=True)) as archive:
        for source, page, member, file, category, digest in archive.execute("""
            SELECT sa.source, sa.page, sa.member, a.archived_path, a.category, a.sha256
            FROM slide_assets sa JOIN assets a ON a.sha256 = sa.asset_sha256
            ORDER BY sa.source, sa.page, sa.member
        """):
            links[(source, page)].append({
                "member": member,
                "file": file,
                "category": category,
                "sha256": digest,
            })
        for source, page, shape, member in archive.execute("""
            SELECT DISTINCT o.source, o.page, o.shape_id, r.target
            FROM objects o
            JOIN object_references r ON (r.source, r.page, r.object_path) =
                                        (o.source, o.page, o.object_path)
            JOIN parts p ON (p.source, p.member) = (r.source, r.target)
            WHERE p.asset_sha256 IS NOT NULL
            ORDER BY o.source, o.page, o.shape_id, r.target
        """):
            slots[(source, page, str(shape))].append(member)
    return links, slots


def build_catalog(root, output):
    # Existing slot catalogs can be searched without opening native PPTX files.
    from native_reference import NativeLibrary

    inventory = list(iter_jsonl(root / "pages.jsonl"))
    manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8-sig"))
    expected = {
        (source["id"], page)
        for source in manifest["sources"]
        for page in range(1, source["pages"] + 1)
    }
    received = {(row["source"], row["page"]) for row in inventory}
    if received != expected or len(inventory) != len(expected):
        raise ValueError("pages.jsonl does not match every source deck and page in manifest.json")

    library = NativeLibrary(root)
    material_links, slot_links = load_material_links(root)
    output.parent.mkdir(parents=True, exist_ok=True)
    counts = Counter()
    temporary = None
    try:
        with NamedTemporaryFile(
            mode="w", encoding="utf-8", newline="\n", delete=False, dir=output.parent
        ) as stream:
            temporary = Path(stream.name)
            for row in inventory:
                source, page = row["source"], row["page"]
                details = library.inspect(source, page)
                shapes = [
                    {field: shape[field] for field in SLOT_FIELDS if field in shape}
                    for shape in details["shapes"]
                ]
                for shape in shapes:
                    members = slot_links.get((source, page, str(shape["id"])))
                    if members:
                        shape["material_members"] = members
                kinds = Counter(shape["type"] for shape in shapes)
                record = {
                    **{field: row[field] for field in PAGE_FIELDS},
                    "canvas_emu": [details["width_emu"], details["height_emu"]],
                    "shape_types": dict(sorted(kinds.items())),
                    "fillable_slots": sum(kinds[kind] for kind in FILLABLE),
                    "shapes": shapes,
                    "materials": material_links[(source, page)],
                }
                if "layout_equivalent_to" in row:
                    record["layout_equivalent_to"] = row["layout_equivalent_to"]
                stream.write(json.dumps(record, ensure_ascii=False, separators=(",", ":")) + "\n")
                counts[source] += 1
        os.replace(temporary, output)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return {"pages": sum(counts.values()), "sources": dict(counts), "catalog": str(output)}


def match_rank(row, terms):
    recipe = row["recipe"].casefold()
    text = " ".join((
        row["source"],
        row["label"],
        row["role"],
        row["priority"],
        recipe,
        *row["profile"],
        *(TYPE_LABELS.get(kind, kind) for kind in row["shape_types"]),
    )).casefold()
    if any(term not in text for term in terms):
        return None
    return (
        SOURCE_ORDER.get(row["source"], len(SOURCE_ORDER)),
        row["hidden"],
        -sum(term in recipe for term in terms),
        row["page"],
    )


def search(root, input_path, *, source, page, query, limit, slots, assets):
    terms = query.casefold().split()
    matches = []
    for row in iter_jsonl(input_path):
        if source and row["source"] != source:
            continue
        if page is not None and row["page"] != page:
            continue
        rank = match_rank(row, terms)
        if rank is not None:
            matches.append((rank, row))
    matches.sort(key=lambda entry: entry[0])

    for _, row in matches[:limit]:
        record = {
            "source": row["source"],
            "page": row["page"],
            "priority": row["priority"],
            "recipe": row["recipe"],
            "profile": row["profile"],
            "hidden": bool(row["hidden"]),
            "shape_types": row["shape_types"],
            "fillable_slots": row["fillable_slots"],
            "native": str(root / row["native_path"]),
            "preview": str(root / row["preview_path"]),
            "slide_part": row["member"],
            "linked_materials": len(row["materials"]),
        }
        if slots:
            record["shapes"] = row["shapes"]
            record["canvas_emu"] = row["canvas_emu"]
            record["source_geometry_sha256"] = row["geometry_sha256"]
        if assets:
            record["materials"] = [
                {**material, "file": str(root / material["file"])}
                for material in row["materials"]
            ]
        yield record


def suggest(root, input_path, *, source=None, intent=None, figures=None, text_slots=0,
            table=False, min_width=4.0, min_height=2.5, limit=10):
    """List runnable recipes first, then structural candidates for manual review."""
    from kw16_component_selection import load_component_profiles
    from native_distillation import load_profiles

    recipes = defaultdict(list)
    for whole, profiles in (
        (True, load_profiles(root)),
        (False, load_component_profiles(root)),
    ):
        for row in profiles:
            if intent is None or row["intent"] == intent:
                slots = row.get("figures", [])
                fields = {slot["field"] for slot in row.get("slots", {}).values() if "field" in slot}
                fields.update(f"figures.{i}.caption" for i, slot in enumerate(slots) if slot.get("label"))
                if row["intent"] == "comparison_table":
                    fields.add("table.rows")
                elif row["intent"] == "gantt_schedule":
                    fields.update({"timeline.periods", "timeline.items[].title",
                                   "timeline.items[].start", "timeline.items[].end"})
                recipes[(row["source"], row["page"])].append({
                    "id": row["id"], "intent": row["intent"], "kind": "page" if whole else "component",
                    "figure_count": len(slots), "fields": sorted(fields),
                })

    ranked = []
    for row in iter_jsonl(input_path):
        if row["hidden"] or source and source != row["source"]:
            continue
        src, page = row["source"], row["page"]
        known = recipes[(src, page)]
        if figures is not None:
            known = [item for item in known if item["figure_count"] == figures]
        if source is None and row.get("layout_equivalent_to") and not known:
            continue
        image_boxes = [
            {"id": shape["id"], "box_cm": shape["box_cm"], "parent": shape["parent"]}
            for shape in row["shapes"]
            if shape["type"] in {"picture", "chart"} and shape.get("box_cm")
            and shape["box_cm"][2] >= min_width and shape["box_cm"][3] >= min_height
        ]
        # Unprofiled source pictures are only visual candidates. Decorative,
        # overlapping or inherited material still needs local inspection.
        available_figures = len(image_boxes)
        editable_text = row["shape_types"].get("text", 0)
        tables = row["shape_types"].get("table", 0)
        if (editable_text < text_slots or table and not tables or
                figures is not None and available_figures < figures and not known):
            continue
        if intent and not known and figures is None and not text_slots and not table:
            continue
        tier = 0 if known else 1
        fit = abs(available_figures - figures) if figures is not None else 0
        rank = (tier, src not in {"gut", "guide"}, fit, SOURCE_ORDER.get(src, len(SOURCE_ORDER)), page)
        ranked.append((rank, {
            "source": src, "page": page,
            "status": "recipe" if known else "inspect",
            "recipes": known, "layout_tags": row["profile"], "text_slots": editable_text,
            "native_tables": tables, "candidate_image_boxes": image_boxes,
            "preview": str(root / row["preview_path"]),
            "native": str(root / row["native_path"]),
        }))
    ranked.sort(key=lambda item: item[0])
    seen = set()
    count = 0
    for _, row in ranked:
        if row["source"] in {"kw16", "kw16-9"}:
            identity = Path(row["preview"]).stem
            if identity in seen and row["status"] != "recipe":
                continue
            seen.add(identity)
        yield row
        count += 1
        if count >= limit:
            break


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    for command in ("build", "search", "suggest"):
        cli = commands.add_parser(command)
        cli.add_argument("--archive", type=Path, default=ARCHIVE)
        if command == "build":
            cli.add_argument("--output", type=Path, help="default: archive/native-pages.jsonl")
        elif command == "search":
            cli.add_argument("--catalog", type=Path, help="default: archive/native-pages.jsonl")
            cli.add_argument("--source", help="素材库清单中的来源别名")
            cli.add_argument("--page", type=int)
            cli.add_argument("--query", default="", help="构图、形状或来源关键词")
            cli.add_argument("--limit", default=12, type=int)
            cli.add_argument("--slots", action="store_true", help="列出原生对象 ID、类型和坐标")
            cli.add_argument("--assets", action="store_true", help="列出所选页关联的本地素材")
        else:
            cli.add_argument("--catalog", type=Path, help="default: archive/native-pages.jsonl")
            cli.add_argument("--source", help="素材库清单中的来源别名")
            cli.add_argument("--intent", help="查已编译的语义版式，例如 method_evidence")
            cli.add_argument("--figures", type=int, help="图数要求，含可检查的未验证候选")
            cli.add_argument("--text-slots", type=int, default=0, help="最少原生文字槽")
            cli.add_argument("--table", action="store_true", help="需要原生表格")
            cli.add_argument("--min-width", type=float, default=4.0, help="候选图位最低宽度 cm")
            cli.add_argument("--min-height", type=float, default=2.5, help="候选图位最低高度 cm")
            cli.add_argument("--limit", type=int, default=10)
    args = parser.parse_args()
    root = args.archive.resolve()
    try:
        if args.command == "build":
            result = build_catalog(root, args.output or root / CATALOG)
            print(json.dumps(result, ensure_ascii=False))
        elif args.command == "search":
            if args.limit < 1 or (args.page is not None and args.page < 1):
                parser.error("--limit and --page must be positive")
            for result in search(
                root, args.catalog or root / CATALOG, source=args.source, page=args.page,
                query=args.query, limit=args.limit, slots=args.slots, assets=args.assets,
            ):
                print(json.dumps(result, ensure_ascii=False))
        else:
            if (args.limit < 1 or args.text_slots < 0 or args.figures is not None and args.figures < 0
                    or not 0 < args.min_width <= 34 or not 0 < args.min_height <= 20):
                parser.error("invalid layout requirements")
            for result in suggest(
                root, args.catalog or root / CATALOG, source=args.source, intent=args.intent,
                figures=args.figures, text_slots=args.text_slots, table=args.table,
                min_width=args.min_width, min_height=args.min_height, limit=args.limit,
            ):
                print(json.dumps(result, ensure_ascii=False))
    except (OSError, ValueError, sqlite3.Error) as error:
        parser.exit(1, f"{error}\n")


if __name__ == "__main__":
    main()
