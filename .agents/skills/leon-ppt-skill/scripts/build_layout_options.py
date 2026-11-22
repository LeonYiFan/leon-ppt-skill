"""Build selectable PowerPoint layouts from the current plan and its layout manifest."""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path
import sys

import build_report as report


def read_object(path):
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict):
        raise ValueError(f"{path.name} 须为 JSON 对象")
    return data


def page_number(value, count):
    try:
        number = int(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"页码无效 {value!r}") from error
    if str(number) != value or not 1 <= number <= count:
        raise ValueError(f"页码超出范围 {value!r}（有效范围 1–{count}）")
    return number


def layout_groups(manifest, count):
    if not isinstance(slides := manifest.get("slides"), dict):
        raise ValueError("验收清单缺少 slides")
    if set(slides) != {str(n) for n in range(1, count + 1)}:
        raise ValueError("方案与验收清单页码不一致")
    for key, record in slides.items():
        number = page_number(key, count)
        selection = record.get("selection") or {}
        options = selection.get("options", [])
        if not isinstance(options, list) or len(options) > 3:
            raise ValueError(f"第 {number} 页的版式备选无效")
        if len(options) < 2:
            continue
        required = {"source", "page", "profile", "status", "intent"}
        if any(not isinstance(option, dict) or not required <= option.keys() for option in options):
            raise ValueError(f"第 {number} 页的备选缺少来源、页码或意图")
        profiles = [o["profile"] for o in options]
        if len(set(profiles)) != len(profiles) or profiles[0] != selection.get("profile"):
            raise ValueError(f"第 {number} 页的首选或备选 ID 不一致")
        if any(o["status"] not in {"selected", "selected_component"} for o in options):
            raise ValueError(f"第 {number} 页的备选状态无效")
        yield number, options


def choose_rows(groups, requested_pages, choices=None):
    for number in sorted(requested_pages):
        if number not in groups:
            raise ValueError(f"第 {number} 页没有至少两个可用版式")
        options = groups[number]
        choices_for_page = dict(zip("ABC", options))
        if choices is None:
            yield from ((number, letter, option) for letter, option in choices_for_page.items())
            continue
        letter = choices[number]
        if not isinstance(letter, str) or letter not in choices_for_page:
            allowed = "、".join(choices_for_page)
            raise ValueError(f"第 {number} 页请选择 {allowed}")
        yield number, letter, choices_for_page[letter]


def build_variants(plan, original_manifest, rows, base, template, native_library):
    preview = {**plan, "slides": []}
    mapping = []
    for index, (number, letter, option) in enumerate(rows, 1):
        page = deepcopy(plan["slides"][number - 1])
        page.update(native_auto=option["intent"], native_profile=option["profile"])
        preview["slides"].append(page)
        mapping.append({"preview_slide": index, "original_slide": number, "choice": letter,
                        "source": option["source"], "page": option["page"],
                        "profile": option["profile"]})
    package, current = BytesIO(), {}
    previous_base = report.BASE
    try:
        report.BASE = str(base)
        report.build(preview, template, package, native_library, audit_manifest=current)
    finally:
        report.BASE = previous_base
    for record, (_, _, option) in zip(mapping, rows, strict=True):
        original = original_manifest["slides"][str(record["original_slide"])]
        actual = current["slides"][str(record["preview_slide"])]
        selection = actual.get("selection") or {}
        fields = ("profile", "source", "page", "status")
        if any(selection.get(field) != option[field] for field in fields):
            raise ValueError(f"第 {record['original_slide']} 页备选已变更，请重新生成清单")
        old_figures = Counter(item["sha256"] for item in original["evidence"])
        new_figures = Counter(item["sha256"] for item in actual["evidence"])
        if old_figures != new_figures or original.get("notes_sources") != actual.get("notes_sources"):
            raise ValueError(f"第 {record['original_slide']} 页的原图或备注来源已变化，请重新生成清单")
    current["layout_options"] = mapping
    return package, current


def rebase_figure_paths(plan, source, destination):
    if source == destination:
        return

    def rebase(path):
        resolved = Path(os.path.abspath(source / path))
        try:
            return os.path.relpath(resolved, destination)
        except ValueError:
            # Relative paths cannot cross drives on Windows.
            return str(resolved)

    for slide in plan["slides"]:
        for figure in slide.get("figures") or ():
            figure["path"] = rebase(figure["path"])
        for card in slide.get("cards") or ():
            if card.get("figure"):
                card["figure"] = rebase(card["figure"])
        native = slide.get("native")
        native_specs = [native] if native else []
        native_specs.extend(slide.get("native_components") or ())
        for spec in native_specs:
            if "images" in spec:
                spec["images"] = {slot: rebase(path) for slot, path in spec["images"].items()}
    audit = plan.get("audit") or {}
    if "figure_sources" in audit:
        audit["figure_sources"] = {rebase(path): value for path, value in audit["figure_sources"].items()}


def validate_outputs(*paths):
    if len(set(paths)) != len(paths):
        raise ValueError("PPTX 与验收清单须使用不同路径")
    for path in paths:
        if not path.parent.is_dir():
            raise ValueError(f"输出目录不存在 {path.parent}")
        if path.exists():
            raise ValueError(f"输出文件已存在 {path}")


def write_json(path, data):
    with path.open("x", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
        file.write("\n")


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--template", default=report.DEFAULT_TEMPLATE)
    parser.add_argument("--native-library", default=report.DEFAULT_NATIVE_LIBRARY)
    parser.add_argument("--pages", help="指定原稿页码，例如 3,5,7")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--output", type=Path, help="输出包含 A/B/C 备选的 PPTX")
    mode.add_argument("--choices", type=Path, help='选项文件，例如 {"3":"B","5":"A"}')
    parser.add_argument("--preview-manifest", type=Path)
    parser.add_argument("--chosen-plan", type=Path)
    args = parser.parse_args()
    if args.output and (args.preview_manifest is None or args.chosen_plan is not None):
        parser.error("--output 须搭配 --preview-manifest")
    if args.choices and (args.chosen_plan is None or args.preview_manifest is not None or args.pages is not None):
        parser.error("--choices 须搭配 --chosen-plan；具体页码填写在选项文件中")
    return args


def main():
    args = parse_args()
    plan_bytes = args.plan.read_bytes()
    plan = json.loads(plan_bytes.decode("utf-8-sig"))
    manifest = read_object(args.manifest)
    if manifest.get("plan_sha256") != sha256(plan_bytes).hexdigest():
        raise ValueError("验收清单与当前计划不匹配，请以 --layout-options 重新构建")
    if not isinstance(plan, dict) or not isinstance(slides := plan.get("slides"), list) or not slides:
        raise ValueError("方案需包含非空 slides 列表")
    count = len(slides)
    groups = dict(layout_groups(manifest, count))
    base = args.plan.resolve().parent
    if args.choices:
        chosen_plan = args.chosen_plan.resolve()
        validate_outputs(chosen_plan)
        choices = {page_number(key, count): letter for key, letter in read_object(args.choices).items()}
        if not choices:
            raise ValueError("请选择至少一页版式")
        rows = list(choose_rows(groups, choices, choices))
        build_variants(plan, manifest, rows, base, args.template, args.native_library)
        for number, _, option in rows:
            slides[number - 1].update(native_auto=option["intent"], native_profile=option["profile"])
        rebase_figure_paths(plan, base, chosen_plan.parent)
        write_json(chosen_plan, plan)
        print(f"已写入 {len(rows)} 页的版式选择 {chosen_plan}")
        return

    deck_path, manifest_path = args.output.resolve(), args.preview_manifest.resolve()
    validate_outputs(deck_path, manifest_path)
    pages = ({page_number(value.strip(), count) for value in args.pages.split(",")}
             if args.pages is not None else set(groups))
    rows = list(choose_rows(groups, pages))
    if not rows:
        raise ValueError("当前计划没有可比较的版式，请检查清单的 options")
    package, preview_manifest = build_variants(plan, manifest, rows, base, args.template, args.native_library)
    with deck_path.open("xb") as file:
        file.write(package.getbuffer())
    write_json(manifest_path, preview_manifest)
    for row in preview_manifest["layout_options"]:
        print(f"预览第 {row['preview_slide']} 页 · 原稿 P{row['original_slide']} 选项 {row['choice']} · "
              f"{row['source']} P{row['page']} · {row['profile']}")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError) as error:
        sys.exit(str(error))
