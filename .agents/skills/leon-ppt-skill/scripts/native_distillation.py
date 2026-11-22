"""Bind authored content to indexed source slots without dropping authored fields."""
from pathlib import Path
import json


PROFILE_FILES = ("native-layout-profiles.json", "gut-distilled-profiles.json",
                 "guide-distilled-profiles.json")
EXTRA_PROFILES_FILE = "supplemental-layout-profiles.json"
REFERENCES = Path(__file__).resolve().parents[1] / "references"
FORM_FIELDS = {
    "items", "metric", "metrics", "detail", "stages", "formula", "headers",
    "lead", "method_heading", "shared_result", "strengths", "outcome", "timeline", "hub",
    "active_stage",
}


class CandidateMismatch(ValueError):
    """The requested content does not fit one particular source profile."""


def load_profiles(archive_root=None):
    directory = Path(archive_root) / "profiles" if archive_root is not None else REFERENCES
    if not directory.is_dir():
        directory = REFERENCES
    paths = [directory / name for name in (*PROFILE_FILES, EXTRA_PROFILES_FILE)
             if (directory / name).is_file()]
    return [row for file in paths for row in json.loads(file.read_text(encoding="utf-8"))["profiles"]]


def _leaf_fields(value, field):
    if isinstance(value, dict):
        for key, child in value.items():
            yield from _leaf_fields(child, f"{field}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _leaf_fields(child, f"{field}.{index}")
    elif value is not None and (not isinstance(value, str) or value.strip()):
        yield field, value


def authored_fields(page):
    component = page.get("component")
    component = component if isinstance(component, dict) else {}
    payload = {key: page[key] for key in FORM_FIELDS & page.keys() if key != "active_stage"}
    if "items" in payload and "items" in component:
        raise ValueError("items 只能在页级或 component 中填写一次")
    if "items" in component:
        payload["items"] = component["items"]
    return {name: value for key, body in payload.items() for name, value in _leaf_fields(body, key)}


def slot_values(page, profile):
    """Yield configured text slots and validate full consumption of data fields."""
    fields = authored_fields(page)
    for index, figure in enumerate(page.get("figures") or ()):
        slot = profile["figures"][index]
        if not slot.get("label") and figure.get("caption"):
            fields[f"figures.{index}.caption"] = figure["caption"]
    used = set()
    for shape_id, slot in profile.get("slots", {}).items():
        field = slot.get("field")
        value = fields.get(field) if field is not None else slot.get("value")
        if value is None and slot.get("optional"):
            continue
        if not isinstance(value, str) or not value.strip():
            raise CandidateMismatch(f"槽位 {field or shape_id} 需要非空文本")
        lines = value.strip().splitlines()
        if len(lines) > slot.get("max_lines", 1) or any(not line.strip() for line in lines):
            raise CandidateMismatch(f"槽位 {field or shape_id} 超过允许行数")
        if field is not None:
            used.add(field)
        yield shape_id, value.strip(), slot
    if extra := fields.keys() - used:
        raise CandidateMismatch(f"版式未覆盖：{', '.join(sorted(extra))}")


def _records(value, count, fields):
    """Match an authored structure without guessing from the page wording."""
    return (isinstance(value, list) and len(value) == count
            and all(isinstance(item, dict) and item.keys() == fields for item in value))


def _structured_intent(page, figure_count):
    fields = FORM_FIELDS & page.keys()
    items = page.get("items")
    stages = page.get("stages")
    if figure_count == 1 and fields == {"items"} and _records(items, 4, {"title", "body", "tag", "result"}):
        return "study_overview"
    if (figure_count == 1 and fields == {"lead", "detail"}
            and isinstance(page["detail"], dict) and page["detail"].keys() == {"body"}):
        return "study_context"
    if figure_count == 2:
        if fields == {"headers", "metrics"}:
            return "cohort_readouts"
        if fields == {"items"}:
            if _records(items, 2, {"title", "value", "body"}):
                return "paired_counts"
            if _records(items, 2, {"body"}):
                return "paired_evidence"
        if fields == {"items", "lead"} and _records(items, 2, {"title", "body"}):
            return "paired_gaps"
        if (fields == {"items", "stages", "lead", "outcome"}
                and _records(items, 2, {"title", "body"})
                and _records(stages, 3, {"title", "body"})):
            return "staged_design"
    if figure_count == 3 and fields == {"lead", "outcome"}:
        return "evidence_to_question"
    if figure_count == 6 and fields == {"headers", "stages"}:
        return "stratified_evidence"

    if figure_count:
        return None
    if fields == {"headers", "items"} and _records(items, 3, {"title", "body", "right_title", "right_body"}):
        return "paired_comparison"
    if fields == {"lead", "items"} and _records(items, 2, {"title", "steps"}):
        return "hypothesis_branches"
    if fields == {"lead", "items"} and _records(items, 3, {"title", "tag", "body"}):
        return "objective_tracks"
    if fields == {"lead", "outcome", "items"} and _records(items, 3, {"title", "body"}):
        return "convergent_validation"
    if fields == {"items"} and _records(items, 3, {"title", "body"}):
        return "parallel_cards"
    if fields == {"items"} and _records(items, 4, {"title", "body", "tag"}):
        return "milestones"
    if fields == {"items", "hub"} and _records(items, 6, {"title", "body"}):
        return "evidence_cluster"
    if fields == {"items"} and _records(items, 5, {"title", "body", "tag"}):
        return "stage_handoffs"
    if (fields == {"stages", "items"} and isinstance(stages, list) and len(stages) == 4
            and _records(items, 2, {"title", "steps"})):
        return "paired_schedule"
    if fields == {"detail"} and isinstance(page.get("table"), dict):
        if isinstance(page["detail"], dict) and page["detail"].keys() == {"footer"}:
            return "comparison_table"
    if fields == {"timeline"} and isinstance(timeline := page["timeline"], dict):
        tasks = timeline.get("items")
        if isinstance(tasks, list) and tasks and _records(tasks, len(tasks), {"title", "start", "end"}):
            return "gantt_schedule"
        if _records(tasks, 3, {"title", "steps"}):
            return "milestones"
    return None


def infer_intent(page):
    """Recognize existing page schemas that have a genuine native counterpart."""
    if page.get("type") in {"native", "cover", "agenda", "divider", "thanks"} or page.get("native_components"):
        return None
    if page.get("type") == "triple" and len(page.get("cards") or ()) == 3:
        return "parallel_cards"
    component = page.get("component") or {}
    if isinstance(component, dict) and component:
        return {
            "risk_response": "risk_response",
            "parallel_workstreams": "workstreams",
            "deliverable_checkpoints": "milestones",
            "independent_findings": "parallel_cards",
            "same_basis_comparison": "paired_comparison",
        }.get(component.get("intent"))
    if page.get("type") != "bullets":
        return None
    figures = page.get("figures") or []
    if not isinstance(figures, list):
        return None
    if any(not isinstance(figure, dict) or "box" in figure for figure in figures):
        return None
    if intent := _structured_intent(page, len(figures)):
        return intent
    if 1 <= len(figures) <= 4:
        if page.get("metrics"):
            return "metric_evidence"
        if page.get("detail") and page.get("items") and page.get("stages"):
            return "method_evidence"
        if (len(figures) == 1 and page.get("lead") and page.get("method_heading")
                and _records(page.get("items"), 2, {"title", "body", "tag"})):
            return "method_evidence"
        return "figure_evidence"
    return None
