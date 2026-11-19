"""Bind authored content to indexed source slots without dropping authored fields."""
from pathlib import Path
import json


PROFILE_FILES = ("native-layout-profiles.json", "gut-distilled-profiles.json",
                 "guide-distilled-profiles.json")
REFERENCES = Path(__file__).resolve().parents[1] / "references"
FORM_FIELDS = {
    "items", "metric", "metrics", "detail", "stages", "formula", "headers",
    "lead", "method_heading", "shared_result", "strengths", "outcome", "timeline", "hub",
    "active_stage",
}


class CandidateMismatch(ValueError):
    """The requested content does not fit one particular source profile."""


def load_profiles():
    return [
        row
        for name in PROFILE_FILES
        for row in json.loads((REFERENCES / name).read_text(encoding="utf-8"))["profiles"]
    ]


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
    if page.get("type") == "bullets" and 1 <= len(page.get("figures") or ()) <= 4:
        if page.get("metrics"):
            return "metric_evidence"
        if page.get("detail") and page.get("items") and page.get("stages"):
            return "method_evidence"
        if all(isinstance(figure, dict) and "box" not in figure for figure in page["figures"]):
            return "figure_evidence"
    return None
