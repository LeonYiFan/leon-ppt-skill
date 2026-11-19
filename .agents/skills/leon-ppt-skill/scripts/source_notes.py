"""Move citation paragraphs into speaker notes without rewriting the talk."""

import re
from unicodedata import normalize


SOURCE_SHAPE_NAME = "图源说明"
LABEL = r"图源|图片来源|数据来源|资料来源|参考来源|来源"
PREFIX = re.compile(rf"^(?:【(?:{LABEL})】|(?:{LABEL})\s*[：:｜|])\s*")


def source_payload(line):
    return PREFIX.sub("", line.strip(), count=1).strip()


def source_key(line):
    return " ".join(normalize("NFKC", source_payload(line)).split())


def append_note_sources(notes, sources):
    """Return the notes and exact lines that carry the supplied citations."""
    notes = "" if notes is None else str(notes).strip()
    existing = {}
    for line in notes.splitlines():
        if key := source_key(line):
            existing.setdefault(key, line.strip())
    added, citations = [], []
    seen = set()
    for value in sources:
        if value is None:
            continue
        if not isinstance(value, str):
            raise ValueError("source 须为文字")
        for line in value.splitlines():
            key = source_key(line)
            if not key or key in seen:
                continue
            seen.add(key)
            citation = existing.get(key)
            if citation is None:
                citation = "【图源】" + source_payload(line)
                added.append(citation)
            citations.append(citation)
    return "\n".join(part for part in [notes, *added] if part), citations
