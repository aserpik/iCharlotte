"""Stage A — per-topic question generation."""
from __future__ import annotations

import json
import re
from typing import Dict

from .prompts import build_per_topic_questions_prompt
from .source_digest import _parse_llm_json


# "med.json#factual_anchors[0]"  →  ("med.json", "factual_anchors", 0)
_REF_RE = re.compile(r"^([^#]+)#([a-z_]+)\[(\d+)\]$")


def _resolve_refs(refs, digests_by_source) -> str:
    """Build text payload listing only the referenced digest entries."""
    if not refs:
        return ""
    blocks = []
    for r in refs:
        m = _REF_RE.match(r.strip())
        if not m:
            continue
        src_id, field, idx = m.group(1), m.group(2), int(m.group(3))
        digest = digests_by_source.get(src_id)
        if not digest:
            continue
        entries = digest.get(field) or []
        if 0 <= idx < len(entries):
            blocks.append(
                f"=== {src_id} :: {field}[{idx}] ===\n"
                + json.dumps(entries[idx], indent=2, ensure_ascii=False)
            )
    return "\n\n".join(blocks)


def _full_digest_payload(digests_by_source) -> str:
    blocks = []
    for src_id, digest in digests_by_source.items():
        blocks.append(f"=== {src_id} ===\n" + json.dumps(digest, indent=2, ensure_ascii=False))
    return "\n\n".join(blocks)


def generate_questions_for_topic(
    *, topic: dict, digests_by_source: Dict[str, dict],
    llm_caller, deponent_name: str, deponent_role: str,
    style: str, free_text_notes: str, flags: Dict[str, bool],
) -> dict:
    """Generate a TopicQuestions-shaped dict for one topic.

    On LLM failure, returns {"topic_id": ..., "questions": [], "error": "..."}.
    """
    is_lawyer_added = bool(topic.get("lawyer_added"))
    refs = topic.get("relevant_digest_refs") or []

    if is_lawyer_added or not refs:
        digest_text = _full_digest_payload(digests_by_source)
    else:
        digest_text = _resolve_refs(refs, digests_by_source)

    prompt, text_payload = build_per_topic_questions_prompt(
        deponent_name=deponent_name, deponent_role=deponent_role, style=style,
        topic_title=topic["title"], strategic_note=topic.get("strategic_note", ""),
        digest_excerpts_text=digest_text, free_text_notes=free_text_notes,
        include_strategic_note=bool(flags.get("strategic_note")),
        include_source_facts=bool(flags.get("source_facts")),
        include_impeachment_hook=bool(flags.get("impeachment_hook")),
        include_objection_alts=bool(flags.get("objection_alts")),
    )

    try:
        raw = llm_caller.call(
            prompt=prompt, text=text_payload,
            task_type="general", agent_id="DepoPrep", pass_name="topic_questions",
        )
        data = _parse_llm_json(raw)
        data["topic_id"] = topic["id"]  # force consistency
        # Ensure required shape.
        if not isinstance(data.get("questions"), list):
            data["questions"] = []
        return data
    except Exception as e:
        return {"topic_id": topic["id"], "questions": [], "error": str(e)}
