"""Stage B - dedup + coverage check."""
from __future__ import annotations

import re
from typing import Dict, List

from .prompts import build_dedup_prompt
from .source_digest import _parse_llm_json


_REF_RE = re.compile(r"^([^.]+)\.q(\d+)$")


def _summarize_topic_outputs(topic_outputs: List[dict]) -> str:
    lines = []
    for t in topic_outputs:
        tid = t["topic_id"]
        for q in t.get("questions", []):
            lines.append(f"{tid}.q{q['n']}: {q.get('text', '')[:140]}")
    return "\n".join(lines)


def _summarize_digests(digests_by_source: Dict[str, dict]) -> str:
    return "\n".join(
        f"{src_id}: {(d.get('summary') or '').strip()[:140]}"
        for src_id, d in digests_by_source.items()
    )


def dedup_and_coverage(
    *, topic_outputs: List[dict], digests_by_source: Dict[str, dict], llm_caller,
) -> dict:
    """Run the dedup/coverage LLM call. On any failure return an empty result."""
    prompt, text_payload = build_dedup_prompt(
        topic_outputs_summary=_summarize_topic_outputs(topic_outputs),
        digest_summary=_summarize_digests(digests_by_source),
    )
    try:
        raw = llm_caller.call(
            prompt=prompt, text=text_payload,
            task_type="general", agent_id="DepoPrep", pass_name="dedup",
        )
        data = _parse_llm_json(raw)
        if not isinstance(data, dict):
            raise ValueError("dedup payload not a dict")
        data.setdefault("duplicates", [])
        data.setdefault("coverage_gaps", [])
        data.setdefault("renumber_after_dedup", True)
        return data
    except Exception:
        return {"duplicates": [], "coverage_gaps": [], "renumber_after_dedup": False}


def apply_dedup(topic_outputs: List[dict], dedup: dict) -> List[dict]:
    """Return a NEW list of topic_outputs with duplicate-drops applied and renumbered."""
    drops_by_topic: Dict[str, set] = {}
    for d in dedup.get("duplicates", []):
        drop_ref = d.get("drop", "")
        m = _REF_RE.match(drop_ref)
        if not m:
            continue
        topic_id, n = m.group(1), int(m.group(2))
        drops_by_topic.setdefault(topic_id, set()).add(n)

    out = []
    for t in topic_outputs:
        tid = t["topic_id"]
        drops = drops_by_topic.get(tid, set())
        kept = [q for q in t.get("questions", []) if q["n"] not in drops]
        if dedup.get("renumber_after_dedup", True):
            for i, q in enumerate(kept, 1):
                q = dict(q)
                q["n"] = i
                kept[i - 1] = q
        new_topic = dict(t)
        new_topic["questions"] = kept
        out.append(new_topic)
    return out
