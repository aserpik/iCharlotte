"""Stage C — phrasing-only polish with structural validation."""
from __future__ import annotations

import json
from typing import Dict

from .prompts import build_polish_prompt
from .source_digest import _parse_llm_json


def _shape_sig(outline: dict) -> Dict[str, int]:
    """Return {topic_id: question_count} for structural comparison."""
    return {
        t["topic_id"]: len(t.get("questions", []))
        for t in outline.get("topics", [])
    }


def polish_outline(*, outline: dict, llm_caller) -> dict:
    """Run polish LLM. If the result changes topic_ids or question counts, revert to original."""
    try:
        prompt, _ = build_polish_prompt(outline_text=json.dumps(outline, indent=2))
        raw = llm_caller.call(
            prompt=prompt,
            text=json.dumps(outline, indent=2),
            task_type="general",
            agent_id="DepoPrep",
            pass_name="polish",
        )
        polished = _parse_llm_json(raw)
        if _shape_sig(polished) != _shape_sig(outline):
            return outline  # structural violation; revert
        return polished
    except Exception:
        return outline
