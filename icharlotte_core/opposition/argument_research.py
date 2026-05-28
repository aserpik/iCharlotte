"""Retrieval-first grounding for opposition drafting.

Per argument: generate CourtListener search queries, hybrid-search CA case
law, fetch real opinion text for the top candidates, then have an LLM
re-rank/select the best 3-5 cases with a VERBATIM supporting passage.
Returns RetrievedAuthority records the drafter cites from.
"""

from __future__ import annotations

import concurrent.futures
import json
import logging
import os
import re
from typing import Any, Callable

from icharlotte_core.opposition.models import RetrievedAuthority

logger = logging.getLogger(__name__)

LLMCallback = Callable[[str, str], str]


def _loads_json(text: str) -> dict[str, Any]:
    if not isinstance(text, str):
        return {}
    cleaned = text.strip()
    fenced = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", cleaned, re.DOTALL | re.I)
    if fenced:
        cleaned = fenced.group(1).strip()
    try:
        data = json.loads(cleaned)
    except (TypeError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def generate_search_queries(argument: str, *, llm_callback: LLMCallback) -> list[str]:
    """Turn one argument into 1-2 CourtListener search queries."""
    from icharlotte_core.prompt_manager import get_prompt
    from icharlotte_core.opposition import prompts as default_prompts

    template = get_prompt("oppose_motion", "research_queries") or default_prompts.RESEARCH_QUERIES_PROMPT
    user_prompt = template.format(argument=argument or "")
    try:
        response = llm_callback("", user_prompt) or ""
    except Exception:
        logger.warning("research query generation failed", exc_info=True)
        return []
    data = _loads_json(response)
    raw = data.get("queries") if isinstance(data, dict) else None
    if not isinstance(raw, list):
        return []
    queries = [str(q).strip() for q in raw if str(q).strip()]
    return queries[:2]


def _normalize_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def _format_candidates(candidates: list[dict], *, excerpt_chars: int = 6000) -> str:
    blocks: list[str] = []
    for c in candidates:
        excerpt = (c.get("text") or "")[:excerpt_chars]
        blocks.append(
            f"[{c.get('cluster_id')}] {c.get('case_name', '')}, {c.get('citation', '')}\n{excerpt}"
        )
    return "\n\n".join(blocks)


def select_authorities(
    proposition: str,
    candidates: list[dict],
    *,
    argument_text: str,
    argument_id: str = "",
    llm_callback: LLMCallback,
) -> list[RetrievedAuthority]:
    """LLM picks the best candidates; citation comes from metadata, and the
    quoted passage must appear verbatim in the candidate's opinion text."""
    from icharlotte_core.prompt_manager import get_prompt
    from icharlotte_core.opposition import prompts as default_prompts

    if not candidates:
        return []

    by_id = {str(c.get("cluster_id")): c for c in candidates}
    template = get_prompt("oppose_motion", "rerank_select") or default_prompts.RERANK_SELECT_PROMPT
    user_prompt = template.format(
        proposition=proposition or "",
        candidates=_format_candidates(candidates),
    )
    try:
        response = llm_callback("", user_prompt) or ""
    except Exception:
        logger.warning("rerank/select failed", exc_info=True)
        return []

    data = _loads_json(response)
    selections = data.get("selections") if isinstance(data, dict) else None
    if not isinstance(selections, list):
        return []

    out: list[RetrievedAuthority] = []
    for sel in selections:
        if not isinstance(sel, dict):
            continue
        cand = by_id.get(str(sel.get("id")))
        if not cand:
            continue
        passage = str(sel.get("passage", "")).strip()
        if not passage or _normalize_ws(passage) not in _normalize_ws(cand.get("text", "")):
            continue  # drop unverifiable / fabricated passages
        out.append(
            RetrievedAuthority(
                argument_id=argument_id,
                argument_text=argument_text,
                cluster_id=str(cand.get("cluster_id") or ""),
                case_name=cand.get("case_name", ""),
                citation=cand.get("citation", ""),
                supports=str(sel.get("supports", "")).strip(),
                passage=passage,
                opinion_url=cand.get("opinion_url", ""),
            )
        )
    return out
