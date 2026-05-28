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
