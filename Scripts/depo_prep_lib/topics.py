"""Stage 1.3 - cluster per-source digests into 8-15 topics."""
from __future__ import annotations

import json
from dataclasses import dataclass
from typing import List, Optional

from .prompts import build_topic_clustering_prompt
from .schemas import validate_topics_dict, Topic
from .source_digest import _parse_llm_json  # reuse fence-stripping JSON parser


_MAX_TOPICS = 20
_MIN_TOPICS_WARN = 3


@dataclass
class TopicsResult:
    topics: List[dict]   # plain dicts (matching Topic schema), ready to serialize
    warning: Optional[str]


def _digests_to_summary_text(digests: List[dict]) -> str:
    """Render the digest list as a single text payload for the clustering prompt."""
    blocks = []
    for d in digests:
        blocks.append(f"=== DIGEST: {d.get('source_id', 'unknown')} ===\n"
                      + json.dumps(d, indent=2, ensure_ascii=False))
    return "\n\n".join(blocks)


def cluster_topics(
    *,
    digests: List[dict],
    llm_caller,
    deponent_name: str,
    deponent_role: str,
    style: str,
    free_text_notes: str,
) -> TopicsResult:
    prompt, text_payload = build_topic_clustering_prompt(
        deponent_name=deponent_name, deponent_role=deponent_role,
        style=style, free_text_notes=free_text_notes,
        digests_summary_text=_digests_to_summary_text(digests),
    )

    raw = llm_caller.call(
        prompt=prompt, text=text_payload,
        task_type="general", agent_id="DepoPrep", pass_name="topic_clustering",
    )

    data = _parse_llm_json(raw)
    validate_topics_dict(data)

    topics = list(data["topics"])
    warning = None
    if len(topics) > _MAX_TOPICS:
        topics = topics[:_MAX_TOPICS]
        warning = f"LLM produced more than {_MAX_TOPICS} topics; truncated."
    elif len(topics) < _MIN_TOPICS_WARN:
        warning = ("Source material appears thin - only "
                   f"{len(topics)} topic(s) emerged. Consider adding more sources "
                   "or detail in your strategy notes.")

    # Normalize each topic through the Topic dataclass to fill defaults.
    normalized = [Topic.from_dict(t).to_dict() for t in topics]

    return TopicsResult(topics=normalized, warning=warning)
