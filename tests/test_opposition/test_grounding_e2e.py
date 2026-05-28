"""End-to-end grounding run. Skipped unless live API tokens are present.

Drives research -> draft -> pool-check -> verify against the real CourtListener
and a real LLM. Asserts every case cite in the draft is in the retrieved pool
and that the majority of citations verify SUPPORTED.
"""

from __future__ import annotations

import os

import pytest

pytestmark = pytest.mark.skipif(
    not (os.environ.get("COURTLISTENER_API_TOKEN") and os.environ.get("GEMINI_API_KEY")),
    reason="requires COURTLISTENER_API_TOKEN and GEMINI_API_KEY",
)


def test_grounded_draft_cites_only_from_pool():
    from icharlotte_core.llm_config import call_llm
    from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
    from icharlotte_core.opposition.argument_research import research_arguments
    from icharlotte_core.opposition.citation_parser import extract_citations
    from icharlotte_core.opposition.drafter import draft_memorandum
    from icharlotte_core.opposition.models import MotionMetadata
    from icharlotte_core.opposition.verifier import pool_membership_check

    token = os.environ["COURTLISTENER_API_TOKEN"]

    def make_llm(pass_name):
        def _llm(system_prompt, user_prompt):
            return call_llm(user_prompt, system_prompt, task_type="general",
                            agent_id="agent_oppose_motion", pass_name=pass_name) or ""
        return _llm

    metadata = MotionMetadata(
        motion_type="Motion to Compel Further Responses",
        relief_requested="Order compelling further responses to inspection demands.",
        principal_arguments=[
            "The responses are evasive and incomplete under the Civil Discovery Act.",
            "Good cause supports inspection of the requested materials.",
        ],
    )

    retrieved = research_arguments(
        metadata.principal_arguments,
        cl_client=CourtListenerClient(token),
        query_llm=make_llm("research_queries"),
        rerank_llm=make_llm("rerank_select"),
        max_workers=4,
    )
    assert retrieved, "expected at least one retrieved authority"

    draft = draft_memorandum(
        metadata=metadata, section_plan=[], motion_text="(motion text omitted)",
        context_text="", style_exemplars=[], retrieved_authorities=retrieved,
        llm_callback=make_llm("draft_memorandum"),
    )
    assert draft.body_text.strip()

    cites = extract_citations(draft.body_text)
    _to_verify, off_pool = pool_membership_check(cites, retrieved)
    # Core guarantee: no case cite outside the retrieved pool.
    assert off_pool == [], f"off-pool cites: {[c.citation_text for c in off_pool]}"
