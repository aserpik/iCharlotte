import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest


@pytest.fixture
def session_dir(tmp_path):
    sd = tmp_path / "session"
    sd.mkdir()
    (sd / "digests").mkdir()
    return sd


def _make_session(session_dir, deponent_sources=None, context_sources=None):
    payload = {
        "version": 1,
        "phase": "awaiting_input",
        "deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
        "style": "discovery", "free_text_notes": "",
        "per_topic_flags": {"strategic_note": True, "source_facts": False,
                             "impeachment_hook": False, "objection_alts": False},
        "case_root": str(session_dir.parent),
        "deponent_sources": deponent_sources or [], "context_sources": context_sources or [],
        "digests_index": ["src1.pdf"],
        "topics_warning": None,
    }
    (session_dir / "session.json").write_text(json.dumps(payload), encoding="utf-8")


def _make_topics(session_dir, topic_count=2):
    topics = [
        {"id": f"t{i:02d}", "title": f"Topic {i}", "strategic_note": "s",
         "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False}
        for i in range(1, topic_count + 1)
    ]
    (session_dir / "topics.json").write_text(json.dumps({"topics": topics}), encoding="utf-8")


def _make_digest(session_dir, source_id="src1.pdf"):
    (session_dir / "digests" / f"{source_id}.json").write_text(
        json.dumps({"source_id": source_id, "source_kind": "other",
                    "deponent_statements": [], "factual_anchors": [],
                    "inconsistencies": [], "summary": ""}), encoding="utf-8")


def _phase2_llm():
    """Returns valid payloads for question generation, dedup, polish."""
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kwargs):
        if pass_name == "topic_questions":
            return json.dumps({"topic_id": "t??", "questions": [
                {"n": 1, "text": "Generated Q1"}, {"n": 2, "text": "Generated Q2"}]})
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": ["Gap A"],
                                "renumber_after_dedup": True})
        if pass_name == "polish":
            # Echo the input outline unchanged (well-shaped polish).
            return text
        return ""
    caller = MagicMock()
    caller.call.side_effect = call
    return caller


def test_phase2_writes_outline_docx_and_md(session_dir):
    from Scripts.depo_prep_lib import phase2

    _make_session(session_dir)
    _make_topics(session_dir, topic_count=2)
    _make_digest(session_dir)

    phase2.run_phase2(
        session_path=str(session_dir / "session.json"),
        llm_caller=_phase2_llm(),
        progress=lambda *a, **k: None,
    )

    assert (session_dir / "outline.docx").exists()
    assert (session_dir / "outline.md").exists()
    md = (session_dir / "outline.md").read_text(encoding="utf-8")
    assert "Jane Doe" in md
    assert "Topic 1" in md
    assert "Gap A" in md


def test_phase2_handles_per_topic_failure(session_dir):
    from Scripts.depo_prep_lib import phase2

    _make_session(session_dir)
    _make_topics(session_dir, topic_count=2)
    _make_digest(session_dir)

    # LLM throws on topic_questions, returns valid for dedup/polish.
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kwargs):
        if pass_name == "topic_questions":
            raise RuntimeError("rate limit")
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": [], "renumber_after_dedup": False})
        if pass_name == "polish":
            return text
        return ""
    caller = MagicMock(); caller.call.side_effect = call

    # Should not raise - Phase 2 finishes with empty topic Qs and an error banner.
    phase2.run_phase2(
        session_path=str(session_dir / "session.json"),
        llm_caller=caller,
        progress=lambda *a, **k: None,
    )
    md = (session_dir / "outline.md").read_text(encoding="utf-8")
    assert "Topic 1" in md
    # Render still produces a doc; user sees gaps but topics present.
