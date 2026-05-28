"""End-to-end Depo Prep integration test (no subprocess)."""
import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest

FIXTURES = Path(__file__).parent / "fixtures" / "depo_prep_sources"


def _stub_llm():
    """Stub LLM keyed by pass_name to return canned, well-shaped payloads."""
    caller = MagicMock()

    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kw):
        if pass_name == "source_digest":
            return json.dumps({
                "source_id": "placeholder",  # orchestrator overwrites
                "source_kind": "deposition_transcript",
                "deponent_statements": [
                    {"text": "I had no prior back problems.", "location": "p.1:3",
                     "context": "Direct exam"}
                ],
                "factual_anchors": [
                    {"fact": "Denied chiropractic care", "location": "p.1:5",
                     "topic_tags": ["prior_care"]}
                ],
                "inconsistencies": [],
                "summary": "Witness denies any prior back issues.",
            })
        if pass_name == "topic_clustering":
            return json.dumps({"topics": [
                {"id": "t01", "title": "Prior back issues",
                 "strategic_note": "Establish lack of prior issues for causation.",
                 "relevant_digest_refs": [],
                 "default_checked": True, "lawyer_added": False},
                {"id": "t02", "title": "Mechanism of collision",
                 "strategic_note": "Pin down what plaintiff observed.",
                 "relevant_digest_refs": [],
                 "default_checked": True, "lawyer_added": False},
                {"id": "t03", "title": "Treatment timeline",
                 "strategic_note": "Walk through care, gaps.",
                 "relevant_digest_refs": [],
                 "default_checked": True, "lawyer_added": False},
            ]})
        if pass_name == "topic_questions":
            return json.dumps({"topic_id": "tXX", "questions": [
                {"n": 1, "text": "Before August 15, 2024, did you experience any back pain?"},
                {"n": 2, "text": "Did you ever see a chiropractor before that date?"},
            ]})
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": [
                "No question addresses prior auto accidents."
            ], "renumber_after_dedup": True})
        if pass_name == "polish":
            return text  # echo unchanged
        return ""

    caller.call.side_effect = call
    return caller


def test_full_phase1_phase2_pipeline(tmp_path):
    from Scripts.depo_prep_lib import phase1, phase2

    case_root = tmp_path / "Smith v. Jones"
    case_root.mkdir()

    config = {
        "deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
        "deponent_sources": [str(FIXTURES / "jane_doe_depo_excerpt.txt")],
        "context_sources": [str(FIXTURES / "complaint_excerpt.txt")],
        "style": "lockdown", "free_text_notes": "Focus on causation.",
        "per_topic_flags": {
            "strategic_note": True, "source_facts": True,
            "impeachment_hook": False, "objection_alts": False,
        },
        "case_root": str(case_root),
    }

    llm = _stub_llm()

    # Phase 1
    session_path = phase1.run_phase1(
        config=config, llm_caller=llm, progress=lambda *a, **k: None)
    assert Path(session_path).exists()
    session_dir = Path(session_path).parent
    topics = json.loads((session_dir / "topics.json").read_text(encoding="utf-8"))
    assert len(topics["topics"]) == 3

    # Simulate user editing topics: uncheck one, add a custom topic.
    topics["topics"][1]["default_checked"] = False
    topics["topics"].append({
        "id": "t99", "title": "Social media activity post-accident",
        "strategic_note": "Lawyer-added. Look for inconsistent depictions.",
        "relevant_digest_refs": [], "default_checked": True, "lawyer_added": True,
    })
    (session_dir / "topics.json").write_text(json.dumps(topics), encoding="utf-8")

    # Phase 2
    phase2.run_phase2(session_path=session_path, llm_caller=llm,
                      progress=lambda *a, **k: None)

    docx = session_dir / "outline.docx"
    md = session_dir / "outline.md"
    assert docx.exists()
    assert md.exists()

    md_text = md.read_text(encoding="utf-8")
    assert "Jane Doe" in md_text
    assert "Prior back issues" in md_text
    # The unchecked topic (Mechanism of collision) should NOT appear.
    assert "Mechanism of collision" not in md_text
    # Lawyer-added topic SHOULD appear.
    assert "Social media activity post-accident" in md_text
    # Coverage gap appears.
    assert "prior auto accidents" in md_text


def test_full_pipeline_handles_per_topic_failure(tmp_path):
    """If one topic's LLM call fails, the rest of the outline still renders."""
    from Scripts.depo_prep_lib import phase1, phase2

    case_root = tmp_path / "C"
    case_root.mkdir()
    config = {
        "deponent_name": "X", "deponent_role": "P",
        "deponent_sources": [str(FIXTURES / "jane_doe_depo_excerpt.txt")],
        "context_sources": [],
        "style": "discovery", "free_text_notes": "",
        "per_topic_flags": {"strategic_note": False, "source_facts": False,
                             "impeachment_hook": False, "objection_alts": False},
        "case_root": str(case_root),
    }

    # LLM: fail every other topic_questions call.
    call_count = {"topic_questions": 0}
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kw):
        if pass_name == "source_digest":
            return json.dumps({"source_id": "x", "source_kind": "other",
                               "deponent_statements": [], "factual_anchors": [],
                               "inconsistencies": [], "summary": ""})
        if pass_name == "topic_clustering":
            return json.dumps({"topics": [
                {"id": f"t{i:02d}", "title": f"T{i}", "strategic_note": "s",
                 "relevant_digest_refs": [], "default_checked": True, "lawyer_added": False}
                for i in range(1, 4)
            ]})
        if pass_name == "topic_questions":
            call_count["topic_questions"] += 1
            if call_count["topic_questions"] % 2 == 0:
                raise RuntimeError("rate limit")
            return json.dumps({"topic_id": "tXX",
                                "questions": [{"n": 1, "text": "Question"}]})
        if pass_name == "dedup":
            return json.dumps({"duplicates": [], "coverage_gaps": [], "renumber_after_dedup": False})
        if pass_name == "polish":
            return text
        return ""

    caller = MagicMock()
    caller.call.side_effect = call

    session_path = phase1.run_phase1(config=config, llm_caller=caller,
                                      progress=lambda *a, **k: None)
    phase2.run_phase2(session_path=session_path, llm_caller=caller,
                      progress=lambda *a, **k: None)

    session_dir = Path(session_path).parent
    md = (session_dir / "outline.md").read_text(encoding="utf-8")
    # All three topics should appear; some with questions, some without.
    assert "T1" in md
    assert "T2" in md
    assert "T3" in md
