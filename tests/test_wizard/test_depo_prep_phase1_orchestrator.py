"""Phase 1 orchestrator — exercised at the function level (no subprocess)."""
import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture
def fake_case_root(tmp_path):
    root = tmp_path / "Smith v. Jones"
    (root / "RECORDS").mkdir(parents=True)
    (root / "PLEADINGS").mkdir()
    return root


def _make_text_source(path, content):
    path.write_text(content, encoding="utf-8")
    return path


def _mock_llm():
    """LLM mock that returns digest payload for extraction, topics payload for general."""
    caller = MagicMock()
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kwargs):
        if pass_name == "source_digest":
            return json.dumps({
                "source_id": "echo",
                "source_kind": "other",
                "deponent_statements": [],
                "factual_anchors": [{"fact": "x", "location": "p.1", "topic_tags": ["t"]}],
                "inconsistencies": [],
                "summary": "ok",
            })
        if pass_name == "topic_clustering":
            return json.dumps({"topics": [
                {"id": f"t{i:02d}", "title": f"Topic {i}", "strategic_note": "n",
                 "relevant_digest_refs": [], "default_checked": True}
                for i in range(1, 11)
            ]})
        return ""
    caller.call.side_effect = call
    return caller


def test_phase1_writes_session_and_topics(fake_case_root, tmp_path, capsys):
    from Scripts.depo_prep_lib import phase1

    src1 = _make_text_source(fake_case_root / "RECORDS" / "med.txt", "medical records text")
    src2 = _make_text_source(fake_case_root / "PLEADINGS" / "complaint.txt", "complaint text")

    config = {
        "deponent_name": "Jane Doe",
        "deponent_role": "Plaintiff",
        "deponent_sources": [str(src1)],
        "context_sources": [str(src2)],
        "style": "discovery",
        "free_text_notes": "Focus on causation.",
        "per_topic_flags": {
            "strategic_note": True, "source_facts": True,
            "impeachment_hook": False, "objection_alts": False,
        },
        "case_root": str(fake_case_root),
    }

    session_path = phase1.run_phase1(
        config=config,
        llm_caller=_mock_llm(),
        progress=lambda n, msg=None: None,
    )

    assert Path(session_path).exists()
    session_dir = Path(session_path).parent
    assert (session_dir / "topics.json").exists()
    topics = json.loads((session_dir / "topics.json").read_text(encoding="utf-8"))
    assert len(topics["topics"]) == 10

    digests_dir = session_dir / "digests"
    assert (digests_dir / "med.txt.json").exists()
    assert (digests_dir / "complaint.txt.json").exists()


def test_phase1_passes_extraction_task_type_to_llm(fake_case_root):
    from Scripts.depo_prep_lib import phase1

    src = _make_text_source(fake_case_root / "RECORDS" / "med.txt", "text")
    config = {
        "deponent_name": "J", "deponent_role": "P",
        "deponent_sources": [str(src)], "context_sources": [],
        "style": "discovery", "free_text_notes": "",
        "per_topic_flags": {"strategic_note": False, "source_facts": False,
                             "impeachment_hook": False, "objection_alts": False},
        "case_root": str(fake_case_root),
    }

    caller = _mock_llm()
    phase1.run_phase1(config=config, llm_caller=caller, progress=lambda *a, **kw: None)

    # At least one call must have been for source_digest with task_type="extraction".
    digest_calls = [c for c in caller.call.call_args_list
                    if c.kwargs.get("pass_name") == "source_digest"]
    assert digest_calls
    assert all(c.kwargs.get("task_type") == "extraction" for c in digest_calls)
