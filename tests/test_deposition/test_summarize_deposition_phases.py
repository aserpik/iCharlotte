"""Tests for the two-phase summarize_deposition agent."""

import json
import os
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest

# Make Scripts/ importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import summarize_deposition  # noqa: E402
from icharlotte_core.deposition import session_manager  # noqa: E402


FAKE_TRANSCRIPT = (
    "DEPOSITION OF JOHN SMITH\n"
    "Taken on January 15, 2024\n\n"
    "Q. Please state your name.\n"
    "A. John Smith.\n"
    "Q. Where do you live?\n"
    "A. Riverside, California.\n"
) * 50


def _stub_extract_with_dynamic_ocr(self, path):
    return SimpleNamespace(
        success=True,
        text=FAKE_TRANSCRIPT,
        char_count=len(FAKE_TRANSCRIPT),
        page_count=20,
        ocr_pages=[],
        ocr_percentage=0.0,
        error=None,
    )


@pytest.fixture
def isolated_sessions(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path / "sessions")
    return tmp_path


def test_phase1_writes_session_json_and_caches_text(isolated_sessions, capsys, monkeypatch):
    canned_topics = json.dumps([
        {"title": "Pre-Accident Medical History", "rank": 1, "discussion_density": "high"},
        {"title": "Mechanism Of Injury", "rank": 2, "discussion_density": "high"},
    ])

    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        _stub_extract_with_dynamic_ocr,
    )

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=canned_topics):
        input_path = str(isolated_sessions / "Smith Depo.pdf")
        Path(input_path).write_bytes(b"%PDF-1.4\n")  # presence only; extractor is stubbed
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        result = summarize_deposition.process_topics(input_path, logger)

    assert result is True
    out = capsys.readouterr().out
    awaiting_lines = [ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")]
    assert awaiting_lines, "AWAITING_INPUT token not printed"
    session_path = Path(awaiting_lines[-1][len("AWAITING_INPUT:"):])
    assert session_path.exists()

    data = json.loads(session_path.read_text(encoding="utf-8"))
    assert data["phase"] == "awaiting_input"
    assert data["user_config"] is None
    assert len(data["topics"]) == 2
    assert data["topics"][0]["title"] == "Pre-Accident Medical History"
    assert Path(data["cached_text_path"]).exists()
    assert Path(data["cached_text_path"]).read_text(encoding="utf-8") == FAKE_TRANSCRIPT


def test_phase1_handles_malformed_llm_json(isolated_sessions, capsys, monkeypatch):
    # Bulleted list instead of JSON
    malformed = "- Topic A\n- Topic B\n- Topic C\n"

    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        _stub_extract_with_dynamic_ocr,
    )

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=malformed):
        input_path = str(isolated_sessions / "X.pdf")
        Path(input_path).write_bytes(b"%PDF-1.4\n")
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        result = summarize_deposition.process_topics(input_path, logger)

    assert result is True
    out = capsys.readouterr().out
    awaiting_lines = [ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")]
    session_path = Path(awaiting_lines[-1][len("AWAITING_INPUT:"):])
    data = json.loads(session_path.read_text(encoding="utf-8"))
    titles = [t["title"] for t in data["topics"]]
    assert "Topic A" in titles
    assert "Topic B" in titles
    assert "Topic C" in titles


def test_phase1_caps_topic_count_at_25(isolated_sessions, capsys, monkeypatch):
    many = json.dumps([
        {"title": f"Topic {i}", "rank": i, "discussion_density": "medium"}
        for i in range(1, 51)
    ])

    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        _stub_extract_with_dynamic_ocr,
    )

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=many):
        input_path = str(isolated_sessions / "Y.pdf")
        Path(input_path).write_bytes(b"%PDF-1.4\n")
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        summarize_deposition.process_topics(input_path, logger)

    out = capsys.readouterr().out
    session_path = Path([ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")][-1][len("AWAITING_INPUT:"):])
    data = json.loads(session_path.read_text(encoding="utf-8"))
    assert len(data["topics"]) == 25
    # Truncation keeps the top-ranked topics
    assert data["topics"][0]["title"] == "Topic 1"
    assert data["topics"][-1]["title"] == "Topic 25"


# ---------------------------------------------------------------------------
# Phase 2
# ---------------------------------------------------------------------------

def _write_ready_session(tmp_path, *, cross_check, selected, added, bullets=5, label="Plaintiff", rules=""):
    session_path = tmp_path / "session.json"
    cached_path = tmp_path / "session.txt"
    cached_path.write_text(FAKE_TRANSCRIPT, encoding="utf-8")
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "ready_for_summary",
        "input_path": str(tmp_path / "Smith Depo.pdf"),
        "cached_text_path": str(cached_path),
        "deponent_name": "John Smith",
        "deposition_date": "January 15, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "3850.084",
        "topics": [{"id": 1, "title": "Topic A", "rank": 1, "discussion_density": "high"}],
        "user_config": {
            "selected_topics": selected,
            "added_topics": added,
            "bullets_per_topic": bullets,
            "deponent_label": label,
            "custom_rules": rules,
            "cross_check_enabled": cross_check,
            "bias": "neutral",
            "bias_custom": "",
            "context_doc_paths": [],
        },
    })
    return session_path


def test_phase2_reads_session_and_generates_summary(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False,
        selected=["Topic A"], added=["Topic B"],
    )

    canned_summary = "**Topic A**\n- Bullet about A.\n\n**Topic B**\n- Bullet about B."
    called_with = {}

    def fake_save(content, output_path, deponent, date, logger):
        called_with["content"] = content
        called_with["output_path"] = output_path
        return True

    monkeypatch.setattr(summarize_deposition, "save_to_docx", fake_save)
    # Stub out registry / case-data writes — they're orthogonal to this test.
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=canned_summary) as mock_call:
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        ok = summarize_deposition.process_summary(str(session_path), logger)

    assert ok is True
    assert mock_call.call_count == 1  # cross-check disabled
    assert called_with["content"] == canned_summary
    # Session + cached text cleaned up on success
    assert not session_path.exists()


def test_phase2_cross_check_runs_only_when_enabled(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=True,
        selected=["Topic A"], added=[],
    )

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    long_enough = "**Topic A**\n" + ("- Filler bullet.\n" * 20)

    with patch.object(summarize_deposition.LLMCaller, "call", return_value=long_enough) as mock_call:
        logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
        summarize_deposition.process_summary(str(session_path), logger)

    assert mock_call.call_count == 2  # summary + cross-check


def test_phase2_prompt_includes_all_selected_plus_added_topics_in_order(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False,
        selected=["Pre-Accident History", "Mechanism Of Injury"],
        added=["Communications With Providers"],
        bullets=7,
        label="Mr. Smith",
        rules="Use past tense.",
    )

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        captured["text"] = text
        return "**Pre-Accident History**\n- B.\n\n**Mechanism Of Injury**\n- B.\n\n**Communications With Providers**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
    summarize_deposition.process_summary(str(session_path), logger)

    prompt = captured["prompt"]
    # All three topics appear in order
    a = prompt.find("Pre-Accident History")
    b = prompt.find("Mechanism Of Injury")
    c = prompt.find("Communications With Providers")
    assert 0 < a < b < c
    assert "Mr. Smith" in prompt
    assert "7" in prompt  # bullets per topic
    assert "Use past tense." in prompt


def test_phase2_prompt_substitution_strips_user_braces(tmp_path, monkeypatch):
    """User-typed {placeholder} text in deponent_label or custom_rules must not leak across slots."""
    session_path = _write_ready_session(
        tmp_path, cross_check=False,
        selected=["Topic A"], added=[],
        bullets=5,
        label="Mr. {topic_list} Smith",  # adversarial input
        rules="Avoid {bullets_per_topic} or {deponent_label} references.",
    )

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        return "**Topic A**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
    summarize_deposition.process_summary(str(session_path), logger)

    prompt = captured["prompt"]
    # Neither the literal placeholder strings nor the curly braces from user input
    # should remain in the final prompt.
    assert "{topic_list}" not in prompt
    assert "{bullets_per_topic}" not in prompt
    assert "{deponent_label}" not in prompt
    assert "{custom_rules}" not in prompt
    # The braces from the user input were stripped, so we expect:
    assert "Mr. topic_list Smith" in prompt
    assert "Avoid bullets_per_topic or deponent_label references." in prompt


@pytest.mark.parametrize("bias_value,bias_custom,expected_substring", [
    ("neutral", "", "neutral, balanced tone"),
    ("pro_plaintiff", "", "most favorable to the plaintiff"),
    ("pro_defense", "", "most favorable to the defense"),
    ("custom", "Highlight inconsistencies in injury testimony.",
     "Highlight inconsistencies in injury testimony."),
])
def test_phase2_resolves_bias_directive_for_each_preset(
    tmp_path, monkeypatch, bias_value, bias_custom, expected_substring
):
    """Each bias preset routes the expected directive language into the prompt."""
    session_path = tmp_path / "session.json"
    cached_path = tmp_path / "session.txt"
    cached_path.write_text(FAKE_TRANSCRIPT, encoding="utf-8")
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "ready_for_summary",
        "input_path": str(tmp_path / "X.pdf"),
        "cached_text_path": str(cached_path),
        "deponent_name": "X",
        "deposition_date": "Jan 1, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "0000.000",
        "topics": [{"id": 1, "title": "T", "rank": 1, "discussion_density": "high"}],
        "user_config": {
            "selected_topics": ["T"],
            "added_topics": [],
            "bullets_per_topic": 5,
            "deponent_label": "Plaintiff",
            "custom_rules": "",
            "cross_check_enabled": False,
            "bias": bias_value,
            "bias_custom": bias_custom,
            "context_doc_paths": [],
        },
    })

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)

    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        return "**T**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    logger = summarize_deposition.AgentLogger("BiasTest", log_to_file=False)
    summarize_deposition.process_summary(str(session_path), logger)
    assert expected_substring in captured["prompt"]
