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


def test_phase1_chunks_large_transcript_for_topic_discovery(isolated_sessions, capsys, monkeypatch):
    large_transcript = FAKE_TRANSCRIPT + (
        "\nQ. Please describe the incident.\n"
        "A. I was asked about treatment, pain, work limits, and prior injuries.\n"
    ) * 2200

    def fake_extract(self, path):
        return SimpleNamespace(
            success=True,
            text=large_transcript,
            char_count=len(large_transcript),
            page_count=130,
            ocr_pages=[],
            ocr_percentage=0.0,
            error=None,
        )

    calls = []

    def fake_call(self, prompt, text, task_type=None, agent_id=None, pass_name=None, **kwargs):
        calls.append({
            "prompt": prompt,
            "text_len": len(text),
            "task_type": task_type,
            "agent_id": agent_id,
            "pass_name": pass_name,
        })
        if len(text) > 95_000:
            return None
        if "candidate topic lists" in prompt.lower():
            return json.dumps([
                {"title": "Incident Description", "rank": 1, "discussion_density": "high"},
                {"title": "Medical Treatment", "rank": 2, "discussion_density": "high"},
            ])
        return json.dumps([
            {"title": "Incident Description", "rank": 1, "discussion_density": "high"},
            {"title": "Medical Treatment", "rank": 2, "discussion_density": "medium"},
        ])

    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        fake_extract,
    )
    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    input_path = str(isolated_sessions / "Long Depo.pdf")
    Path(input_path).write_bytes(b"%PDF-1.4\n")
    logger = summarize_deposition.AgentLogger("DepositionTest", log_to_file=False)
    result = summarize_deposition.process_topics(input_path, logger)

    assert result is True
    discovery_calls = [
        c for c in calls
        if c["agent_id"] == "agent_sum_depo" and c["pass_name"] == "topic_discovery"
    ]
    assert len(discovery_calls) >= 2
    assert all(c["text_len"] <= 95_000 for c in discovery_calls)

    out = capsys.readouterr().out
    session_path = Path([ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")][-1][len("AWAITING_INPUT:"):])
    data = json.loads(session_path.read_text(encoding="utf-8"))
    assert [t["title"] for t in data["topics"]][:2] == [
        "Incident Description",
        "Medical Treatment",
    ]


def test_deposition_chunk_caps_limit_known_large_transcripts_to_two_calls():
    block = (
        "\nQ. Please describe the incident, medical treatment, symptoms, "
        "work limits, and damages.\n"
        "A. The witness gave detailed testimony about those topics.\n"
    )
    transcript = FAKE_TRANSCRIPT + (block * 1140)
    assert 166_000 < len(transcript) < 180_000

    topic_chunks = summarize_deposition._split_topic_discovery_text(transcript)
    summary_chunks = summarize_deposition._split_topic_discovery_text(
        transcript,
        max_chars=summarize_deposition._DEPOSITION_SUMMARY_CHUNK_CHAR_CAP,
    )

    assert len(topic_chunks) == 2
    assert len(summary_chunks) == 2
    assert max(len(chunk) for chunk in topic_chunks) <= 95_000
    assert max(len(chunk) for chunk in summary_chunks) <= 95_000


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
            "audience": "neutral",
            "audience_custom": "",
            "tone": "recitation",
            "tone_custom": "",
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
    # Session is intentionally kept alive after success — see test_phase2_keeps_session_alive_after_success.
    assert session_path.exists()


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


def _write_session_with_posture(tmp_path, audience, audience_custom="",
                                  tone="recitation", tone_custom="",
                                  legacy_bias=None, legacy_bias_custom=""):
    """Build a ready_for_summary session with arbitrary audience/tone keys."""
    session_path = tmp_path / "session.json"
    cached_path = tmp_path / "session.txt"
    cached_path.write_text(FAKE_TRANSCRIPT, encoding="utf-8")
    user_config = {
        "selected_topics": ["T"],
        "added_topics": [],
        "bullets_per_topic": 5,
        "deponent_label": "Plaintiff",
        "custom_rules": "",
        "cross_check_enabled": False,
        "context_doc_paths": [],
    }
    if legacy_bias is not None:
        # Legacy session: old "bias" keys only, no "audience" key at all.
        user_config["bias"] = legacy_bias
        user_config["bias_custom"] = legacy_bias_custom
    else:
        user_config["audience"] = audience
        user_config["audience_custom"] = audience_custom
        user_config["tone"] = tone
        user_config["tone_custom"] = tone_custom
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
        "user_config": user_config,
    })
    return session_path


def _run_and_capture_prompt(session_path, monkeypatch):
    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        return "**T**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)
    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("PostureTest", log_to_file=False),
    )
    return captured["prompt"]


@pytest.mark.parametrize("audience,expected_phrase", [
    ("neutral", "balanced, work-product summary suitable for either side"),
    ("pro_plaintiff", "summarizing this transcript for plaintiff's trial team"),
    ("pro_defense", "summarizing this transcript for defense's trial team"),
])
def test_phase2_audience_preset_inserts_expected_directive(
    tmp_path, monkeypatch, audience, expected_phrase
):
    """Each audience preset routes the expected directive language into the prompt."""
    session_path = _write_session_with_posture(tmp_path, audience)
    prompt = _run_and_capture_prompt(session_path, monkeypatch)
    assert expected_phrase in prompt


def test_phase2_audience_custom_uses_user_text(tmp_path, monkeypatch):
    session_path = _write_session_with_posture(
        tmp_path, "custom", audience_custom="Insurance carrier evaluating settlement value.",
    )
    prompt = _run_and_capture_prompt(session_path, monkeypatch)
    assert "Insurance carrier evaluating settlement value." in prompt


def test_phase2_audience_custom_falls_back_to_neutral_when_blank(tmp_path, monkeypatch):
    """If 'custom' is picked but the field is empty, fall back to neutral wording."""
    session_path = _write_session_with_posture(tmp_path, "custom", audience_custom="")
    prompt = _run_and_capture_prompt(session_path, monkeypatch)
    assert "balanced, work-product summary suitable for either side" in prompt


@pytest.mark.parametrize("tone,expected_phrase", [
    ("recitation", "Recite the testimony without characterization"),
    ("editorial", "you may briefly characterize testimony"),
])
def test_phase2_tone_preset_inserts_expected_directive(
    tmp_path, monkeypatch, tone, expected_phrase
):
    session_path = _write_session_with_posture(tmp_path, "neutral", tone=tone)
    prompt = _run_and_capture_prompt(session_path, monkeypatch)
    assert expected_phrase in prompt


def test_phase2_tone_custom_uses_user_text(tmp_path, monkeypatch):
    session_path = _write_session_with_posture(
        tmp_path, "neutral", tone="custom",
        tone_custom="Use plain-English jury-presentation style.",
    )
    prompt = _run_and_capture_prompt(session_path, monkeypatch)
    assert "Use plain-English jury-presentation style." in prompt


def test_phase2_legacy_bias_key_still_maps_to_audience_directive(tmp_path, monkeypatch):
    """Sessions written before the rename use 'bias' instead of 'audience' — they
    must still route to the same directive without explosion."""
    session_path = _write_session_with_posture(
        tmp_path, audience=None, legacy_bias="pro_plaintiff",
    )
    prompt = _run_and_capture_prompt(session_path, monkeypatch)
    assert "summarizing this transcript for plaintiff's trial team" in prompt
    # Legacy sessions have no tone key — should default to recitation directive.
    assert "Recite the testimony without characterization" in prompt


def test_phase2_concatenates_context_documents_into_prompt(tmp_path, monkeypatch):
    """Two context docs (one PDF, one DOCX) are extracted and injected into the prompt."""
    pdf_path = tmp_path / "complaint.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")
    docx_path = tmp_path / "med.docx"
    docx_path.write_bytes(b"PK\x03\x04")

    # Mock both extractors
    from types import SimpleNamespace
    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        lambda self, p: SimpleNamespace(
            success=True, text="PDF context text", char_count=16, page_count=1,
            ocr_pages=[], ocr_percentage=0.0, error=None,
        ),
    )
    monkeypatch.setattr(
        "icharlotte_core.document_processor.extract_docx_text",
        lambda p: "DOCX context text",
    )

    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(pdf_path), str(docx_path)]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)

    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        return "**T**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    logger = summarize_deposition.AgentLogger("CtxTest", log_to_file=False)
    summarize_deposition.process_summary(str(session_path), logger)

    p = captured["prompt"]
    assert "=== CONTEXT DOC: complaint.pdf ===" in p
    assert "PDF context text" in p
    assert "=== CONTEXT DOC: med.docx ===" in p
    assert "DOCX context text" in p
    # Context section must be a directive ("use as a roadmap"), not a soft hint,
    # so the model is steered to use status-report theories to drive bullet
    # selection. Otherwise the context doc has weak influence on the output.
    assert "USE AS A ROADMAP FOR BULLET SELECTION" in p
    assert "Lead each topic with bullets that confirm, contradict, or undermine" in p


def test_phase2_per_doc_char_cap_truncates_long_docs(tmp_path, monkeypatch):
    pdf_path = tmp_path / "big.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")

    huge = "X" * 200_000
    from types import SimpleNamespace
    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        lambda self, p: SimpleNamespace(
            success=True, text=huge, char_count=len(huge), page_count=100,
            ocr_pages=[], ocr_percentage=0.0, error=None,
        ),
    )

    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(pdf_path)]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)

    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("TruncTest", log_to_file=False),
    )
    assert "[...truncated at 100000 chars]" in captured["prompt"]


def test_phase2_missing_context_doc_logged_and_skipped(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(tmp_path / "ghost.pdf")]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    logger = summarize_deposition.AgentLogger("MissTest", log_to_file=False)
    ok = summarize_deposition.process_summary(str(session_path), logger)
    assert ok is True
    assert "=== CONTEXT DOC:" not in captured["prompt"]


def test_phase2_no_context_docs_leaves_context_section_empty(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("EmptyCtxTest", log_to_file=False),
    )
    # The injected context block has a distinctive lead-in; rule 9 of the prompt
    # template mentions "ADDITIONAL CASE CONTEXT" by name, so we look for the
    # block-specific phrase instead.
    assert "USE AS A ROADMAP FOR BULLET SELECTION" not in captured["prompt"]
    assert "=== CONTEXT DOC:" not in captured["prompt"]


def test_phase2_doc_extension_unsupported_skipped(tmp_path, monkeypatch):
    txt_path = tmp_path / "notes.txt"
    txt_path.write_text("not a supported doc", encoding="utf-8")

    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(txt_path)]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    ok = summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("ExtTest", log_to_file=False),
    )
    assert ok is True
    assert "=== CONTEXT DOC:" not in captured["prompt"]


def test_phase2_default_output_path_uses_deponent_name(tmp_path, monkeypatch):
    """Without --output_path, output filename is derived from the deponent name."""
    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    captured = {}

    def fake_save(content, output_path, deponent, date, logger):
        captured["output_path"] = output_path
        return True

    monkeypatch.setattr(summarize_deposition, "save_to_docx", fake_save)
    monkeypatch.setattr(summarize_deposition, "_register_outputs",
                        lambda *a, **kw: None, raising=False)
    # get_output_directory walks the input path; point it at tmp.
    monkeypatch.setattr(summarize_deposition, "get_output_directory",
                        lambda p: str(tmp_path / "AI OUTPUT"))
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: "**T**\n- B.",
    )

    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("OutNameTest", log_to_file=False),
    )

    # Session was built with deponent_name="John Smith".
    assert os.path.basename(captured["output_path"]) == "Deposition of John Smith.docx"


def test_phase2_default_output_falls_back_to_input_stem(tmp_path, monkeypatch):
    """If deponent_name is empty, the filename derives from the input file stem."""
    session_path = tmp_path / "session.json"
    cached_path = tmp_path / "session.txt"
    cached_path.write_text(FAKE_TRANSCRIPT, encoding="utf-8")
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "ready_for_summary",
        "input_path": str(tmp_path / "Smith Depo.pdf"),
        "cached_text_path": str(cached_path),
        "deponent_name": "",  # missing
        "deposition_date": "",
        "deponent_type": "Witness",
        "file_number": "0000.000",
        "topics": [{"id": 1, "title": "T", "rank": 1, "discussion_density": "high"}],
        "user_config": {
            "selected_topics": ["T"], "added_topics": [], "bullets_per_topic": 5,
            "deponent_label": "Witness", "custom_rules": "",
            "cross_check_enabled": False, "bias": "neutral", "bias_custom": "",
            "context_doc_paths": [],
        },
    })

    captured = {}
    monkeypatch.setattr(summarize_deposition, "save_to_docx",
                        lambda c, p, d, dt, lg: captured.update(output_path=p) or True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs",
                        lambda *a, **kw: None, raising=False)
    monkeypatch.setattr(summarize_deposition, "get_output_directory",
                        lambda p: str(tmp_path / "AI OUTPUT"))
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: "**T**\n- B.",
    )

    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("FallbackNameTest", log_to_file=False),
    )

    assert os.path.basename(captured["output_path"]) == "Smith Depo - Summary.docx"


def test_phase2_explicit_output_path_override_is_respected(tmp_path, monkeypatch):
    """--output_path takes precedence over the per-deponent default."""
    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    captured = {}
    monkeypatch.setattr(summarize_deposition, "save_to_docx",
                        lambda c, p, d, dt, lg: captured.update(output_path=p) or True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs",
                        lambda *a, **kw: None, raising=False)
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: "**T**\n- B.",
    )

    override = str(tmp_path / "explicit_target.docx")
    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("OverrideTest", log_to_file=False),
        output_path_override=override,
    )
    assert captured["output_path"] == override


def test_save_to_docx_versions_up_when_file_exists(tmp_path):
    """Re-running for the same deponent should create v.2 rather than appending."""
    target = tmp_path / "Deposition of John Smith.docx"
    # Pre-existing file simulating a prior run.
    target.write_bytes(b"placeholder")
    pre_size = target.stat().st_size

    logger = summarize_deposition.AgentLogger("VersionUpTest", log_to_file=False)
    ok = summarize_deposition.save_to_docx(
        "**Topic**\n- Bullet.", str(target),
        "John Smith", "January 15, 2024", logger,
    )

    assert ok is True
    # Original file untouched.
    assert target.exists()
    assert target.stat().st_size == pre_size
    # v.2 created alongside.
    v2 = tmp_path / "Deposition of John Smith v.2.docx"
    assert v2.exists()
    assert v2.stat().st_size > 0


def test_phase2_keeps_session_alive_after_success(tmp_path, monkeypatch):
    """After Phase 2 succeeds, the session JSON and cached transcript remain on disk
    so the user can re-open the popup to generate another version."""
    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    cached_path = Path(session_manager.read_session(session_path)["cached_text_path"])

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: "**T**\n- B.",
    )

    logger = summarize_deposition.AgentLogger("KeepAliveTest", log_to_file=False)
    ok = summarize_deposition.process_summary(str(session_path), logger)
    assert ok is True
    # Session + cached text MUST still exist after a successful phase 2.
    assert session_path.exists()
    assert cached_path.exists()


def test_phase2_chunks_large_cached_transcript_for_summary(tmp_path, monkeypatch):
    large_text = FAKE_TRANSCRIPT + (
        "\nQ. Tell me about the incident.\n"
        "A. We covered liability, medical treatment, symptoms, work limits, and damages.\n"
    ) * 1500
    session_path = _write_ready_session(
        tmp_path,
        cross_check=False,
        selected=["Incident Description", "Medical Treatment"],
        added=[],
    )
    session = session_manager.read_session(session_path)
    Path(session["cached_text_path"]).write_text(large_text, encoding="utf-8")

    calls = []

    def fake_call(self, prompt, text, task_type=None, agent_id=None, pass_name=None, **kwargs):
        calls.append({
            "prompt": prompt,
            "text_len": len(text),
            "task_type": task_type,
            "agent_id": agent_id,
            "pass_name": pass_name,
        })
        if text and len(text) > 95_000:
            return None
        if "partial deposition summaries" in prompt.lower():
            return "**Incident Description**\n- Final consolidated point."
        return "**Incident Description**\n- Chunk point."

    saved = {}

    def fake_save(summary, output_path, deponent_name, deposition_date, logger):
        saved["summary"] = summary
        saved["output_path"] = output_path
        return True

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)
    monkeypatch.setattr(summarize_deposition, "save_to_docx", fake_save)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)

    ok = summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("Phase2ChunkTest", log_to_file=False),
    )

    assert ok is True
    summary_calls = [
        c for c in calls
        if c["agent_id"] == "agent_sum_depo" and c["pass_name"] == "summary"
    ]
    assert len(summary_calls) >= 2
    assert all(c["text_len"] <= 95_000 for c in summary_calls)
    assert saved["summary"] == "**Incident Description**\n- Final consolidated point."
