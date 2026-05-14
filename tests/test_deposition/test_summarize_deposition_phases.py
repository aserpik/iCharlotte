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
