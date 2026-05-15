"""End-to-end smoke test: phase 1 → user config → phase 2 → output docx."""

import json
import os
import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from icharlotte_core.deposition import session_manager

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
AGENT_SCRIPT = PROJECT_ROOT / "Scripts" / "summarize_deposition.py"


@pytest.fixture
def mock_llm_env(tmp_path, monkeypatch):
    """Stand up a fake LLM via an environment variable that the agent honors in test mode.

    For this smoke test we rely on monkey-patching at the Python level rather than spawning
    a real subprocess — see the inline note. If you want a true subprocess test, wrap LLMCaller
    behavior behind an env var so the subprocess can pick it up.
    """
    return tmp_path


def test_full_flow_in_process(tmp_path, monkeypatch, capsys):
    """In-process smoke test of the phase 1 → phase 2 handoff.

    True subprocess test requires building an env-var-driven stub of LLMCaller; this
    test instead exercises main() in-process which still covers the session-JSON round-trip
    and the AWAITING_INPUT contract.
    """
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path / "sessions")

    sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))
    import summarize_deposition

    # Build a fake input pdf
    input_path = tmp_path / "Smith Depo.pdf"
    input_path.write_bytes(b"%PDF-1.4\n")

    # Stub extraction
    from types import SimpleNamespace
    fake_text = ("DEPOSITION OF JOHN SMITH\nTaken on January 15, 2024\n\n"
                 "Q. State your name.\nA. John Smith.\n") * 30

    def fake_extract(self, p):
        return SimpleNamespace(success=True, text=fake_text, char_count=len(fake_text),
                               page_count=10, ocr_pages=[], ocr_percentage=0.0, error=None)

    monkeypatch.setattr(summarize_deposition.DocumentProcessor, "extract_with_dynamic_ocr", fake_extract)
    monkeypatch.setattr(summarize_deposition, "save_to_docx",
                         lambda content, output_path, deponent, date, logger:
                             Path(output_path).parent.mkdir(parents=True, exist_ok=True) or
                             Path(output_path).write_text(content, encoding="utf-8") or True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None, raising=False)

    canned_topics = json.dumps([
        {"title": "Pre-Accident History", "rank": 1, "discussion_density": "high"},
        {"title": "Mechanism Of Injury", "rank": 2, "discussion_density": "high"},
    ])
    canned_summary = "**Pre-Accident History**\n- Bullet.\n\n**Mechanism Of Injury**\n- Bullet."

    call_responses = iter([canned_topics, canned_summary])

    def fake_call(self, prompt, text, task_type=None, **kw):
        return next(call_responses)

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    # Phase 1
    logger = summarize_deposition.AgentLogger("DepoSmoke", log_to_file=False)
    assert summarize_deposition.process_topics(str(input_path), logger) is True

    out = capsys.readouterr().out
    awaiting = [ln for ln in out.splitlines() if ln.startswith("AWAITING_INPUT:")]
    assert awaiting, "phase 1 did not emit AWAITING_INPUT"
    session_path = Path(awaiting[-1][len("AWAITING_INPUT:"):])
    assert session_path.exists()

    # Simulate the dialog writing user_config
    session_manager.update_user_config(session_path, {
        "selected_topics": ["Pre-Accident History", "Mechanism Of Injury"],
        "added_topics": [],
        "bullets_per_topic": 5,
        "deponent_label": "Plaintiff",
        "custom_rules": "",
        "cross_check_enabled": False,
        "bias": "neutral",
        "bias_custom": "",
        "context_doc_paths": [],
    })

    # Phase 2
    assert summarize_deposition.process_summary(str(session_path), logger) is True

    # Session intentionally kept alive so the user can generate additional versions.
    assert session_path.exists()
