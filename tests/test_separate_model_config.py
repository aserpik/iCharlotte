"""Tests that the Separate task is provider-agnostic via the shared LLMCaller.

The Separate task (Scripts/separate.py) no longer calls the Gemini SDK directly.
It routes through icharlotte_core.llm_config.LLMCaller, using agent
``agent_separate`` / pass ``main``, so the model (and provider) is whatever the
user configures in the Prompt Engineering Workbench — Gemini, Claude, or OpenAI,
with automatic fallback. These tests verify that wiring.
"""

import importlib.util
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SEPARATE_PATH = PROJECT_ROOT / "Scripts" / "separate.py"


@pytest.fixture(scope="module")
def separate_module():
    # Importing separate.py pulls in fitz/pypdf; skip cleanly if absent.
    pytest.importorskip("fitz")
    pytest.importorskip("pypdf")
    spec = importlib.util.spec_from_file_location("separate_under_test", SEPARATE_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(autouse=True)
def _reset_caller(separate_module):
    # Each test installs its own fake .call(); start from a clean caller.
    separate_module._llm_caller = None
    yield
    separate_module._llm_caller = None


def test_chunk_analysis_routes_through_llmcaller(separate_module, monkeypatch):
    captured = {}

    def fake_call(self, prompt, text, task_type="general", agent_id=None,
                  pass_name=None, **kwargs):
        captured.update(
            prompt=prompt, text=text, task_type=task_type,
            agent_id=agent_id, pass_name=pass_name,
        )
        return "1|Plaintiff's Complaint|2023-01-01|1|3\n2|Exhibit A|2023-02-02|4|6"

    monkeypatch.setattr(separate_module.LLMCaller, "call", fake_call)

    headers = ["Page 1: Complaint", "Page 2: ...", "Page 3: ...",
               "Page 4: Exhibit A", "Page 5: ...", "Page 6: ..."]
    docs = separate_module.analyze_headers_chunk(headers, start_page_num=1, next_id=1)

    # Goes through the Workbench-configurable agent/pass, not a hardcoded model.
    assert captured["agent_id"] == "agent_separate"
    assert captured["pass_name"] == "main"
    # Headers are passed as the document content, not embedded in the instruction.
    assert "\n".join(headers) == captured["text"]
    assert "HEADERS:" not in captured["prompt"]

    assert [d["title"] for d in docs] == ["Plaintiff's Complaint", "Exhibit A"]
    assert docs[0]["start"] == 1 and docs[0]["end"] == 3


def test_chunk_analysis_works_with_non_gemini_response(separate_module, monkeypatch):
    # Whatever provider LLMCaller routed to (e.g. Claude/OpenAI), the parsing is
    # provider-independent: it just consumes the returned text.
    def fake_call(self, *args, **kwargs):
        return "```\n1|Settlement Agreement|2024-05-01|1|10\n```"

    monkeypatch.setattr(separate_module.LLMCaller, "call", fake_call)

    docs = separate_module.analyze_headers_chunk(["Page 1: x"], start_page_num=1, next_id=1)

    assert len(docs) == 1
    assert docs[0]["title"] == "Settlement Agreement"


def test_chunk_analysis_returns_empty_when_all_models_fail(separate_module, monkeypatch):
    # LLMCaller.call returns None when every configured model/provider fails.
    monkeypatch.setattr(separate_module.LLMCaller, "call", lambda self, *a, **k: None)

    docs = separate_module.analyze_headers_chunk(["Page 1: x"], start_page_num=1, next_id=1)

    assert docs == []


def test_shipped_config_preserves_flash_default():
    import json

    cfg = json.loads((PROJECT_ROOT / "config" / "llm_preferences.json").read_text(encoding="utf-8"))
    sep = cfg["agents"]["agent_separate"]
    # Default (when no Workbench pass override is set) stays on the cheap Flash
    # models the task used before the refactor.
    assert sep["use_default"] is False
    seq = sep["model_sequence"]
    assert (seq[0]["provider"], seq[0]["model"]) == ("Gemini", "gemini-3.5-flash")
    assert (seq[1]["provider"], seq[1]["model"]) == ("Gemini", "gemini-3.1-flash-lite")


def test_source_no_longer_calls_gemini_sdk_directly():
    src = SEPARATE_PATH.read_text(encoding="utf-8")
    # The whole point of the refactor: no direct Gemini SDK coupling remains.
    assert "genai.Client" not in src
    assert "generate_content" not in src
    assert "from google import genai" not in src
    # And it does go through the shared caller.
    assert "LLMCaller" in src
    assert 'agent_id="agent_separate"' in src
    assert 'pass_name="main"' in src
