"""Tests for webcompanion.chat — persistence wrappers + turn manager."""
import time

import pytest

from webcompanion import chat
from webcompanion import chat as chatmod
from icharlotte_core.chat import persistence as _persistence


@pytest.fixture
def data_dir(tmp_path, monkeypatch):
    monkeypatch.setattr(_persistence, "GEMINI_DATA_DIR", str(tmp_path))
    return tmp_path


def test_default_system_prompt_constant():
    assert "helpful legal assistant" in chat.DEFAULT_SYSTEM_PROMPT
    assert "disclaimers" in chat.DEFAULT_SYSTEM_PROMPT


def test_create_and_list_conversation(data_dir):
    conv_id = chat.create_conversation("9999", name="Test thread")
    convs = chat.list_conversations("9999")
    assert len(convs) == 1
    assert convs[0].id == conv_id and convs[0].name == "Test thread"
    # new conversations get the default legal-assistant system prompt
    assert "helpful legal assistant" in convs[0].system_prompt


def test_create_defaults_provider_model(data_dir):
    conv_id = chat.create_conversation("9999")
    conv = chat.get_conversation("9999", conv_id)
    assert conv.provider == "Gemini" and conv.model == "gemini-3.5-flash"


def test_get_missing_conversation_returns_none(data_dir):
    assert chat.get_conversation("9999", "nope") is None


def test_append_message_persists(data_dir):
    conv_id = chat.create_conversation("9999")
    chat.append_message("9999", conv_id, role="user", content="hello")
    conv = chat.get_conversation("9999", conv_id)
    assert len(conv.messages) == 1
    assert conv.messages[0].role == "user" and conv.messages[0].content == "hello"


def test_append_assistant_message_records_model(data_dir):
    conv_id = chat.create_conversation("9999")
    chat.append_message("9999", conv_id, role="assistant", content="hi",
                        model_used="gemini-3.5-flash")
    conv = chat.get_conversation("9999", conv_id)
    assert conv.messages[0].model_used == "gemini-3.5-flash"


def _wait_turn(mgr, turn_id, status, timeout=5.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        t = mgr.get_turn(turn_id)
        if t and t["status"] == status:
            return t
        time.sleep(0.02)
    raise AssertionError(f"turn did not reach {status}: {mgr.get_turn(turn_id)}")


def test_turn_generates_and_persists(data_dir, monkeypatch):
    monkeypatch.setattr(chatmod.LLMHandler, "generate",
                        staticmethod(lambda **kw: "the answer"))
    conv_id = chat.create_conversation("9999")
    mgr = chatmod.ChatTurnManager()
    turn_id = mgr.start_turn("9999", conv_id, user_text="what is the law?",
                             provider="Gemini", model="gemini-3.5-flash",
                             attach_rel_files=[], research_on=False)
    # user message persisted immediately
    conv = chat.get_conversation("9999", conv_id)
    assert conv.messages[0].role == "user"
    _wait_turn(mgr, turn_id, "done")
    conv = chat.get_conversation("9999", conv_id)
    assert [m.role for m in conv.messages] == ["user", "assistant"]
    assert conv.messages[1].content == "the answer"
    assert conv.messages[1].model_used == "gemini-3.5-flash"


def test_turn_passes_history_excluding_current(data_dir, monkeypatch):
    captured = {}
    def fake_generate(**kw):
        captured.update(kw)
        return "reply2"
    monkeypatch.setattr(chatmod.LLMHandler, "generate", staticmethod(fake_generate))
    conv_id = chat.create_conversation("9999")
    chat.append_message("9999", conv_id, role="user", content="first")
    chat.append_message("9999", conv_id, role="assistant", content="answer1")
    mgr = chatmod.ChatTurnManager()
    turn_id = mgr.start_turn("9999", conv_id, user_text="second",
                             provider="Gemini", model="gemini-3.5-flash",
                             attach_rel_files=[], research_on=False)
    _wait_turn(mgr, turn_id, "done")
    # history is the prior 2 messages; current user_text passed separately
    assert captured["history"] == [
        {"role": "user", "content": "first"},
        {"role": "assistant", "content": "answer1"},
    ]
    assert captured["user_prompt"] == "second"


def test_turn_generate_failure_marks_failed(data_dir, monkeypatch):
    def boom(**kw):
        raise RuntimeError("llm down")
    monkeypatch.setattr(chatmod.LLMHandler, "generate", staticmethod(boom))
    conv_id = chat.create_conversation("9999")
    mgr = chatmod.ChatTurnManager()
    turn_id = mgr.start_turn("9999", conv_id, user_text="x",
                             provider="Gemini", model="gemini-3.5-flash",
                             attach_rel_files=[], research_on=False)
    t = _wait_turn(mgr, turn_id, "failed")
    assert "llm down" in t["error"]
    # user message stays; no assistant message appended
    conv = chat.get_conversation("9999", conv_id)
    assert [m.role for m in conv.messages] == ["user"]


def test_one_turn_per_conversation(data_dir, monkeypatch):
    monkeypatch.setattr(chatmod.LLMHandler, "generate",
                        staticmethod(lambda **kw: (time.sleep(0.3) or "slow")))
    conv_id = chat.create_conversation("9999")
    mgr = chatmod.ChatTurnManager()
    first = mgr.start_turn("9999", conv_id, user_text="a", provider="Gemini",
                           model="gemini-3.5-flash", attach_rel_files=[], research_on=False)
    with pytest.raises(ValueError):
        mgr.start_turn("9999", conv_id, user_text="b", provider="Gemini",
                       model="gemini-3.5-flash", attach_rel_files=[],
                       research_on=False)
    _wait_turn(mgr, first, "done")  # let the first turn finish before teardown


def test_conversation_frees_after_successful_turn(data_dir, monkeypatch):
    monkeypatch.setattr(chatmod.LLMHandler, "generate",
                        staticmethod(lambda **kw: "ok"))
    conv_id = chat.create_conversation("9999")
    mgr = chatmod.ChatTurnManager()
    t1 = mgr.start_turn("9999", conv_id, user_text="first", provider="Gemini",
                        model="gemini-3.5-flash", attach_rel_files=[], research_on=False)
    _wait_turn(mgr, t1, "done")
    # conversation is no longer busy -> a second turn starts without ValueError
    t2 = mgr.start_turn("9999", conv_id, user_text="second", provider="Gemini",
                        model="gemini-3.5-flash", attach_rel_files=[], research_on=False)
    _wait_turn(mgr, t2, "done")
    conv = chat.get_conversation("9999", conv_id)
    assert [m.role for m in conv.messages] == ["user", "assistant", "user", "assistant"]


def test_persist_model_choice_updates_conversation(data_dir):
    conv_id = chat.create_conversation("9999")  # defaults Gemini/flash
    chat._persist_model_choice("9999", conv_id, "Claude",
                               "claude-sonnet-4-20250514")
    conv = chat.get_conversation("9999", conv_id)
    assert conv.provider == "Claude" and conv.model == "claude-sonnet-4-20250514"


def test_extract_context_concatenates_and_caps(data_dir, monkeypatch, tmp_path):
    f1 = tmp_path / "a.txt"; f1.write_text("AAA", encoding="utf-8")
    f2 = tmp_path / "b.txt"; f2.write_text("BBB", encoding="utf-8")

    class _Res:
        def __init__(self, text): self.text = text; self.error = ""
    monkeypatch.setattr(chatmod.DocumentProcessor, "extract_text",
                        lambda self, p, ocr_enabled=True: _Res(open(p, encoding="utf-8").read()))
    mgr = chatmod.ChatTurnManager()
    text = mgr._extract_context("t", str(tmp_path), ["a.txt", "b.txt"])
    assert "AAA" in text and "BBB" in text
    assert len(text) <= chatmod._CONTEXT_CHAR_CAP


def test_extract_context_skips_failed_file(data_dir, monkeypatch, tmp_path):
    (tmp_path / "ok.txt").write_text("OK", encoding="utf-8")
    class _Res:
        def __init__(self, text, error=""): self.text = text; self.error = error
    def fake_extract(self, p, ocr_enabled=True):
        if p.endswith("bad.txt"):
            return _Res("", "boom")
        return _Res("OK")
    monkeypatch.setattr(chatmod.DocumentProcessor, "extract_text", fake_extract)
    mgr = chatmod.ChatTurnManager()
    text = mgr._extract_context("t", str(tmp_path), ["ok.txt", "bad.txt"])
    assert "OK" in text  # bad file skipped, no raise


def test_extract_context_empty_when_none(data_dir):
    mgr = chatmod.ChatTurnManager()
    assert mgr._extract_context("t", "/case", []) == ""
