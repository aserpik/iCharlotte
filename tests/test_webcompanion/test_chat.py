"""Tests for webcompanion.chat — persistence wrappers + turn manager."""
import pytest

from webcompanion import chat
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
