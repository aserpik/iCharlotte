"""Tests for webcompanion.chat_models."""
from webcompanion import chat_models


def test_available_models_nonempty_and_shaped():
    models = chat_models.available_models()
    assert len(models) >= 2
    for m in models:
        assert set(m) >= {"provider", "model", "label"}
    # includes the chat default
    assert any(m["model"] == "gemini-3.5-flash" for m in models)


def test_available_models_deduped():
    models = chat_models.available_models()
    pairs = [(m["provider"], m["model"]) for m in models]
    assert len(pairs) == len(set(pairs))
