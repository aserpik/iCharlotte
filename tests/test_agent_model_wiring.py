"""Tests for get_primary_model_for_agent model resolution.

Some call sites take an explicit (provider, model) rather than going through
LLMCaller, so they resolve it from config via get_primary_model_for_agent. The
Workbench model-default therefore takes effect for them too.
"""

from icharlotte_core import llm_config as llm_config_module
from icharlotte_core.llm_config import LLMConfig, ModelSpec, get_primary_model_for_agent


def _reset_llm_config():
    LLMConfig._instance = None


def test_primary_model_prefers_configured_available_provider(tmp_path, monkeypatch):
    monkeypatch.setattr(llm_config_module, "CONFIG_FILE", str(tmp_path / "llm_preferences.json"))
    _reset_llm_config()
    try:
        config = LLMConfig()
        config.update_agent_config(
            "func_chat",
            model_sequence=[
                ModelSpec(provider="OpenAI", model="gpt-4o"),
                ModelSpec(provider="Gemini", model="gemini-3.5-flash"),
            ],
            use_default=False,
        )
        # Only Gemini has a key available -> skip OpenAI, pick the Gemini model.
        monkeypatch.setattr(
            LLMConfig, "is_provider_available",
            lambda self, provider: provider == "Gemini",
        )
        assert get_primary_model_for_agent("func_chat") == ("Gemini", "gemini-3.5-flash")
    finally:
        _reset_llm_config()


def test_primary_model_falls_back_to_default_when_unconfigured(tmp_path, monkeypatch):
    monkeypatch.setattr(llm_config_module, "CONFIG_FILE", str(tmp_path / "llm_preferences.json"))
    _reset_llm_config()
    try:
        # No task profiles, no agent config -> default tuple.
        monkeypatch.setattr(
            LLMConfig, "get_model_sequence_for_agent",
            lambda self, agent_id, fallback_task=None, pass_name=None: [],
        )
        assert get_primary_model_for_agent("nonexistent") == ("Gemini", "gemini-3.5-flash")
        assert get_primary_model_for_agent("nonexistent", default=("Claude", "x")) == ("Claude", "x")
    finally:
        _reset_llm_config()
