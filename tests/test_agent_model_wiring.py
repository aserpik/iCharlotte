"""Tests for get_primary_model_for_agent and the email/sent-monitor wiring.

These features take an explicit (provider, model) rather than going through
LLMCaller, so they resolve it from config via get_primary_model_for_agent. The
Workbench model-default therefore takes effect for them too.
"""

import ast
from pathlib import Path

from icharlotte_core import llm_config as llm_config_module
from icharlotte_core.llm_config import LLMConfig, ModelSpec, get_primary_model_for_agent

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _reset_llm_config():
    LLMConfig._instance = None


def test_primary_model_prefers_configured_available_provider(tmp_path, monkeypatch):
    monkeypatch.setattr(llm_config_module, "CONFIG_FILE", str(tmp_path / "llm_preferences.json"))
    _reset_llm_config()
    try:
        config = LLMConfig()
        config.update_agent_config(
            "func_email_compose",
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
        assert get_primary_model_for_agent("func_email_compose") == ("Gemini", "gemini-3.5-flash")
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


def _calls_with_agent(source_path, func_substr):
    """Return True if the source calls get_primary_model_for_agent with the given agent id."""
    tree = ast.parse(source_path.read_text(encoding="utf-8"), filename=str(source_path))
    for node in ast.walk(tree):
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "get_primary_model_for_agent"
            and node.args
            and isinstance(node.args[0], ast.Constant)
            and node.args[0].value == func_substr
        ):
            return True
    return False


def test_email_and_sent_monitor_resolve_models_from_config():
    assert _calls_with_agent(
        PROJECT_ROOT / "icharlotte_core" / "ui" / "email_update_tab.py", "func_email_compose"
    )
    assert _calls_with_agent(
        PROJECT_ROOT / "icharlotte_core" / "sent_items_monitor.py", "func_sent_monitor"
    )
    assert _calls_with_agent(
        PROJECT_ROOT / "icharlotte_core" / "ui" / "email_tab.py", "func_email_intelligence"
    )


def test_hardcoded_flash_model_strings_removed_from_email_sites():
    for rel in [
        "icharlotte_core/ui/email_update_tab.py",
        "icharlotte_core/sent_items_monitor.py",
    ]:
        src = (PROJECT_ROOT / rel).read_text(encoding="utf-8")
        # The LLM call sites no longer pin "gemini-3.5-flash" literally.
        assert '"gemini-3.5-flash"' not in src, rel
