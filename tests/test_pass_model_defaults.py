import ast
from pathlib import Path

from icharlotte_core import llm_config as llm_config_module
from icharlotte_core.llm_config import LLMCaller, LLMConfig, ModelSpec


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _reset_llm_config():
    LLMConfig._instance = None


def test_pass_model_defaults_override_agent_and_task_fallbacks(tmp_path, monkeypatch):
    monkeypatch.setattr(llm_config_module, "CONFIG_FILE", str(tmp_path / "llm_preferences.json"))
    _reset_llm_config()
    try:
        config = LLMConfig()
        config.update_agent_config(
            "agent_sum_depo",
            model_sequence=[ModelSpec(provider="Gemini", model="gemini-3.5-flash")],
            use_default=False,
        )
        config.update_pass_model_config(
            "agent_sum_depo",
            "summary",
            [ModelSpec(provider="OpenAI", model="gpt-4o-mini")],
        )

        summary_sequence = config.get_model_sequence_for_agent(
            "agent_sum_depo",
            fallback_task="summary",
            pass_name="summary",
        )
        cross_check_sequence = config.get_model_sequence_for_agent(
            "agent_sum_depo",
            fallback_task="cross_check",
            pass_name="cross_check",
        )

        assert [(m.provider, m.model) for m in summary_sequence] == [("OpenAI", "gpt-4o-mini")]
        assert [(m.provider, m.model) for m in cross_check_sequence] == [("Gemini", "gemini-3.5-flash")]

        _reset_llm_config()
        reloaded = LLMConfig()
        reloaded_sequence = reloaded.get_model_sequence_for_agent(
            "agent_sum_depo",
            fallback_task="summary",
            pass_name="summary",
        )
        assert [(m.provider, m.model) for m in reloaded_sequence] == [("OpenAI", "gpt-4o-mini")]
    finally:
        _reset_llm_config()


class _FakeConfig:
    def __init__(self):
        self.calls = []

    def get_model_sequence_for_agent(self, agent_id, fallback_task=None, pass_name=None):
        self.calls.append((agent_id, fallback_task, pass_name))
        return [ModelSpec(provider="Gemini", model="gemini-3.5-flash")]

    def get_agent_config(self, agent_id):
        return llm_config_module.AgentConfig(
            agent_id=agent_id,
            display_name=agent_id,
            model_sequence=[],
            use_default=True,
        )

    def is_provider_available(self, provider):
        return False


def test_llm_caller_forwards_pass_name_to_model_resolution(monkeypatch, tmp_path):
    monkeypatch.setattr(llm_config_module, "CONFIG_FILE", str(tmp_path / "llm_preferences.json"))
    _reset_llm_config()
    try:
        caller = LLMCaller()
        fake_config = _FakeConfig()
        caller.config = fake_config

        result = caller.call(
            prompt="prompt",
            text="text",
            task_type="summary",
            agent_id="agent_sum_depo",
            pass_name="summary",
        )

        assert result is None
        assert fake_config.calls == [("agent_sum_depo", "summary", "summary")]
    finally:
        _reset_llm_config()


def _llm_call_keyword_sets(source_path: Path):
    tree = ast.parse(source_path.read_text(encoding="utf-8"), filename=str(source_path))
    keyword_sets = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Attribute) or node.func.attr != "call":
            continue
        keywords = {}
        for kw in node.keywords:
            if kw.arg and isinstance(kw.value, ast.Constant):
                keywords[kw.arg] = kw.value.value
        keyword_sets.append(keywords)
    return keyword_sets


def _has_llm_call(keyword_sets, **expected):
    for keywords in keyword_sets:
        if all(keywords.get(name) == value for name, value in expected.items()):
            return True
    return False


def test_deposition_runtime_calls_send_workbench_pass_names():
    keyword_sets = _llm_call_keyword_sets(PROJECT_ROOT / "Scripts" / "summarize_deposition.py")
    selector_keyword_sets = _llm_call_keyword_sets(
        PROJECT_ROOT / "icharlotte_core" / "deposition" / "testimony_selector.py"
    )

    assert _has_llm_call(
        keyword_sets,
        agent_id="agent_sum_depo",
        task_type="summary",
        pass_name="topic_discovery",
    )
    assert _has_llm_call(
        keyword_sets,
        agent_id="agent_sum_depo",
        task_type="summary",
        pass_name="summary",
    )
    assert _has_llm_call(
        keyword_sets,
        pass_agent_id="agent_sum_depo",
        task_type="cross_check",
        pass_name="cross_check",
    )
    assert _has_llm_call(
        selector_keyword_sets,
        agent_id="agent_depo_extract",
        task_type="extraction",
        pass_name="extraction",
    )


def test_discovery_runtime_calls_send_workbench_pass_names():
    keyword_sets = _llm_call_keyword_sets(PROJECT_ROOT / "Scripts" / "summarize_discovery.py")

    assert _has_llm_call(
        keyword_sets,
        agent_id="agent_sum_disc",
        task_type="extraction",
        pass_name="extraction",
    )
    assert _has_llm_call(
        keyword_sets,
        agent_id="agent_sum_disc",
        task_type="summary",
        pass_name="summary",
    )
    assert _has_llm_call(
        keyword_sets,
        pass_agent_id="agent_sum_disc",
        task_type="cross_check",
        pass_name="cross_check",
    )
    assert _has_llm_call(
        keyword_sets,
        pass_agent_id="agent_sum_disc",
        task_type="summary",
        pass_name="consolidate",
        model_override="gemini-3.1-flash-lite",
    )


def test_med_chron_runtime_calls_send_workbench_agent_and_pass():
    keyword_sets = _llm_call_keyword_sets(PROJECT_ROOT / "Scripts" / "med_chron.py")

    assert _has_llm_call(
        keyword_sets,
        agent_id="agent_med_chron",
        pass_name="main",
        task_type="summary",
    )


def test_liability_and_exposure_route_through_config():
    for script, agent_id in [("liability.py", "agent_liability"), ("exposure.py", "agent_exposure")]:
        path = PROJECT_ROOT / "Scripts" / script
        keyword_sets = _llm_call_keyword_sets(path)
        assert _has_llm_call(keyword_sets, agent_id=agent_id, pass_name="main"), script
        # call_gemini no longer talks to the Gemini SDK directly.
        assert "client.models.generate_content" not in path.read_text(encoding="utf-8"), script


def test_summarize_runtime_calls_send_workbench_pass_names():
    keyword_sets = _llm_call_keyword_sets(PROJECT_ROOT / "Scripts" / "summarize.py")

    assert _has_llm_call(
        keyword_sets,
        agent_id="agent_summarize",
        task_type="summary",
        pass_name="summary",
    )
    assert _has_llm_call(
        keyword_sets,
        pass_agent_id="agent_summarize",
        task_type="cross_check",
        pass_name="cross_check",
    )
