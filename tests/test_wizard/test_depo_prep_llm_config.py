"""Verify the DepoPrep agent is registered in config/llm_preferences.json.

The Depo Prep pipeline calls LLMCaller.call(..., agent_id="DepoPrep", ...) from
several places (topic_clustering, topic_questions, dedup, polish). LLMConfig
falls back to a default AgentConfig when the agent_id is unknown, so missing
this entry would silently use whatever the default provider is configured to
instead of the deposition-tuned model sequence we want.

This test guards against the entry being removed or renamed.
"""

import json
from pathlib import Path


CONFIG_PATH = Path(__file__).resolve().parents[2] / "config" / "llm_preferences.json"


def _load_agents() -> dict:
    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    # Accept either nested {agents: {...}} or flat {DepoPrep: {...}} shape.
    return cfg.get("agents", cfg)


def test_depo_prep_agent_in_llm_preferences():
    agents = _load_agents()
    assert "DepoPrep" in agents, (
        f"DepoPrep not in agents (keys: {list(agents)[:10]})"
    )


def test_depo_prep_has_model_sequence():
    """DepoPrep should have a non-empty model_sequence (mirroring agent_sum_depo)."""
    agents = _load_agents()
    depo = agents["DepoPrep"]
    seq = depo.get("model_sequence")
    assert seq, f"DepoPrep.model_sequence is missing or empty: {depo}"
    assert isinstance(seq, list)
    first = seq[0]
    assert "model" in first and first["model"], "First model in sequence has no 'model' key"
    assert "provider" in first and first["provider"], "First model in sequence has no 'provider' key"


def test_depo_prep_mirrors_agent_sum_depo_models():
    """The model names should match agent_sum_depo so we don't introduce unknown IDs."""
    agents = _load_agents()
    depo = agents["DepoPrep"]
    sum_depo = agents["agent_sum_depo"]
    depo_models = [m["model"] for m in depo["model_sequence"]]
    sum_depo_models = [m["model"] for m in sum_depo["model_sequence"]]
    assert depo_models == sum_depo_models, (
        f"DepoPrep models {depo_models} should match agent_sum_depo models "
        f"{sum_depo_models}"
    )
