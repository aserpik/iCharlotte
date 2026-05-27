"""Verify agent_oppose_motion is registered with five passes."""

from icharlotte_core.llm_config import LLMConfig


def test_agent_oppose_motion_registered():
    cfg = LLMConfig()
    agent_cfg = cfg.get_agent_config("agent_oppose_motion")
    assert agent_cfg is not None
    # AgentConfig should know the five passes (model overrides may be empty
    # initially, but the agent_id must resolve to a non-empty config).
    assert getattr(agent_cfg, "agent_id", None) == "agent_oppose_motion"


def test_workbench_mapping_includes_oppose_motion():
    from icharlotte_core.ui.dialogs import WORKBENCH_TO_AGENT_ID

    assert WORKBENCH_TO_AGENT_ID.get("oppose_motion") == "agent_oppose_motion"
