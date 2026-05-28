"""Verify PromptManager seeds all five oppose_motion prompts on first run."""

import os
import tempfile

import pytest

from icharlotte_core.prompt_manager import PromptManager


@pytest.fixture
def fresh_manager():
    with tempfile.TemporaryDirectory() as tmp:
        # Initialize with prompts_dir pointed at an empty temp tree.
        mgr = PromptManager(prompts_dir=tmp)
        yield mgr


def test_seed_creates_all_oppose_motion_prompts(fresh_manager):
    fresh_manager.seed_pipeline_prompts()

    expected_passes = [
        "analyze_motion",
        "generate_outline",
        "draft_memorandum",
        "verify_citation",
        "find_replacement",
    ]
    for pass_name in expected_passes:
        text = fresh_manager.get_prompt("oppose_motion", pass_name)
        assert text is not None, f"missing prompt: oppose_motion:{pass_name}"
        assert text.strip(), f"empty prompt: oppose_motion:{pass_name}"


def test_seed_is_idempotent(fresh_manager):
    fresh_manager.seed_pipeline_prompts()
    first = fresh_manager.get_prompt("oppose_motion", "draft_memorandum")
    # Run a second time; should not duplicate or wipe the prompt.
    fresh_manager.seed_pipeline_prompts()
    second = fresh_manager.get_prompt("oppose_motion", "draft_memorandum")
    assert first == second


def test_seed_includes_research_and_rerank_prompts(tmp_path):
    from icharlotte_core.prompt_manager import PromptManager

    pm = PromptManager(prompts_dir=str(tmp_path))
    pm.seed_pipeline_prompts()

    assert pm.get_prompt("oppose_motion", "research_queries")
    assert pm.get_prompt("oppose_motion", "rerank_select")
