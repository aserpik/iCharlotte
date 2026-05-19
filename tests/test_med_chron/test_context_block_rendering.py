"""Tests for context-block rendering in the Med-Chron custom-analysis pipeline."""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))


def test_custom_wrapper_template_has_context_block_placeholder():
    """The wrapper template must define a {context_block} placeholder so
    Phase 2 can inject (or omit) user-supplied context documents."""
    from MED_CHRON_ANALYSES.catalog import load_prompt

    wrapper = load_prompt("_custom_wrapper.txt")
    assert "{context_block}" in wrapper
    assert "{user_instruction}" in wrapper  # existing placeholder must remain
