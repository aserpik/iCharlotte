"""Tests for the oppose_motion default-prompt constants module."""

from icharlotte_core.opposition import prompts


def test_module_exposes_all_five_pass_constants():
    expected = [
        "ANALYZE_MOTION_PROMPT",
        "GENERATE_OUTLINE_PROMPT",
        "DRAFT_MEMORANDUM_PROMPT",
        "VERIFY_CITATION_PROMPT",
        "FIND_REPLACEMENT_PROMPT",
    ]
    for name in expected:
        assert hasattr(prompts, name), f"Missing constant: {name}"
        value = getattr(prompts, name)
        assert isinstance(value, str)
        assert value.strip(), f"{name} is empty"


def test_draft_prompt_does_not_reference_authority_block():
    # The redesigned drafter no longer receives a pre-fetched authority block.
    assert "authority_block" not in prompts.DRAFT_MEMORANDUM_PROMPT.lower()


def test_verify_prompt_returns_json_verdict_keys():
    # The verifier prompt must instruct the LLM to return verdict / evidence / note.
    text = prompts.VERIFY_CITATION_PROMPT
    assert '"verdict"' in text or "verdict:" in text
    assert "SUPPORTED" in text
    assert "PARTIAL" in text
    assert "NOT_SUPPORTED" in text


def test_draft_prompt_supports_style_exemplar_blocks():
    # The drafter prompt must reference the style_exemplars placeholder.
    assert "{style_exemplars}" in prompts.DRAFT_MEMORANDUM_PROMPT
