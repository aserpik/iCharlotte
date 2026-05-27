"""Tests for the redesigned drafter signature (no authority_block; uses style exemplars)."""

from __future__ import annotations

from icharlotte_core.opposition.drafter import draft_memorandum
from icharlotte_core.opposition.models import MotionMetadata, SectionPlanItem


def _captured_prompts():
    captures = {"system": None, "user": None}

    def llm(system_prompt, user_prompt):
        captures["system"] = system_prompt
        captures["user"] = user_prompt
        return (
            '{"title": "Opposition to Motion to Compel", '
            '"body_text": "# I. INTRODUCTION\\n\\n*Smith v. Jones* (2010) 50 Cal.4th 100 controls."}'
        )

    return llm, captures


def test_drafter_runs_with_empty_style_exemplars_list():
    llm, _ = _captured_prompts()
    draft = draft_memorandum(
        metadata=MotionMetadata(motion_type="Motion to Compel", relief_requested="x", principal_arguments=["a"]),
        section_plan=[SectionPlanItem(id="i", path=["I"], text="Introduction")],
        motion_text="motion",
        context_text="context",
        style_exemplars=[],
        llm_callback=llm,
    )
    assert draft.body_text.strip()


def test_drafter_injects_style_exemplar_blocks_into_user_prompt():
    llm, captures = _captured_prompts()
    draft_memorandum(
        metadata=MotionMetadata(motion_type="Motion to Compel", relief_requested="x", principal_arguments=["a"]),
        section_plan=[SectionPlanItem(id="i", path=["I"], text="Introduction")],
        motion_text="motion",
        context_text="context",
        style_exemplars=[
            "First exemplar text here, paragraph one.\n\nSecond paragraph.",
            "Second exemplar with formal tone.",
        ],
        llm_callback=llm,
    )
    user = captures["user"] or ""
    assert "<style_exemplar_1>" in user
    assert "First exemplar text here" in user
    assert "<style_exemplar_2>" in user
    assert "Second exemplar with formal tone" in user


def test_drafter_when_no_exemplars_indicates_default_voice():
    llm, captures = _captured_prompts()
    draft_memorandum(
        metadata=MotionMetadata(motion_type="MTC", relief_requested="x", principal_arguments=["a"]),
        section_plan=[],
        motion_text="m",
        context_text="c",
        style_exemplars=[],
        llm_callback=llm,
    )
    user = captures["user"] or ""
    # Either explicit empty-state marker OR no style_exemplar blocks at all.
    assert "<style_exemplar_1>" not in user
