"""Tests for the redesigned drafter signature (no authority_block; uses style exemplars)."""

from __future__ import annotations

from icharlotte_core.opposition.drafter import draft_memorandum
from icharlotte_core.opposition.models import MotionMetadata, SectionPlanItem
from icharlotte_core.opposition.models import RetrievedAuthority


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


def test_drafter_injects_labeled_authority_pool():
    llm, captures = _captured_prompts()
    draft_memorandum(
        metadata=MotionMetadata(motion_type="Motion to Compel", relief_requested="x", principal_arguments=["a"]),
        section_plan=[SectionPlanItem(id="i", path=["I"], text="Introduction")],
        motion_text="motion",
        context_text="context",
        style_exemplars=[],
        retrieved_authorities=[
            RetrievedAuthority(
                argument_text="Discovery cutoff bars the motion",
                cluster_id="111",
                case_name="A v. B",
                citation="2 Cal.5th 2",
                supports="discretion is broad",
                passage="The court held discretion is broad.",
            )
        ],
        llm_callback=llm,
    )
    user = captures["user"] or ""
    assert "Discovery cutoff bars the motion" in user
    assert "A v. B" in user
    assert "2 Cal.5th 2" in user
    assert "The court held discretion is broad." in user


def test_authority_pool_uses_year_parenthetical_and_is_parseable():
    """The pool must render "Name (year) cite" (no comma before the year) so the
    drafter copies a citation the parser can recognize — otherwise the output
    panel can't make cites selectable/verifiable."""
    from icharlotte_core.opposition.drafter import _format_authority_pool
    from icharlotte_core.opposition.citation_parser import extract_citations

    pool = _format_authority_pool([
        RetrievedAuthority(argument_text="Forensic imaging",
                           case_name="Ellis v. Toshiba America Information Systems, Inc.",
                           citation="218 Cal. App. 4th 853", year="2013",
                           supports="s", passage="p"),
    ])
    assert "Ellis v. Toshiba America Information Systems, Inc. (2013) 218 Cal. App. 4th 853" in pool
    # A body that copies that line verbatim (italicized) must parse to one cite.
    body = ("The point is settled. *Ellis v. Toshiba America Information Systems, Inc.* "
            "(2013) 218 Cal. App. 4th 853.")
    cites = extract_citations(body)
    assert len(cites) == 1 and cites[0].year == "2013"
    assert cites[0].reporter_citation == "218 Cal. App. 4th 853"


def test_authority_pool_omits_year_parenthetical_when_year_missing():
    from icharlotte_core.opposition.drafter import _format_authority_pool
    pool = _format_authority_pool([
        RetrievedAuthority(argument_text="x", case_name="A v. B", citation="2 Cal.5th 2"),
    ])
    assert "A v. B 2 Cal.5th 2" in pool and "()" not in pool


def test_drafter_pool_empty_message_when_no_authorities():
    llm, captures = _captured_prompts()
    draft_memorandum(
        metadata=MotionMetadata(motion_type="MTC", relief_requested="x", principal_arguments=["a"]),
        section_plan=[],
        motion_text="m",
        context_text="c",
        style_exemplars=[],
        retrieved_authorities=[],
        llm_callback=llm,
    )
    user = (captures["user"] or "").lower()
    assert "no" in user and "authority" in user
