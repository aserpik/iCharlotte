import pytest

from Scripts.depo_prep_lib.prompts import (
    build_per_source_digest_prompt,
    build_topic_clustering_prompt,
    build_per_topic_questions_prompt,
    build_dedup_prompt,
    build_polish_prompt,
    STYLE_DIRECTIVES,
)


def test_style_directives_has_all_four():
    assert set(STYLE_DIRECTIVES.keys()) == {"discovery", "lockdown", "expert", "friendly"}
    for v in STYLE_DIRECTIVES.values():
        assert isinstance(v, str) and len(v) > 50  # non-trivial directive


def test_per_source_digest_prompt_mentions_deponent_and_kind_hint():
    prompt, text_payload = build_per_source_digest_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", source_text="...transcript...",
        source_filename="jane_doe_depo.pdf",
    )
    assert "Jane Doe" in prompt
    assert "Plaintiff" in prompt
    assert "deponent_statements" in prompt  # field name guidance
    assert "factual_anchors" in prompt
    assert "inconsistencies" in prompt
    assert "JSON" in prompt
    assert text_payload == "...transcript..."


def test_topic_clustering_prompt_includes_style_and_count_range():
    prompt, text_payload = build_topic_clustering_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff",
        style="lockdown",
        free_text_notes="Focus on causation and prior injuries.",
        digests_summary_text="...digests...",
    )
    assert "8" in prompt and "15" in prompt  # 8-15 topic range
    assert "lockdown" in prompt.lower() or "lock-down" in prompt.lower() or STYLE_DIRECTIVES["lockdown"][:30] in prompt
    assert "causation" in prompt  # free text injected
    assert text_payload == "...digests..."


def test_per_topic_questions_prompt_conditionally_includes_field_instructions():
    # All flags off → only basic question text
    prompt, _ = build_per_topic_questions_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", style="discovery",
        topic_title="Pre-existing conditions",
        strategic_note="Establish chronic LBP",
        digest_excerpts_text="...digest excerpts...",
        free_text_notes="",
        include_strategic_note=False,
        include_source_facts=False,
        include_impeachment_hook=False,
        include_objection_alts=False,
    )
    assert "purpose" not in prompt.lower()
    assert "source_facts" not in prompt.lower()
    assert "impeachment" not in prompt.lower()
    assert "objection" not in prompt.lower()

    # All flags on → all field instructions present
    prompt2, _ = build_per_topic_questions_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", style="discovery",
        topic_title="Pre-existing conditions",
        strategic_note="Establish chronic LBP",
        digest_excerpts_text="...digest excerpts...",
        free_text_notes="",
        include_strategic_note=True,
        include_source_facts=True,
        include_impeachment_hook=True,
        include_objection_alts=True,
    )
    assert "purpose" in prompt2.lower()
    assert "source_facts" in prompt2
    assert "impeachment_hook" in prompt2 or "impeachment" in prompt2.lower()
    assert "objection_alts" in prompt2 or "objection" in prompt2.lower()


def test_per_topic_questions_truncates_free_text_above_2000_chars():
    long_notes = "x" * 5000
    prompt, _ = build_per_topic_questions_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", style="discovery",
        topic_title="t", strategic_note="s", digest_excerpts_text="d",
        free_text_notes=long_notes,
        include_strategic_note=True, include_source_facts=False,
        include_impeachment_hook=False, include_objection_alts=False,
    )
    # The truncation cap is 2000 chars; we should see fewer than 2050 'x's plus a marker.
    assert prompt.count("x") < 2100
    assert "truncated" in prompt.lower()


def test_dedup_prompt_takes_topic_outputs_summary():
    prompt, text = build_dedup_prompt(topic_outputs_summary="Topic 1: 4 Qs\nTopic 2: 5 Qs",
                                       digest_summary="med_records: causation, gaps")
    assert "duplicates" in prompt.lower()
    assert "coverage" in prompt.lower()
    assert "Topic 1" in text or "Topic 1" in prompt


def test_polish_prompt_forbids_substantive_changes():
    prompt, text = build_polish_prompt(outline_text="full outline here")
    # The polish prompt must explicitly forbid adding/dropping/changing questions.
    assert "no new" in prompt.lower() or "do not add" in prompt.lower()
    assert "drop" in prompt.lower() or "remove" in prompt.lower()
    assert "phrasing" in prompt.lower() or "transitions" in prompt.lower()
    assert text == "full outline here"


def test_per_topic_questions_prompt_no_trailing_comma_in_schema_example():
    """The assembled JSON schema example must not have a trailing comma before the
    closing object brace — the LLM mimics what it sees."""
    import re
    prompt, _ = build_per_topic_questions_prompt(
        deponent_name="J", deponent_role="P", style="discovery",
        topic_title="t", strategic_note="s", digest_excerpts_text="d",
        free_text_notes="",
        include_strategic_note=True, include_source_facts=True,
        include_impeachment_hook=True, include_objection_alts=True,
    )
    # The schema example ends with the closing brace of the question object.
    # Find the last "objection_alts" mention and the next "}" — there must NOT be
    # a comma immediately before that brace.
    idx = prompt.rfind("objection_alts")
    assert idx != -1
    tail = prompt[idx:]
    # The closing brace of the question object is on the line after the last field.
    # Match: any field value + optional whitespace/newline + closing brace.
    # We assert there is no `,` immediately before the first `}` we see after the
    # last field.
    m = re.search(r"objection_alts.*?(?P<close>[,\s]*})", tail, re.DOTALL)
    assert m is not None
    closing_segment = m.group("close")
    assert "," not in closing_segment, (
        f"Trailing comma before closing brace in schema example. "
        f"Segment: {closing_segment!r}")


def test_per_topic_questions_prompt_no_trailing_comma_when_only_one_optional_field():
    """Single-flag case: also no trailing comma."""
    prompt, _ = build_per_topic_questions_prompt(
        deponent_name="J", deponent_role="P", style="discovery",
        topic_title="t", strategic_note="s", digest_excerpts_text="d",
        free_text_notes="",
        include_strategic_note=True, include_source_facts=False,
        include_impeachment_hook=False, include_objection_alts=False,
    )
    # Find the purpose line; immediately after the closing quote/value there must
    # not be a comma followed by whitespace/newline and then `}`.
    idx = prompt.rfind('"purpose"')
    assert idx != -1
    tail = prompt[idx:]
    import re
    m = re.search(r'"purpose".*?(?P<close>[,\s]*})', tail, re.DOTALL)
    assert m is not None
    assert "," not in m.group("close")
