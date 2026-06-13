"""Tests for Wizard launcher task categories + search filtering (spec #5).

Pure-logic tests — no Qt required. Cover the registry taxonomy fields and the
`filter_tasks` grouping/search helper.
"""
from icharlotte_core.ui.wizard.registry import (
    CATEGORY_ORDER,
    TASK_REGISTRY,
    filter_tasks,
    list_tasks,
)


def test_category_order_constant():
    assert CATEGORY_ORDER == [
        "General",
        "Summarize",
        "Discovery",
        "Medical",
        "Motions",
    ]


def test_every_task_has_a_valid_category():
    for spec in list_tasks():
        assert spec.category in CATEGORY_ORDER, (
            f"{spec.task_id} has invalid category {spec.category!r}"
        )


def test_expected_category_assignments():
    # General
    assert TASK_REGISTRY["chat"].category == "General"
    assert TASK_REGISTRY["separate"].category == "General"
    assert TASK_REGISTRY["case_intake_docket"].category == "General"
    # Summarize
    assert TASK_REGISTRY["summarize_documents"].category == "Summarize"
    assert TASK_REGISTRY["summarize_discovery"].category == "Summarize"
    assert TASK_REGISTRY["summarize_depositions"].category == "Summarize"
    # Discovery
    assert TASK_REGISTRY["respond_to_discovery"].category == "Discovery"
    assert TASK_REGISTRY["depo_prep"].category == "Discovery"
    # Medical
    assert TASK_REGISTRY["medical_records"].category == "Medical"
    assert TASK_REGISTRY["med_chron_analysis"].category == "Medical"
    assert TASK_REGISTRY["med_record_extractor"].category == "Medical"
    assert TASK_REGISTRY["subpoena_tracker"].category == "Medical"
    # Motions
    assert TASK_REGISTRY["motion_drafting"].category == "Motions"
    assert TASK_REGISTRY["oppose_motion"].category == "Motions"
    assert TASK_REGISTRY["generate_motion"].category == "Motions"
    assert TASK_REGISTRY["mediation_brief"].category == "Motions"


def test_empty_query_returns_all_tasks_grouped_in_category_order():
    grouped = filter_tasks(list_tasks(), "")
    # Only non-empty categories, in CATEGORY_ORDER.
    assert list(grouped.keys()) == CATEGORY_ORDER
    total = sum(len(v) for v in grouped.values())
    assert total == len(list_tasks())
    # Per-category counts per the spec.
    assert len(grouped["General"]) == 3
    assert len(grouped["Summarize"]) == 3
    assert len(grouped["Discovery"]) == 2
    assert len(grouped["Medical"]) == 4
    assert len(grouped["Motions"]) == 2


def test_whitespace_query_is_treated_as_empty():
    assert filter_tasks(list_tasks(), "   ") == filter_tasks(list_tasks(), "")


def test_keyword_alias_match_finds_respond_to_discovery():
    grouped = filter_tasks(list_tasks(), "rfp")
    ids = {s.task_id for specs in grouped.values() for s in specs}
    assert "respond_to_discovery" in ids


def test_keyword_alias_ime_finds_medical_records():
    grouped = filter_tasks(list_tasks(), "IME")
    ids = {s.task_id for specs in grouped.values() for s in specs}
    assert "medical_records" in ids


def test_title_substring_match_is_case_insensitive():
    grouped = filter_tasks(list_tasks(), "depo")
    ids = {s.task_id for specs in grouped.values() for s in specs}
    assert {"depo_prep", "summarize_depositions"} <= ids


def test_no_match_returns_empty_mapping():
    grouped = filter_tasks(list_tasks(), "zzzznotarealtask")
    assert grouped == {}


def test_filtered_result_preserves_category_order():
    # A query that hits Summarize + Discovery should still order Summarize first.
    grouped = filter_tasks(list_tasks(), "depo")
    keys = list(grouped.keys())
    assert "Summarize" in keys and "Discovery" in keys
    # keys must be a subsequence of CATEGORY_ORDER
    assert keys == [c for c in CATEGORY_ORDER if c in keys]


def test_every_task_keywords_is_a_list():
    for spec in list_tasks():
        assert isinstance(spec.keywords, list)
