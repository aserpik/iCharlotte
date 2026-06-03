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
        "Discovery",
        "Medical",
        "Motions & Drafting",
        "General",
    ]


def test_every_task_has_a_valid_category():
    for spec in list_tasks():
        assert spec.category in CATEGORY_ORDER, (
            f"{spec.task_id} has invalid category {spec.category!r}"
        )


def test_expected_category_assignments():
    assert TASK_REGISTRY["summarize_discovery"].category == "Discovery"
    assert TASK_REGISTRY["respond_to_discovery"].category == "Discovery"
    assert TASK_REGISTRY["subpoena_tracker"].category == "Discovery"
    assert TASK_REGISTRY["medical_records"].category == "Medical"
    assert TASK_REGISTRY["med_record_extractor"].category == "Medical"
    assert TASK_REGISTRY["oppose_motion"].category == "Motions & Drafting"
    assert TASK_REGISTRY["summarize_documents"].category == "General"
    assert TASK_REGISTRY["chat"].category == "General"


def test_empty_query_returns_all_tasks_grouped_in_category_order():
    grouped = filter_tasks(list_tasks(), "")
    # Only non-empty categories, in CATEGORY_ORDER.
    assert list(grouped.keys()) == CATEGORY_ORDER
    total = sum(len(v) for v in grouped.values())
    assert total == len(list_tasks())
    # Discovery has 5 tasks per the spec.
    assert len(grouped["Discovery"]) == 5


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
    # A query that hits Medical + Discovery should still order Discovery first.
    grouped = filter_tasks(list_tasks(), "records")
    keys = list(grouped.keys())
    # keys must be a subsequence of CATEGORY_ORDER
    assert keys == [c for c in CATEGORY_ORDER if c in keys]


def test_every_task_keywords_is_a_list():
    for spec in list_tasks():
        assert isinstance(spec.keywords, list)
