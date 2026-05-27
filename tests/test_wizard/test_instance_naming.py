"""Tests for tab-title disambiguation when the same task is opened twice."""
from icharlotte_core.ui.wizard.instance_naming import next_instance_suffix


def test_no_existing_returns_empty_suffix():
    assert next_instance_suffix("Summarize Documents", existing_titles=[]) == ""


def test_existing_base_only_returns_2():
    assert next_instance_suffix("Summarize Documents", existing_titles=["Summarize Documents"]) == "(2)"


def test_fills_gap_with_lowest_unused():
    existing = ["Summarize Documents", "Summarize Documents (3)"]
    assert next_instance_suffix("Summarize Documents", existing_titles=existing) == "(2)"


def test_returns_highest_plus_one_when_no_gap():
    existing = ["Summarize Documents", "Summarize Documents (2)", "Summarize Documents (3)"]
    assert next_instance_suffix("Summarize Documents", existing_titles=existing) == "(4)"


def test_ignores_unrelated_titles():
    existing = ["Medical Records", "Summarize Discovery"]
    assert next_instance_suffix("Summarize Documents", existing_titles=existing) == ""
