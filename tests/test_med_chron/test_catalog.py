"""Tests for the Med-Cron analysis catalog."""

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

from MED_CHRON_ANALYSES import catalog  # noqa: E402


def test_catalog_is_non_empty():
    assert len(catalog.CATALOG) > 0


def test_catalog_ids_are_unique():
    ids = [d.id for d in catalog.CATALOG]
    assert len(ids) == len(set(ids)), f"duplicate catalog ids: {ids}"


def test_only_rewrite_is_default_selected():
    defaults = [d for d in catalog.CATALOG if d.default_selected]
    assert len(defaults) == 1
    assert defaults[0].id == "rewrite_chronology"


def test_rewrite_does_not_use_tables():
    rewrite = catalog.CATALOG_BY_ID["rewrite_chronology"]
    assert rewrite.uses_tables is False


def test_every_catalog_entry_has_prompt_file_on_disk():
    prompts_dir = PROJECT_ROOT / "Scripts" / "MED_CHRON_ANALYSES" / "prompts"
    for d in catalog.CATALOG:
        path = prompts_dir / d.prompt_file
        assert path.exists(), f"missing prompt file: {path}"


def test_load_prompt_returns_file_contents():
    text = catalog.load_prompt("rewrite_chronology.txt")
    assert text
    assert isinstance(text, str)


def test_load_prompt_rejects_path_traversal():
    with pytest.raises(ValueError):
        catalog.load_prompt("../../../etc/passwd")


def test_custom_wrapper_contains_placeholder():
    wrapper = catalog.load_prompt("_custom_wrapper.txt")
    assert "{user_instruction}" in wrapper
