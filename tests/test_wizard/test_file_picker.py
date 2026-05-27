"""Tests for default-folder resolution."""
import os
import pytest

from icharlotte_core.ui.wizard.file_picker import (
    find_medical_summary_folder,
    resolve_default_folder,
)


def test_first_existing_pref_wins(tmp_path):
    (tmp_path / "DISCOVERY" / "RESPONSES").mkdir(parents=True)
    (tmp_path / "DISCOVERY").mkdir(exist_ok=True)
    result = resolve_default_folder(str(tmp_path), ["DISCOVERY/RESPONSES", "DISCOVERY"])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path / "DISCOVERY" / "RESPONSES"))


def test_falls_back_to_second_pref(tmp_path):
    (tmp_path / "DISCOVERY").mkdir()
    result = resolve_default_folder(str(tmp_path), ["DISCOVERY/RESPONSES", "DISCOVERY"])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path / "DISCOVERY"))


def test_case_insensitive_match(tmp_path):
    (tmp_path / "discovery" / "responses").mkdir(parents=True)
    result = resolve_default_folder(str(tmp_path), ["DISCOVERY/RESPONSES"])
    assert os.path.normpath(result).lower() == os.path.normpath(str(tmp_path / "discovery" / "responses")).lower()


def test_no_match_returns_case_root(tmp_path):
    result = resolve_default_folder(str(tmp_path), ["DOESNT/EXIST"])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path))


def test_empty_prefs_returns_case_root(tmp_path):
    result = resolve_default_folder(str(tmp_path), [])
    assert os.path.normpath(result) == os.path.normpath(str(tmp_path))


def test_missing_case_root_returns_input_unchanged(tmp_path):
    fake_root = str(tmp_path / "nonexistent")
    # Should not raise — returns the input string even though it doesn't exist.
    result = resolve_default_folder(fake_root, ["RECORDS"])
    assert result == fake_root


# --- find_medical_summary_folder -------------------------------------------

def test_med_summary_finds_canonical_uppercase_dash_variant(tmp_path):
    target = tmp_path / "RECORDS" / "MEDICAL SUMMARY - DO NOT PRODUCE"
    target.mkdir(parents=True)
    result = find_medical_summary_folder(str(tmp_path))
    assert os.path.normpath(result) == os.path.normpath(str(target))


def test_med_summary_finds_en_dash_variant(tmp_path):
    target = tmp_path / "RECORDS" / "Medical Summary – DO NOT PRODUCE"
    target.mkdir(parents=True)
    result = find_medical_summary_folder(str(tmp_path))
    assert os.path.normpath(result) == os.path.normpath(str(target))


def test_med_summary_finds_abbreviated_variant(tmp_path):
    target = tmp_path / "RECORDS" / "Med. Summary - DO NOT PRODUCE"
    target.mkdir(parents=True)
    result = find_medical_summary_folder(str(tmp_path))
    assert os.path.normpath(result) == os.path.normpath(str(target))


def test_med_summary_finds_plain_variant_when_no_do_not_produce(tmp_path):
    target = tmp_path / "RECORDS" / "Medical Summary"
    target.mkdir(parents=True)
    result = find_medical_summary_folder(str(tmp_path))
    assert os.path.normpath(result) == os.path.normpath(str(target))


def test_med_summary_prefers_do_not_produce_over_plain(tmp_path):
    (tmp_path / "RECORDS" / "Medical Summary").mkdir(parents=True)
    preferred = tmp_path / "RECORDS" / "Medical Summary - DO NOT PRODUCE"
    preferred.mkdir(parents=True)
    result = find_medical_summary_folder(str(tmp_path))
    assert os.path.normpath(result) == os.path.normpath(str(preferred))


def test_med_summary_returns_none_when_records_missing(tmp_path):
    assert find_medical_summary_folder(str(tmp_path)) is None


def test_med_summary_returns_none_when_no_matching_subfolder(tmp_path):
    (tmp_path / "RECORDS" / "Discovery Stuff").mkdir(parents=True)
    (tmp_path / "RECORDS" / "Subpoenas").mkdir(parents=True)
    assert find_medical_summary_folder(str(tmp_path)) is None


def test_med_summary_ignores_files_named_like_folder(tmp_path):
    (tmp_path / "RECORDS").mkdir()
    (tmp_path / "RECORDS" / "Medical Summary.docx").write_text("x", encoding="utf-8")
    assert find_medical_summary_folder(str(tmp_path)) is None


def test_med_summary_case_insensitive_records_folder(tmp_path):
    target = tmp_path / "records" / "medical summary - do not produce"
    target.mkdir(parents=True)
    result = find_medical_summary_folder(str(tmp_path))
    assert os.path.normpath(result).lower() == os.path.normpath(str(target)).lower()
