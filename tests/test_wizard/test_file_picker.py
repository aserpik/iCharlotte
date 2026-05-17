"""Tests for default-folder resolution."""
import os
import pytest

from icharlotte_core.ui.wizard.file_picker import resolve_default_folder


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
