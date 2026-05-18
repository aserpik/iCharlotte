"""Tests for the global custom-analyses persistence store."""

import json

import pytest

from icharlotte_core.med_chron import custom_analyses_store


@pytest.fixture(autouse=True)
def _redirect_store(tmp_path, monkeypatch):
    monkeypatch.setattr(
        custom_analyses_store, "_STORE_PATH", tmp_path / "store.json"
    )


def test_load_returns_empty_when_file_missing():
    assert custom_analyses_store.load() == []


def test_save_then_load_round_trip():
    entries = [
        {"label": "A", "instruction": "first"},
        {"label": "B", "instruction": "second"},
    ]
    custom_analyses_store.save(entries)
    assert custom_analyses_store.load() == entries


def test_save_drops_empty_label_or_instruction():
    custom_analyses_store.save([
        {"label": "ok", "instruction": "ok"},
        {"label": "", "instruction": "no label"},
        {"label": "no instruction", "instruction": ""},
        {"label": "  ", "instruction": "whitespace only"},
    ])
    assert custom_analyses_store.load() == [
        {"label": "ok", "instruction": "ok"}
    ]


def test_load_drops_malformed_entries():
    custom_analyses_store.store_path().parent.mkdir(parents=True, exist_ok=True)
    custom_analyses_store.store_path().write_text(json.dumps([
        {"label": "good", "instruction": "good"},
        "not a dict",
        {"label": "missing instruction"},
        {"instruction": "missing label"},
        {"label": 123, "instruction": "non-string label"},
    ]), encoding="utf-8")
    assert custom_analyses_store.load() == [
        {"label": "good", "instruction": "good"}
    ]


def test_load_returns_empty_on_corrupt_json():
    custom_analyses_store.store_path().parent.mkdir(parents=True, exist_ok=True)
    custom_analyses_store.store_path().write_text("not json {", encoding="utf-8")
    assert custom_analyses_store.load() == []


def test_save_overwrites_existing_file():
    custom_analyses_store.save([{"label": "first", "instruction": "first"}])
    custom_analyses_store.save([{"label": "second", "instruction": "second"}])
    assert custom_analyses_store.load() == [
        {"label": "second", "instruction": "second"}
    ]


def test_save_creates_parent_directory():
    custom_analyses_store.save([{"label": "x", "instruction": "x"}])
    assert custom_analyses_store.store_path().exists()
