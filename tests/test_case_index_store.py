"""Shared per-case document-index store: round-trip + format compatibility."""
import json

from icharlotte_core import config
from icharlotte_core import case_index_store as store


def test_load_missing_returns_empty(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    assert store.load_index("1234") == {}


def test_upsert_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    docs = [{"id": "1", "title": "A", "date": "01/02/2020", "start": 1, "end": 3}]
    store.upsert_pdf("1234", r"C:\x\a.pdf", docs)
    assert store.load_index("1234") == {r"C:\x\a.pdf": docs}


def test_upsert_second_pdf_preserves_first(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    store.upsert_pdf("1234", "p1", [{"id": "1"}])
    store.upsert_pdf("1234", "p2", [{"id": "2"}])
    assert set(store.load_index("1234")) == {"p1", "p2"}


def test_upsert_same_pdf_overwrites(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    store.upsert_pdf("1234", "p1", [{"id": "1"}])
    store.upsert_pdf("1234", "p1", [{"id": "9"}])
    assert store.load_index("1234") == {"p1": [{"id": "9"}]}


def test_corrupt_returns_empty(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    with open(store.index_path("1234"), "w", encoding="utf-8") as f:
        f.write("{not valid json")
    assert store.load_index("1234") == {}


def test_on_disk_shape_is_dict_of_lists(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "GEMINI_DATA_DIR", str(tmp_path))
    store.upsert_pdf("1234", "p1", [{"id": "1", "title": "A"}])
    with open(store.index_path("1234"), "r", encoding="utf-8") as f:
        raw = json.load(f)
    assert isinstance(raw, dict)
    assert "p1" in raw
    assert isinstance(raw["p1"], list)
