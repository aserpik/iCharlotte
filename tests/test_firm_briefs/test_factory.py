import os
from icharlotte_core.firm_briefs import factory


def test_paths_under_data_dir(monkeypatch, tmp_path):
    monkeypatch.setattr(factory, "DATA_DIR", str(tmp_path))
    db, vec = factory.index_paths()
    assert db.startswith(str(tmp_path))
    assert vec.startswith(str(tmp_path))


def test_available_false_when_absent(monkeypatch, tmp_path):
    monkeypatch.setattr(factory, "DATA_DIR", str(tmp_path))
    assert factory.index_available() is False


def test_make_index_none_when_unavailable(monkeypatch, tmp_path):
    monkeypatch.setattr(factory, "DATA_DIR", str(tmp_path))
    assert factory.make_index() is None
