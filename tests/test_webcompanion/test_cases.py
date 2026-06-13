"""Tests for webcompanion.cases — path safety and browsing."""
import pytest

from webcompanion import cases


def test_safe_resolve_inside(tmp_path):
    (tmp_path / "DISCOVERY").mkdir()
    p = cases.safe_resolve(str(tmp_path), "DISCOVERY")
    assert p == (tmp_path / "DISCOVERY").resolve()


def test_safe_resolve_root_when_empty(tmp_path):
    assert cases.safe_resolve(str(tmp_path), "") == tmp_path.resolve()


def test_safe_resolve_rejects_traversal(tmp_path):
    with pytest.raises(ValueError):
        cases.safe_resolve(str(tmp_path), "..\\..\\Windows")
    with pytest.raises(ValueError):
        cases.safe_resolve(str(tmp_path), "../etc")


def test_browse_filters_extensions(tmp_path):
    (tmp_path / "sub").mkdir()
    (tmp_path / "a.pdf").write_bytes(b"x")
    (tmp_path / "b.docx").write_bytes(b"x")
    (tmp_path / "c.exe").write_bytes(b"x")
    dirs, files = cases.browse(str(tmp_path), "", (".pdf",))
    assert dirs == ["sub"] and files == ["a.pdf"]
    _, files2 = cases.browse(str(tmp_path), "", (".pdf", ".docx"))
    assert files2 == ["a.pdf", "b.docx"]


def test_browse_missing_dir_returns_empty(tmp_path):
    dirs, files = cases.browse(str(tmp_path), "NOPE", (".pdf",))
    assert dirs == [] and files == []


def test_resolve_start_folder_first_existing(tmp_path):
    (tmp_path / "DISCOVERY").mkdir()
    rel = cases.resolve_start_folder(
        str(tmp_path), ("DISCOVERY/RESPONSES", "DISCOVERY"))
    assert rel == "DISCOVERY"


def test_resolve_start_folder_falls_back_to_root(tmp_path):
    assert cases.resolve_start_folder(str(tmp_path), ("NOPE",)) == ""
