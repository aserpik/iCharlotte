"""Session folder layout + JSON read/write tests."""
import json
import os
from pathlib import Path

import pytest

from Scripts.depo_prep_lib.session_io import (
    SessionPaths,
    build_session_folder_name,
    compute_session_paths,
    write_json,
    read_json,
    file_sha256,
)


def test_build_session_folder_name_includes_deponent_and_date():
    name = build_session_folder_name("Jane Doe", when_iso="2026-05-27T14:32:00")
    assert "Depo Prep" in name
    assert "Jane Doe" in name
    assert "2026-05-27" in name
    # No illegal Windows filename chars
    for bad in '\\/*?:"<>|':
        assert bad not in name


def test_build_session_folder_name_sanitizes_deponent():
    name = build_session_folder_name('Joe "Slick" O\'Malley/Jr', when_iso="2026-05-27T00:00:00")
    for bad in '\\/*?:"<>|':
        assert bad not in name
    assert "Slick" in name  # kept words, dropped quotes


def test_compute_session_paths_creates_expected_subpaths(tmp_path):
    case_root = tmp_path / "Smith v. Jones"
    case_root.mkdir()
    paths = compute_session_paths(
        case_root=str(case_root),
        deponent_name="Jane Doe",
        when_iso="2026-05-27T14:32:00",
    )
    assert isinstance(paths, SessionPaths)
    assert "NOTES" in str(paths.session_dir)
    assert "AI Output" in str(paths.session_dir)
    assert paths.session_json.name == "session.json"
    assert paths.topics_json.name == "topics.json"
    assert paths.digests_dir.name == "digests"
    assert paths.outline_docx.name == "outline.docx"
    assert paths.outline_md.name == "outline.md"


def test_write_and_read_json_roundtrip(tmp_path):
    p = tmp_path / "x.json"
    write_json(p, {"a": 1, "b": ["x", "y"]})
    assert read_json(p) == {"a": 1, "b": ["x", "y"]}


def test_write_json_creates_parent_dirs(tmp_path):
    p = tmp_path / "deep" / "nested" / "x.json"
    write_json(p, {"ok": True})
    assert p.exists()
    assert read_json(p)["ok"] is True


def test_file_sha256_is_stable(tmp_path):
    f = tmp_path / "x.txt"
    f.write_bytes(b"hello world")
    h1 = file_sha256(f)
    h2 = file_sha256(f)
    assert h1 == h2
    assert len(h1) == 64  # hex SHA-256
