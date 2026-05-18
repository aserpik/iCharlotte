"""Tests for the Med-Cron session manager."""

import json
from pathlib import Path

import pytest

from icharlotte_core.med_chron import session_manager


def _make_input_file(tmp_path: Path, name: str = "Acme PT Records.docx") -> Path:
    p = tmp_path / name
    p.write_bytes(b"\x00" * 1024)
    return p


def test_compute_session_paths_layout(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    paths = session_manager.compute_session_paths(str(inp), str(out_dir))

    cache_root = out_dir / ".med_chron"
    assert paths.cache_dir.parent == cache_root
    assert paths.cache_dir.name and len(paths.cache_dir.name) == 12  # hash
    assert paths.session_path.name == "session.json"
    assert paths.narrative_text_path.name == "narrative.txt"
    assert paths.full_text_path.name == "full.txt"


def test_hash_changes_when_mtime_changes(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()

    p1 = session_manager.compute_session_paths(str(inp), str(out_dir))
    # Force a different mtime
    import os, time
    new_time = inp.stat().st_mtime_ns + 1_000_000_000
    os.utime(inp, ns=(new_time, new_time))
    p2 = session_manager.compute_session_paths(str(inp), str(out_dir))

    assert p1.cache_dir != p2.cache_dir


def test_write_then_read_session_round_trip(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    paths = session_manager.compute_session_paths(str(inp), str(out_dir))

    data = {"phase": "awaiting_input", "user_config": None, "k": 1}
    session_manager.write_session(paths.session_path, data)

    loaded = session_manager.read_session(paths.session_path)
    assert loaded == data


def test_update_user_config_flips_phase(tmp_path):
    inp = _make_input_file(tmp_path)
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    paths = session_manager.compute_session_paths(str(inp), str(out_dir))

    session_manager.write_session(
        paths.session_path,
        {"phase": "awaiting_input", "user_config": None},
    )
    session_manager.update_user_config(
        paths.session_path,
        {"selected_catalog_ids": ["rewrite_chronology"], "custom_analyses": []},
    )

    loaded = session_manager.read_session(paths.session_path)
    assert loaded["phase"] == "ready_to_run"
    assert loaded["user_config"]["selected_catalog_ids"] == ["rewrite_chronology"]


def test_write_session_creates_parent_dirs(tmp_path):
    deeply_nested = tmp_path / "a" / "b" / "c" / "session.json"
    session_manager.write_session(deeply_nested, {"phase": "x"})
    assert deeply_nested.exists()
