"""Tests for icharlotte_core.deposition.session_manager."""

import json
import os
from pathlib import Path

import pytest

from icharlotte_core.deposition import session_manager


def test_compute_session_paths_are_deterministic_per_input(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    # Same input twice in same second yields different timestamps but same hash prefix
    p1 = session_manager.compute_session_paths(r"Z:\foo\Smith Depo.pdf")
    p2 = session_manager.compute_session_paths(r"Z:\foo\Smith Depo.pdf")
    assert p1.session_path.name.split("_", 1)[0] == p2.session_path.name.split("_", 1)[0]
    assert p1.cached_text_path.suffix == ".txt"
    assert p1.session_path.suffix == ".json"
    assert p1.session_path.parent == tmp_path


def test_write_and_read_session_round_trip(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    paths = session_manager.compute_session_paths(r"Z:\foo\X.pdf")
    data = {
        "version": 1,
        "phase": "awaiting_input",
        "input_path": r"Z:\foo\X.pdf",
        "cached_text_path": str(paths.cached_text_path),
        "deponent_name": "John Smith",
        "deposition_date": "January 15, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "3850.084",
        "topics": [{"id": 1, "title": "Topic A", "rank": 1, "discussion_density": "high"}],
        "user_config": None,
    }
    session_manager.write_session(paths.session_path, data)
    loaded = session_manager.read_session(paths.session_path)
    assert loaded == data


def test_write_session_is_atomic(tmp_path, monkeypatch):
    """If os.replace fails, the original file is not corrupted."""
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "s.json"
    session_path.write_text('{"phase": "awaiting_input", "version": 1}', encoding="utf-8")

    def boom(*a, **kw):
        raise OSError("simulated replace failure")

    monkeypatch.setattr(os, "replace", boom)
    with pytest.raises(OSError):
        session_manager.write_session(session_path, {"phase": "ready_for_summary", "version": 1})

    # Original file untouched
    assert json.loads(session_path.read_text(encoding="utf-8"))["phase"] == "awaiting_input"


def test_update_user_config_flips_phase_and_writes_config(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "s.json"
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "awaiting_input",
        "user_config": None,
        "topics": [],
    })

    config = {
        "selected_topics": ["A"],
        "added_topics": ["B"],
        "bullets_per_topic": 5,
        "deponent_label": "Plaintiff",
        "custom_rules": "Use past tense.",
        "cross_check_enabled": True,
    }
    session_manager.update_user_config(session_path, config)

    loaded = session_manager.read_session(session_path)
    assert loaded["phase"] == "ready_for_summary"
    assert loaded["user_config"] == config


def test_cleanup_session_removes_both_files(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "s.json"
    cached = tmp_path / "s.txt"
    session_path.write_text('{"cached_text_path": "' + str(cached).replace("\\", "\\\\") + '"}', encoding="utf-8")
    cached.write_text("transcript text", encoding="utf-8")

    session_manager.cleanup_session(session_path)
    assert not session_path.exists()
    assert not cached.exists()


def test_cleanup_session_tolerates_missing_files(tmp_path, monkeypatch):
    monkeypatch.setattr(session_manager, "SESSION_DIR", tmp_path)
    session_path = tmp_path / "missing.json"
    # Should not raise even though neither file exists.
    session_manager.cleanup_session(session_path)
