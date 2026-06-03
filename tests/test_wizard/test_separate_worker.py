"""Tests for SeparateAnalysisWorker output parsing (no real subprocess)."""
import json

import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.pages.separate_page import SeparateAnalysisWorker


def test_parse_json_map_from_stdout(tmp_path):
    docs = [{"id": "1", "title": "X", "date": "", "start": 1, "end": 2}]
    map_path = tmp_path / "map.json"
    map_path.write_text(json.dumps(docs), encoding="utf-8")
    stdout = f"some log line\nJSON_MAP: {map_path}\nmore\n"
    parsed = SeparateAnalysisWorker._parse_docs(stdout)
    assert parsed == docs


def test_parse_missing_marker_returns_none():
    assert SeparateAnalysisWorker._parse_docs("no marker here") is None
