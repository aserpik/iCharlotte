"""Tests for Phase 2 (run) of the Med-Cron agent."""

import json
import os
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import med_chron  # noqa: E402
from icharlotte_core.med_chron import session_manager  # noqa: E402


def _prep_session(tmp_path: Path, *, narrative: str = "narr",
                  full: str = "full text", selected: list[str] = None,
                  custom: list[dict] = None) -> Path:
    """Hand-build a ready_to_run session for direct Phase 2 tests."""
    cache = tmp_path / ".med_chron" / "abc123"
    cache.mkdir(parents=True)
    narrative_path = cache / "narrative.txt"
    narrative_path.write_text(narrative, encoding="utf-8")
    full_path = cache / "full.txt"
    full_path.write_text(full, encoding="utf-8")
    session_path = cache / "session.json"

    data = {
        "version": 1,
        "phase": "ready_to_run",
        "input_path": str(tmp_path / "rec.docx"),
        "narrative_text_path": str(narrative_path),
        "full_text_path": str(full_path),
        "narrative_missing": narrative == "",
        "provider_name": "Acme PT",
        "file_number": "1234.567",
        "catalog": [],
        "user_config": {
            "selected_catalog_ids": selected if selected is not None else ["rewrite_chronology"],
            "custom_analyses": custom or [],
        },
    }
    session_manager.write_session(session_path, data)
    return session_path


def test_run_one_uses_narrative_when_uses_tables_false(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="NARR_ONLY",
        full="NARR_ONLY plus TABLE_TOKEN",
        selected=["rewrite_chronology"],
    )

    captured = {}

    def fake_call(prompt, text, **kw):
        captured["text"] = text
        return "# Result\nbody"

    with patch.object(med_chron.LLMCaller, "call", side_effect=fake_call):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    assert captured["text"] == "NARR_ONLY"
    assert "TABLE_TOKEN" not in captured["text"]


def test_run_one_uses_full_text_when_uses_tables_true(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="NARR_ONLY",
        full="NARR_ONLY plus TABLE_TOKEN",
        selected=["inconsistencies"],
    )

    captured = {}
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda prompt, text, **kw: captured.setdefault("text", text) or "# X"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert "TABLE_TOKEN" in captured["text"]


def test_per_run_failure_does_not_abort_siblings(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="narr",
        full="full",
        selected=["rewrite_chronology", "inconsistencies"],
    )

    call_count = {"n": 0}

    def flaky(prompt, text, **kw):
        call_count["n"] += 1
        if call_count["n"] == 1:
            raise RuntimeError("simulated LLM error")
        return "# Survivor"

    with patch.object(med_chron.LLMCaller, "call", side_effect=flaky):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    out_files = list((tmp_path / "out").rglob("*.docx"))
    assert len(out_files) == 1


def test_all_runs_failed_exits_nonzero(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=["rewrite_chronology", "inconsistencies"],
    )
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=RuntimeError("boom")):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))
    assert rc == 1


def test_skips_rewrite_when_narrative_missing(tmp_path):
    session_path = _prep_session(
        tmp_path,
        narrative="",
        full="full body",
        selected=["rewrite_chronology", "inconsistencies"],
    )

    calls = []
    def stub_call(prompt, text, **kw):
        calls.append((prompt[:30], text[:30]))
        return "# OK"

    with patch.object(med_chron.LLMCaller, "call", side_effect=stub_call):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    assert len(calls) == 1


def test_custom_analysis_wraps_user_instruction_in_template(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=[],
        custom=[{"label": "Left-knee mentions",
                  "instruction": "Find every entry mentioning the left knee."}],
    )

    captured_prompts = []
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda prompt, text, **kw: captured_prompts.append(prompt) or "# X"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert len(captured_prompts) == 1
    assert "Find every entry mentioning the left knee." in captured_prompts[0]
    assert "{user_instruction}" not in captured_prompts[0]


def test_output_filenames_use_analysis_id(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=["rewrite_chronology", "inconsistencies"],
    )
    with patch.object(med_chron.LLMCaller, "call", return_value="# Result"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    out_files = sorted(p.name for p in (tmp_path / "out").rglob("*.docx"))
    assert any("med_chron_rewrite_chronology_rec.docx" == n for n in out_files)
    assert any("med_chron_inconsistencies_rec.docx" == n for n in out_files)


def test_custom_output_includes_index_and_slug(tmp_path):
    session_path = _prep_session(
        tmp_path,
        selected=[],
        custom=[
            {"label": "Left knee", "instruction": "..."},
            {"label": "Left knee", "instruction": "..."},
        ],
    )
    with patch.object(med_chron.LLMCaller, "call", return_value="# Result"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    out_files = sorted(p.name for p in (tmp_path / "out").rglob("*.docx"))
    assert any("custom_1_left_knee_rec.docx" in n for n in out_files)
    assert any("custom_2_left_knee_rec.docx" in n for n in out_files)


def test_bails_if_phase_not_ready_to_run(tmp_path):
    session_path = _prep_session(tmp_path)
    data = session_manager.read_session(session_path)
    data["phase"] = "awaiting_input"
    session_manager.write_session(session_path, data)

    rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))
    assert rc == 1


def test_slug_helper_lowercases_and_replaces_special_chars():
    assert med_chron._slug("Left-knee mentions") == "left-knee_mentions"
    assert med_chron._slug("  whitespace  ") == "whitespace"
    assert med_chron._slug("Has !! punct?") == "has_punct"


def test_max_workers_capped_at_4(tmp_path):
    """Even with many analyses queued, ThreadPoolExecutor uses at most 4 workers."""
    custom = [{"label": f"c{i}", "instruction": "do x"} for i in range(10)]
    session_path = _prep_session(tmp_path, selected=[], custom=custom)

    captured = {}
    real_tpe = med_chron.ThreadPoolExecutor
    class SpyTPE(real_tpe):
        def __init__(self, max_workers=None, *a, **kw):
            captured["max_workers"] = max_workers
            super().__init__(max_workers=max_workers, *a, **kw)

    with patch.object(med_chron, "ThreadPoolExecutor", SpyTPE):
        with patch.object(med_chron.LLMCaller, "call", return_value="# X"):
            med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert captured["max_workers"] == 4  # min(10, 4) → 4
