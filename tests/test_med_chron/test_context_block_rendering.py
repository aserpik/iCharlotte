"""Tests for context-block rendering in the Med-Chron custom-analysis pipeline."""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import json
from unittest.mock import patch

import med_chron  # noqa: E402
from icharlotte_core.med_chron import session_manager  # noqa: E402


def test_custom_wrapper_template_has_context_block_placeholder():
    """The wrapper template must define a {context_block} placeholder so
    Phase 2 can inject (or omit) user-supplied context documents."""
    from MED_CHRON_ANALYSES.catalog import load_prompt

    wrapper = load_prompt("_custom_wrapper.txt")
    assert "{context_block}" in wrapper
    assert "{user_instruction}" in wrapper  # existing placeholder must remain


def _prep_session_with_custom(tmp_path: Path, custom: list[dict]) -> Path:
    """Hand-build a ready_to_run session for context-block tests."""
    cache = tmp_path / ".med_chron" / "abc123"
    cache.mkdir(parents=True)
    narrative_path = cache / "narrative.txt"
    narrative_path.write_text("narr", encoding="utf-8")
    full_path = cache / "full.txt"
    full_path.write_text("full text", encoding="utf-8")
    session_path = cache / "session.json"

    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "ready_to_run",
        "input_path": str(tmp_path / "rec.docx"),
        "narrative_text_path": str(narrative_path),
        "full_text_path": str(full_path),
        "narrative_missing": False,
        "provider_name": "Acme PT",
        "file_number": "1234.567",
        "catalog": [],
        "user_config": {
            "selected_catalog_ids": [],
            "custom_analyses": custom,
        },
    })
    return session_path


def test_phase2_includes_context_block_in_prompt(tmp_path):
    """Custom analysis with one .txt context file should produce a prompt
    that contains BEGIN/END markers AND the file's text."""
    ctx = tmp_path / "status_report.txt"
    ctx.write_text("Defense theory: Plaintiff exaggerates pain.", encoding="utf-8")

    session_path = _prep_session_with_custom(tmp_path, [{
        "label": "Defense targets",
        "instruction": "Identify providers worth deposing.",
        "context_files": [str(ctx)],
    }])

    captured_prompts = []
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda prompt, text, **kw: captured_prompts.append(prompt) or "# X"):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    assert len(captured_prompts) == 1
    prompt = captured_prompts[0]
    assert "--- BEGIN CONTEXT DOCUMENT: status_report.txt ---" in prompt
    assert "Defense theory: Plaintiff exaggerates pain." in prompt
    assert "--- END CONTEXT DOCUMENT ---" in prompt
    assert "ADDITIONAL CONTEXT DOCUMENTS PROVIDED BY THE USER" in prompt
    assert "{context_block}" not in prompt  # placeholder must be substituted


def test_phase2_omits_block_header_when_no_context_files(tmp_path):
    """When context_files is empty/absent, the rendered block must be empty
    string — no stray header text."""
    session_path = _prep_session_with_custom(tmp_path, [{
        "label": "no ctx",
        "instruction": "Find left-knee mentions.",
    }])  # no context_files key at all

    captured_prompts = []
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda prompt, text, **kw: captured_prompts.append(prompt) or "# X"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    prompt = captured_prompts[0]
    assert "ADDITIONAL CONTEXT DOCUMENTS" not in prompt
    assert "BEGIN CONTEXT DOCUMENT" not in prompt
    assert "{context_block}" not in prompt


def test_phase2_skips_missing_context_file_but_still_runs(tmp_path):
    """A context_files path that doesn't exist must be skipped with a warning
    log; the analysis continues with the remaining files."""
    ctx_ok = tmp_path / "good.txt"
    ctx_ok.write_text("USABLE_CONTEXT", encoding="utf-8")
    ctx_missing = tmp_path / "does_not_exist.txt"

    session_path = _prep_session_with_custom(tmp_path, [{
        "label": "mixed",
        "instruction": "Look at this.",
        "context_files": [str(ctx_missing), str(ctx_ok)],
    }])

    captured_prompts = []
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda prompt, text, **kw: captured_prompts.append(prompt) or "# X"):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    prompt = captured_prompts[0]
    assert "USABLE_CONTEXT" in prompt
    assert "does_not_exist.txt" not in prompt  # missing file is silently dropped


def test_phase2_all_context_files_failing_still_runs_with_empty_block(tmp_path):
    """If every attached file fails to extract, the analysis still runs;
    block header is omitted (treated as no-context)."""
    session_path = _prep_session_with_custom(tmp_path, [{
        "label": "all-fail",
        "instruction": "Look at this.",
        "context_files": [str(tmp_path / "ghost1.txt"), str(tmp_path / "ghost2.txt")],
    }])

    captured_prompts = []
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda prompt, text, **kw: captured_prompts.append(prompt) or "# X"):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
    prompt = captured_prompts[0]
    assert "ADDITIONAL CONTEXT DOCUMENTS" not in prompt
    assert "{context_block}" not in prompt


def test_phase2_truncates_oversized_context_file(tmp_path):
    """Per-file truncation cap of MAX_CONTEXT_CHARS prevents runaway prompts."""
    ctx = tmp_path / "big.txt"
    huge = "A" * (med_chron.MAX_CONTEXT_CHARS + 5000)
    ctx.write_text(huge, encoding="utf-8")

    session_path = _prep_session_with_custom(tmp_path, [{
        "label": "big",
        "instruction": "x",
        "context_files": [str(ctx)],
    }])

    captured_prompts = []
    with patch.object(med_chron.LLMCaller, "call",
                      side_effect=lambda prompt, text, **kw: captured_prompts.append(prompt) or "# X"):
        med_chron.process_run(str(session_path), str(tmp_path / "out"))

    prompt = captured_prompts[0]
    assert "context truncated at" in prompt
    # The whole 125k-char blob can't survive; the inner A-block must be at most cap.
    # Sanity-check via the marker only — exact char count is checked indirectly.
    assert prompt.count("A") <= med_chron.MAX_CONTEXT_CHARS + 200  # cap + scaffolding


def test_phase2_backward_compat_no_context_files_key(tmp_path):
    """An older session JSON whose custom_analyses entry has no context_files
    key must still work (treated as empty list)."""
    session_path = _prep_session_with_custom(tmp_path, [{
        "label": "old shape",
        "instruction": "Find left-knee mentions.",
    }])

    with patch.object(med_chron.LLMCaller, "call", return_value="# X"):
        rc = med_chron.process_run(str(session_path), str(tmp_path / "out"))

    assert rc == 0
