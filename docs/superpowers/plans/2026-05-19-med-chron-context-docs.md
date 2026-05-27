# Med Chron Custom Analyses — Per-Row Context Documents Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let users attach `.pdf` / `.docx` / `.txt` context documents to a single custom Med-Chron analysis (via drag-drop onto the instruction textbox or via an "+ Add context" button). Files attach per-row; instruction text persists globally, files do not.

**Architecture:** Add a `context_files: list[str]` field to each `user_config["custom_analyses"][i]` entry in the session JSON. Phase 2 reads each file, builds a labeled context block, and substitutes it into a new `{context_block}` placeholder in `_custom_wrapper.txt`. The global `custom_analyses_store` is unchanged — it persists only `{label, instruction}` so files do not leak across cases. UI gets a chip strip under each instruction textbox plus a drag-drop-aware `QPlainTextEdit` subclass.

**Tech Stack:** Python, PySide6, pypdf, python-docx, pytest, pytest-qt.

**Spec:** [docs/superpowers/specs/2026-05-19-med-chron-context-docs-design.md](../specs/2026-05-19-med-chron-context-docs-design.md)

---

## File Structure

**Modify:**

- `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt` — add `{context_block}` placeholder.
- `Scripts/med_chron.py` — new `_render_context_block(...)` helper, updated `_build_run_list(...)`.
- `icharlotte_core/ui/med_chron_config_form.py` — new `sniff_text_layer(...)` helper, new `ContextDropTextEdit` class, chip-strip additions on `CustomAnalysisRow`, dual-shape commit on `MedChronConfigForm`.

**Create:**

- `tests/test_med_chron/test_context_block_rendering.py` — backend tests for Phase 2 + helper.
- `tests/test_wizard/test_med_chron_context_docs.py` — UI tests for sniff helper, drop event, chip strip, dual-shape commit.

No new modules. All new helpers/classes live alongside the code that uses them, following the existing pattern.

---

## Task 1: Add `{context_block}` placeholder to the custom wrapper template

**Files:**
- Modify: `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt`
- Test: `tests/test_med_chron/test_context_block_rendering.py` (new file)

- [ ] **Step 1: Write the failing test**

Create `tests/test_med_chron/test_context_block_rendering.py` with this content:

```python
"""Tests for context-block rendering in the Med-Cron custom-analysis pipeline."""

import json
import sys
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "Scripts"))

import med_chron  # noqa: E402
from icharlotte_core.med_chron import session_manager  # noqa: E402


def test_custom_wrapper_template_has_context_block_placeholder():
    """The wrapper template must define a {context_block} placeholder so
    Phase 2 can inject (or omit) user-supplied context documents."""
    scripts_dir = PROJECT_ROOT / "Scripts"
    sys.path.insert(0, str(scripts_dir))
    from MED_CHRON_ANALYSES.catalog import load_prompt

    wrapper = load_prompt("_custom_wrapper.txt")
    assert "{context_block}" in wrapper
    assert "{user_instruction}" in wrapper  # existing placeholder must remain
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_med_chron/test_context_block_rendering.py::test_custom_wrapper_template_has_context_block_placeholder -v`

Expected: FAIL with `assert "{context_block}" in wrapper` because the template doesn't yet contain that string.

- [ ] **Step 3: Update the wrapper template**

Replace the entire contents of `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt` with:

```
You will be given the BRIEF SYNOPSIS sections of a medical chronology PLUS
the underlying tables of medical entries.

{context_block}

The user has asked you to perform the following analysis on this
chronology:

{user_instruction}

Ground every finding in the document. Cite specific dates, providers, and
entries where applicable. When context documents are provided above, you
may reference them but the medical chronology is the primary source. Use
Markdown. Be specific.
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_med_chron/test_context_block_rendering.py::test_custom_wrapper_template_has_context_block_placeholder -v`

Expected: PASS.

- [ ] **Step 5: Run the existing Phase 2 custom-analysis test to confirm it still passes (placeholder is unfilled but Phase 2 substitutes it next task — for now confirm the template loads)**

Run: `python -m pytest tests/test_med_chron/test_phase2_runner.py::test_custom_analysis_wraps_user_instruction_in_template -v`

Expected: This test will FAIL because Phase 2 doesn't substitute `{context_block}` yet (the assertion in the existing test only checks `{user_instruction}` is gone — but the literal `{context_block}` will now appear in the prompt). That's fine; Task 2 fixes it. Note the failure but don't fix here.

- [ ] **Step 6: Commit**

```powershell
git add Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt tests/test_med_chron/test_context_block_rendering.py
git -c commit.gpgsign=false commit -m @'
feat(med-chron): add {context_block} placeholder to custom wrapper

First half of the per-row context documents feature. Phase 2 will fill
this placeholder in the next commit; for now the template just declares
it so the renderer has somewhere to inject context.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 2: Render context block in Phase 2

**Files:**
- Modify: `Scripts/med_chron.py` — add `_render_context_block(...)`, update `_build_run_list(...)`.
- Test: `tests/test_med_chron/test_context_block_rendering.py`

- [ ] **Step 1: Write the failing tests**

Append these tests to `tests/test_med_chron/test_context_block_rendering.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_med_chron/test_context_block_rendering.py -v`

Expected: All five new tests FAIL. Most will fail because `med_chron.MAX_CONTEXT_CHARS` doesn't exist or `{context_block}` is not substituted yet. The placeholder test from Task 1 still passes.

- [ ] **Step 3: Add `MAX_CONTEXT_CHARS` and `_render_context_block` to `Scripts/med_chron.py`**

In `Scripts/med_chron.py`, locate the line `# Phase 2: Run selected analyses` (around line 447) and **insert this code immediately before that section comment**:

```python
# =============================================================================
# Context-document rendering for custom analyses
# =============================================================================

MAX_CONTEXT_CHARS = 120_000  # per-file cap to keep prompts bounded


def _render_context_block(context_files: list[str]) -> str:
    """Return the ADDITIONAL CONTEXT DOCUMENTS block, or '' if no usable files.

    Each file is extracted via ``_extract_full_text``. Failures (missing
    file, exception, empty text) are logged and silently skipped. If every
    file fails, returns ''.
    """
    if not context_files:
        return ""

    chunks: list[str] = []
    for path in context_files:
        try:
            if not os.path.exists(path):
                log_event(f"Context file missing, skipping: {path}", level="warning")
                continue
            text = _extract_full_text(path)
            if not text or not text.strip():
                log_event(f"Context file empty after extraction, skipping: {path}", level="warning")
                continue
            if len(text) > MAX_CONTEXT_CHARS:
                text = text[:MAX_CONTEXT_CHARS] + f"\n[…context truncated at {MAX_CONTEXT_CHARS:,} characters…]"
                log_event(f"Context file truncated to {MAX_CONTEXT_CHARS} chars: {path}", level="warning")
            filename = os.path.basename(path)
            chunks.append(
                f"--- BEGIN CONTEXT DOCUMENT: {filename} ---\n{text}\n--- END CONTEXT DOCUMENT ---"
            )
        except Exception as e:
            log_event(f"Failed to extract context file {path}: {e}", level="warning")
            continue

    if not chunks:
        return ""

    body = "\n\n".join(chunks)
    return f"ADDITIONAL CONTEXT DOCUMENTS PROVIDED BY THE USER:\n\n{body}"
```

- [ ] **Step 4: Update `_build_run_list` to use the renderer**

In `Scripts/med_chron.py`, find `_build_run_list` (around line 480). Locate the for-loop that builds custom analyses:

```python
    wrapper = None
    for i, c in enumerate(cfg.get("custom_analyses", []), 1):
        if wrapper is None:
            wrapper = load_prompt("_custom_wrapper.txt")
        label_slug = _slug(c["label"])
        runs.append(RunSpec(
            id=f"custom_{i}_{label_slug}",
            title=c["label"],
            prompt_text=wrapper.replace("{user_instruction}", c["instruction"]),
            input_text=full,
            output_path=os.path.join(
                output_dir, f"med_chron_custom_{i}_{label_slug}_{safe_basename}.docx"
            ),
        ))
```

Replace the inner body (everything inside the `for` loop) with this:

```python
    wrapper = None
    for i, c in enumerate(cfg.get("custom_analyses", []), 1):
        if wrapper is None:
            wrapper = load_prompt("_custom_wrapper.txt")
        label_slug = _slug(c["label"])
        context_block = _render_context_block(c.get("context_files", []) or [])
        prompt_text = wrapper.replace("{user_instruction}", c["instruction"]).replace(
            "{context_block}", context_block
        )
        runs.append(RunSpec(
            id=f"custom_{i}_{label_slug}",
            title=c["label"],
            prompt_text=prompt_text,
            input_text=full,
            output_path=os.path.join(
                output_dir, f"med_chron_custom_{i}_{label_slug}_{safe_basename}.docx"
            ),
        ))
```

- [ ] **Step 5: Run the new tests to verify they pass**

Run: `python -m pytest tests/test_med_chron/test_context_block_rendering.py -v`

Expected: All six tests PASS (1 from Task 1 + 5 new ones from Task 2).

- [ ] **Step 6: Re-run the broader Phase 2 test suite to catch regressions**

Run: `python -m pytest tests/test_med_chron/ -v`

Expected: All tests pass. The `test_custom_analysis_wraps_user_instruction_in_template` test from `test_phase2_runner.py` must now pass again (it was failing at end of Task 1) because Phase 2 substitutes `{context_block}` with empty string when no `context_files` is set.

- [ ] **Step 7: Commit**

```powershell
git add Scripts/med_chron.py tests/test_med_chron/test_context_block_rendering.py
git -c commit.gpgsign=false commit -m @'
feat(med-chron): render per-row context documents into Phase 2 prompt

_build_run_list now reads context_files from each custom_analyses entry,
extracts each file via _extract_full_text, and injects a labeled block
into the {context_block} placeholder in _custom_wrapper.txt. Missing or
unreadable files are logged and skipped; files larger than
MAX_CONTEXT_CHARS (120k) are truncated. When no usable context files are
present, the placeholder is substituted with an empty string and no
header text is emitted.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 3: Add `sniff_text_layer` helper

**Files:**
- Modify: `icharlotte_core/ui/med_chron_config_form.py` — add module-level `sniff_text_layer(path: str) -> tuple[bool, str]`.
- Test: `tests/test_wizard/test_med_chron_context_docs.py` (new file)

- [ ] **Step 1: Write the failing tests**

Create `tests/test_wizard/test_med_chron_context_docs.py` with:

```python
"""Tests for the per-row context-documents UI feature in MedChronConfigForm."""

import json
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

pytest.importorskip("pytestqt")  # NOTE: no underscore — pytest_qt silently skips

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture(autouse=True)
def _isolated_custom_analyses_store(tmp_path, monkeypatch):
    """Redirect the global custom-analyses JSON to a per-test tmp path so
    tests don't see (or pollute) the developer's real saved analyses."""
    from icharlotte_core.med_chron import custom_analyses_store
    monkeypatch.setattr(
        custom_analyses_store,
        "_STORE_PATH",
        tmp_path / "store" / "med_chron_custom_analyses.json",
    )
    yield


# -----------------------------
# Task 3: sniff_text_layer tests
# -----------------------------

def test_sniff_text_layer_txt_has_text(tmp_path):
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "ok.txt"
    p.write_text("This is a useful status report with content.", encoding="utf-8")
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is True


def test_sniff_text_layer_txt_empty(tmp_path):
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "empty.txt"
    p.write_text("   \n  ", encoding="utf-8")
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is False


def test_sniff_text_layer_docx_has_text(tmp_path):
    from docx import Document
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "doc.docx"
    doc = Document()
    doc.add_paragraph("Some legible paragraph content for the sniff to find.")
    doc.save(str(p))
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is True


def test_sniff_text_layer_docx_empty(tmp_path):
    from docx import Document
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    p = tmp_path / "blank.docx"
    Document().save(str(p))
    has_text, _ = sniff_text_layer(str(p))
    assert has_text is False


def test_sniff_text_layer_unreadable_returns_false(tmp_path):
    from icharlotte_core.ui.med_chron_config_form import sniff_text_layer
    # File that does not exist.
    has_text, reason = sniff_text_layer(str(tmp_path / "ghost.pdf"))
    assert has_text is False
    assert reason  # non-empty reason string
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -v`

Expected: All four tests FAIL with `ImportError: cannot import name 'sniff_text_layer'`.

- [ ] **Step 3: Add the helper**

In `icharlotte_core/ui/med_chron_config_form.py`, add **after the module docstring and existing imports** (around line 26, after the `from icharlotte_core.med_chron import custom_analyses_store, session_manager` line), this code:

```python
# ---- Context-document helpers ----

SUPPORTED_CONTEXT_EXTS = (".pdf", ".docx", ".txt")
MAX_CONTEXT_CHARS = 120_000  # mirrors the Phase 2 cap; only used by callers that want it


def sniff_text_layer(path: str) -> tuple[bool, str]:
    """Return (has_text, reason).

    Cheap, attach-time check: does this file appear to have extractable
    text without needing OCR? Used by the UI to warn the user.

    - .txt: any non-whitespace in the first 4 KB.
    - .docx: any non-empty paragraph in the first ~50 paragraphs.
    - .pdf: at least 200 chars of text extractable from pages 0..2.
    - any error / unknown extension: (False, "...").
    """
    import os

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".txt":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                sample = f.read(4096)
            return (bool(sample.strip()), "")
        if ext == ".docx":
            from docx import Document
            doc = Document(path)
            text = "".join(p.text for p in doc.paragraphs[:50])
            return (len(text.strip()) > 0, "")
        if ext == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(path)
            n = min(3, len(reader.pages))
            text = ""
            for i in range(n):
                try:
                    text += reader.pages[i].extract_text() or ""
                except Exception:
                    pass
            return (len(text.strip()) > 200, "")
        return (False, f"unsupported extension {ext}")
    except FileNotFoundError:
        return (False, "file not found")
    except Exception as e:
        return (False, f"could not read file: {e}")
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -v`

Expected: All four sniff tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/ui/med_chron_config_form.py tests/test_wizard/test_med_chron_context_docs.py
git -c commit.gpgsign=false commit -m @'
feat(med-chron-ui): add sniff_text_layer helper

Cheap attach-time check for whether a context document has an extractable
text layer. Used by the UI to surface an "OCR may be required" warning
without blocking the attach. Handles .pdf / .docx / .txt; any error or
unknown extension returns (False, reason).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 4: Add `ContextDropTextEdit` subclass

**Files:**
- Modify: `icharlotte_core/ui/med_chron_config_form.py` — new class above `CustomAnalysisRow`.
- Test: `tests/test_wizard/test_med_chron_context_docs.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_wizard/test_med_chron_context_docs.py`:

```python
# -----------------------------
# Task 4: ContextDropTextEdit
# -----------------------------

def _make_drag_event(qt_event_cls, urls):
    """Build a drag/drop event whose mimeData carries the given QUrl list."""
    from PySide6.QtCore import QMimeData, QPoint, Qt, QUrl
    from PySide6.QtGui import QDragEnterEvent, QDropEvent
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(u) if not isinstance(u, QUrl) else u for u in urls])
    return mime


def test_context_drop_textedit_accepts_supported_file_urls(qtbot, tmp_path):
    from PySide6.QtCore import QMimeData, QPoint, Qt, QUrl
    from PySide6.QtGui import QDragEnterEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    p1 = tmp_path / "a.pdf"
    p1.write_bytes(b"%PDF-1.4 stub")
    p2 = tmp_path / "b.docx"
    p2.write_bytes(b"docx stub")

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    received: list[list[str]] = []
    w.files_dropped.connect(lambda paths: received.append(paths))

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(p1)), QUrl.fromLocalFile(str(p2))])
    event = QDragEnterEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                            Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dragEnterEvent(event)
    assert event.isAccepted()


def test_context_drop_textedit_rejects_unsupported_file_urls(qtbot, tmp_path):
    from PySide6.QtCore import QMimeData, QPoint, Qt, QUrl
    from PySide6.QtGui import QDragEnterEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    p = tmp_path / "bad.png"
    p.write_bytes(b"\x89PNG\r\n")

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(p))])
    event = QDragEnterEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                            Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dragEnterEvent(event)
    # Not accepted as a "file drop", so the base class falls through; since
    # the base QPlainTextEdit doesn't know how to handle a PNG URL either,
    # the event is not accepted.
    assert not event.isAccepted()


def test_context_drop_textedit_emits_files_dropped(qtbot, tmp_path):
    from PySide6.QtCore import QMimeData, QPoint, Qt, QUrl
    from PySide6.QtGui import QDropEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    p = tmp_path / "good.txt"
    p.write_text("hi", encoding="utf-8")

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    received: list[list[str]] = []
    w.files_dropped.connect(lambda paths: received.append(paths))

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(p))])
    event = QDropEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                       Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dropEvent(event)
    assert received == [[str(p)]]


def test_context_drop_textedit_plaintext_drop_still_works(qtbot):
    """Dropping plain text (not file URLs) should fall through to the base
    QPlainTextEdit so the user can still drop a snippet."""
    from PySide6.QtCore import QMimeData, QPoint, Qt
    from PySide6.QtGui import QDropEvent
    from icharlotte_core.ui.med_chron_config_form import ContextDropTextEdit

    w = ContextDropTextEdit()
    qtbot.addWidget(w)

    received: list[list[str]] = []
    w.files_dropped.connect(lambda paths: received.append(paths))

    mime = QMimeData()
    mime.setText("some snippet")
    event = QDropEvent(QPoint(10, 10), Qt.DropAction.CopyAction, mime,
                       Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    w.dropEvent(event)
    # files_dropped should NOT fire for plain text
    assert received == []
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "context_drop_textedit" -v`

Expected: All four tests FAIL with `ImportError: cannot import name 'ContextDropTextEdit'`.

- [ ] **Step 3: Add the class**

In `icharlotte_core/ui/med_chron_config_form.py`, locate the existing class definition:

```python
class CustomAnalysisRow(QWidget):
```

**Immediately above it**, insert this new class:

```python
class ContextDropTextEdit(QPlainTextEdit):
    """QPlainTextEdit that accepts drops of .pdf / .docx / .txt file URLs.

    Emits ``files_dropped(list[str])`` when one or more supported files are
    dropped on the widget. Non-file drops fall through to the base class so
    text editing still works.
    """

    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def _supported_paths(self, mime) -> list[str]:
        if not mime.hasUrls():
            return []
        paths: list[str] = []
        for url in mime.urls():
            if not url.isLocalFile():
                return []
            p = url.toLocalFile()
            ext = os.path.splitext(p)[1].lower()
            if ext not in SUPPORTED_CONTEXT_EXTS:
                return []
            paths.append(p)
        return paths

    def dragEnterEvent(self, event):
        if self._supported_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if self._supported_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        paths = self._supported_paths(event.mimeData())
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
            return
        super().dropEvent(event)
```

You will also need to add the missing imports at the top of `med_chron_config_form.py`. Find the existing PySide6 import block:

```python
from PySide6.QtWidgets import (
    QCheckBox, QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit,
    QPushButton, QScrollArea, QVBoxLayout, QWidget,
)
```

Replace it with (adding `Signal` import and `QFrame`, `QFileDialog`):

```python
import os

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QCheckBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QPlainTextEdit, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "context_drop_textedit" -v`

Expected: All four ContextDropTextEdit tests PASS.

- [ ] **Step 5: Run the full med-chron-config-form test suite to catch regressions**

Run: `python -m pytest tests/test_wizard/test_med_chron_settings_page.py tests/test_wizard/test_med_chron_context_docs.py -v`

Expected: All pass. (Existing tests still construct `CustomAnalysisRow` widgets and exercise the form; we haven't touched the row yet.)

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/ui/med_chron_config_form.py tests/test_wizard/test_med_chron_context_docs.py
git -c commit.gpgsign=false commit -m @'
feat(med-chron-ui): ContextDropTextEdit accepts .pdf/.docx/.txt drops

Subclass of QPlainTextEdit that intercepts file URL drops with supported
extensions and emits files_dropped(list[str]). Any other mime data
(plain text, unsupported file types) falls through to the base class so
normal text editing keeps working.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 5: Add chip strip + add/remove on `CustomAnalysisRow`

**Files:**
- Modify: `icharlotte_core/ui/med_chron_config_form.py` — replace the instruction-textbox section of `CustomAnalysisRow.__init__`, add `add_context_files`, `_remove_context_file`, `_render_chip_strip`, `context_files()`.
- Test: `tests/test_wizard/test_med_chron_context_docs.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_wizard/test_med_chron_context_docs.py`:

```python
# -----------------------------
# Task 5: CustomAnalysisRow chip strip
# -----------------------------

def _make_row(qtbot):
    from icharlotte_core.ui.med_chron_config_form import CustomAnalysisRow
    from PySide6.QtWidgets import QWidget
    parent = QWidget()
    qtbot.addWidget(parent)
    row = CustomAnalysisRow(parent, on_remove=lambda r: None)
    return row


def test_custom_row_starts_with_no_context_files(qtbot):
    row = _make_row(qtbot)
    assert row.context_files() == []


def test_custom_row_add_context_files_appends_and_dedupes(qtbot, tmp_path):
    p1 = tmp_path / "a.txt"
    p1.write_text("hi", encoding="utf-8")
    p2 = tmp_path / "b.txt"
    p2.write_text("hi", encoding="utf-8")

    row = _make_row(qtbot)
    row.add_context_files([str(p1)])
    row.add_context_files([str(p1), str(p2)])  # p1 is a dup
    assert row.context_files() == [str(p1), str(p2)]


def test_custom_row_remove_context_file_clears_chip(qtbot, tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("hi", encoding="utf-8")

    row = _make_row(qtbot)
    row.add_context_files([str(p)])
    assert row.context_files() == [str(p)]
    row._remove_context_file(str(p))
    assert row.context_files() == []


def test_custom_row_is_empty_ignores_context_files(qtbot, tmp_path):
    """A row with only files attached (no label, no instruction) is still
    considered empty so the form doesn't try to persist it."""
    p = tmp_path / "a.txt"
    p.write_text("hi", encoding="utf-8")

    row = _make_row(qtbot)
    row.add_context_files([str(p)])
    assert row.is_empty() is True


def test_custom_row_chip_strip_renders_one_chip_per_file(qtbot, tmp_path):
    p1 = tmp_path / "a.txt"
    p1.write_text("hi", encoding="utf-8")
    p2 = tmp_path / "b.txt"
    p2.write_text("hi", encoding="utf-8")

    row = _make_row(qtbot)
    row.add_context_files([str(p1), str(p2)])

    # The strip lays chips out in row._chip_strip_layout — each chip has the
    # filename in its child QLabel. Iterate widgets and collect labels.
    from PySide6.QtWidgets import QLabel
    chip_texts = []
    for i in range(row._chip_strip_layout.count()):
        w = row._chip_strip_layout.itemAt(i).widget()
        if w is None:
            continue
        # Each chip's filename QLabel uses objectName "chip_filename"
        for child in w.findChildren(QLabel):
            if child.objectName() == "chip_filename":
                chip_texts.append(child.text())
    assert "a.txt" in chip_texts
    assert "b.txt" in chip_texts
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "custom_row" -v`

Expected: All five tests FAIL with `AttributeError: ... has no attribute 'context_files'` (or similar).

- [ ] **Step 3: Replace `CustomAnalysisRow.__init__` body**

In `icharlotte_core/ui/med_chron_config_form.py`, find the `CustomAnalysisRow` class. Replace its `__init__` method **entirely** with:

```python
    def __init__(self, parent: QWidget, on_remove):
        super().__init__(parent)
        self._on_remove = on_remove
        self._context_files: list[str] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        top = QHBoxLayout()
        self.include_cb = QCheckBox()
        self.include_cb.setChecked(True)
        self.include_cb.setToolTip("Include this analysis in the current run")
        top.addWidget(self.include_cb)
        top.addWidget(QLabel("Label:"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("Short name (e.g. 'Left-knee mentions')")
        top.addWidget(self.label_edit, 1)
        self.remove_btn = QPushButton("−")
        self.remove_btn.setFixedSize(24, 24)
        self.remove_btn.setStyleSheet("QPushButton { color: #c62828; font-weight: bold; }")
        self.remove_btn.setToolTip("Remove this analysis (also deletes from global save)")
        self.remove_btn.clicked.connect(self._handle_remove)
        top.addWidget(self.remove_btn)
        layout.addLayout(top)

        layout.addWidget(QLabel("Request:"))
        self.instruction_edit = ContextDropTextEdit()
        self.instruction_edit.setPlaceholderText("Describe the analysis…")
        self.instruction_edit.setFixedHeight(60)
        self.instruction_edit.files_dropped.connect(self.add_context_files)
        layout.addWidget(self.instruction_edit)

        # Chip strip for attached context documents.
        self._chip_strip_container = QWidget()
        self._chip_strip_layout = QHBoxLayout(self._chip_strip_container)
        self._chip_strip_layout.setContentsMargins(0, 0, 0, 0)
        self._chip_strip_layout.setSpacing(4)
        self._chip_strip_layout.addStretch()
        layout.addWidget(self._chip_strip_container)

        # Warning label for files that look like they need OCR.
        self._context_warning_label = QLabel("")
        self._context_warning_label.setStyleSheet(
            "color: #856404; background-color: #FFF3CD; "
            "padding: 4px; border-radius: 3px; font-size: 11px;"
        )
        self._context_warning_label.setWordWrap(True)
        self._context_warning_label.setVisible(False)
        layout.addWidget(self._context_warning_label)

        self._render_chip_strip()

        self.setStyleSheet(
            "CustomAnalysisRow { border: 1px solid #ddd; border-radius: 4px; }"
        )

    def label(self) -> str:
        return self.label_edit.text().strip()

    def instruction(self) -> str:
        return self.instruction_edit.toPlainText().strip()

    def is_included(self) -> bool:
        return self.include_cb.isChecked()

    def is_empty(self) -> bool:
        return not self.label() and not self.instruction()

    def context_files(self) -> list[str]:
        return list(self._context_files)

    def add_context_files(self, paths: list[str]) -> None:
        changed = False
        for p in paths:
            if p in self._context_files:
                continue
            self._context_files.append(p)
            changed = True
        if changed:
            self._render_chip_strip()
            self._refresh_context_warning()

    def _remove_context_file(self, path: str) -> None:
        if path in self._context_files:
            self._context_files.remove(path)
            self._render_chip_strip()
            self._refresh_context_warning()

    def _render_chip_strip(self) -> None:
        # Remove all widgets currently in the strip (preserve trailing stretch).
        while self._chip_strip_layout.count() > 0:
            item = self._chip_strip_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()
        # One chip per attached file.
        for p in self._context_files:
            chip = self._build_chip(p)
            self._chip_strip_layout.addWidget(chip)
        # Trailing "+ Add context" button (wired up in Task 6).
        if not hasattr(self, "_add_ctx_btn") or self._add_ctx_btn is None:
            self._add_ctx_btn = QPushButton("+ Add context")
            self._add_ctx_btn.setStyleSheet(
                "QPushButton { font-size: 11px; padding: 2px 8px; }"
            )
        self._chip_strip_layout.addWidget(self._add_ctx_btn)
        self._chip_strip_layout.addStretch()

    def _build_chip(self, path: str) -> QWidget:
        chip = QFrame()
        chip.setObjectName("ctx_chip")
        chip.setStyleSheet(
            "QFrame#ctx_chip { background: #E3F2FD; border: 1px solid #90CAF9; "
            "border-radius: 10px; padding: 2px 6px; }"
        )
        lay = QHBoxLayout(chip)
        lay.setContentsMargins(4, 0, 4, 0)
        lay.setSpacing(4)
        icon = QLabel("📎")
        icon.setStyleSheet("font-size: 11px;")
        lay.addWidget(icon)
        name = QLabel(os.path.basename(path))
        name.setObjectName("chip_filename")
        name.setStyleSheet("font-size: 11px; color: #0D47A1;")
        name.setToolTip(path)
        lay.addWidget(name)
        x_btn = QPushButton("✕")
        x_btn.setFixedSize(16, 16)
        x_btn.setStyleSheet(
            "QPushButton { font-size: 10px; color: #555; border: none; }"
            "QPushButton:hover { color: #c62828; }"
        )
        x_btn.setToolTip(f"Remove {os.path.basename(path)}")
        x_btn.clicked.connect(lambda _=False, p=path: self._remove_context_file(p))
        lay.addWidget(x_btn)
        return chip

    def _refresh_context_warning(self) -> None:
        bad: list[str] = []
        for p in self._context_files:
            has_text, _ = sniff_text_layer(p)
            if not has_text:
                bad.append(os.path.basename(p))
        if bad:
            self._context_warning_label.setText(
                "⚠ Likely needs OCR at run time: " + ", ".join(bad)
            )
            self._context_warning_label.setVisible(True)
        else:
            self._context_warning_label.setVisible(False)
```

**Then delete the old standalone methods** that now live inside the new `__init__` body — i.e., remove the old `_handle_remove`, `label`, `instruction`, `is_included`, `is_empty` methods that appeared below the old `__init__`. (After the replacement above, the only method remaining outside is `_handle_remove` — keep that one. Re-read the file in your IDE to confirm structure.)

Actually, simpler: the replacement above intentionally keeps `_handle_remove` only via the existing pattern, so the only standalone method that needs to survive untouched in the class body is `_handle_remove`. If you replaced the whole class body, add back:

```python
    def _handle_remove(self):
        self._on_remove(self)
```

Verify by skimming the file: `CustomAnalysisRow` should now contain `__init__`, `label`, `instruction`, `is_included`, `is_empty`, `context_files`, `add_context_files`, `_remove_context_file`, `_render_chip_strip`, `_build_chip`, `_refresh_context_warning`, `_handle_remove`. No duplicate definitions.

- [ ] **Step 4: Run the new tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "custom_row" -v`

Expected: All five new tests PASS.

- [ ] **Step 5: Run the existing form tests to catch regressions**

Run: `python -m pytest tests/test_wizard/test_med_chron_settings_page.py tests/test_wizard/test_med_chron_context_docs.py -v`

Expected: All tests pass. The existing settings-page tests construct `CustomAnalysisRow` widgets via the form's `add_custom_row()` method and check label/instruction methods that we preserved.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/ui/med_chron_config_form.py tests/test_wizard/test_med_chron_context_docs.py
git -c commit.gpgsign=false commit -m @'
feat(med-chron-ui): chip strip + add/remove for per-row context docs

CustomAnalysisRow now embeds a ContextDropTextEdit for the instruction
field, a horizontal chip strip showing one chip per attached file (with
an X to remove), and a yellow warning label that lights up when one of
the attached files fails the text-layer sniff. Drops on the textbox are
routed to add_context_files via the files_dropped signal. The
"+ Add context" button is rendered but not wired yet — Task 6 hooks up
QFileDialog.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 6: Wire "+ Add context" button to `QFileDialog`

**Files:**
- Modify: `icharlotte_core/ui/med_chron_config_form.py` — connect `_add_ctx_btn` clicked signal.
- Test: `tests/test_wizard/test_med_chron_context_docs.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_med_chron_context_docs.py`:

```python
# -----------------------------
# Task 6: + Add context button
# -----------------------------

def test_add_context_button_opens_filedialog_and_attaches(qtbot, tmp_path, monkeypatch):
    """Clicking '+ Add context' opens QFileDialog.getOpenFileNames and any
    selected paths are appended to context_files."""
    p = tmp_path / "picked.pdf"
    p.write_bytes(b"%PDF-1.4")

    # Stub QFileDialog.getOpenFileNames to return our test file.
    from icharlotte_core.ui import med_chron_config_form
    monkeypatch.setattr(
        med_chron_config_form.QFileDialog,
        "getOpenFileNames",
        staticmethod(lambda *a, **kw: ([str(p)], "")),
    )

    row = _make_row(qtbot)
    assert row.context_files() == []
    row._add_ctx_btn.click()
    assert row.context_files() == [str(p)]


def test_add_context_button_cancelled_does_nothing(qtbot, monkeypatch):
    from icharlotte_core.ui import med_chron_config_form
    monkeypatch.setattr(
        med_chron_config_form.QFileDialog,
        "getOpenFileNames",
        staticmethod(lambda *a, **kw: ([], "")),
    )

    row = _make_row(qtbot)
    row._add_ctx_btn.click()
    assert row.context_files() == []
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "add_context_button" -v`

Expected: Tests FAIL — clicking `_add_ctx_btn` does nothing because no slot is wired.

- [ ] **Step 3: Wire the button**

In `icharlotte_core/ui/med_chron_config_form.py`, locate `CustomAnalysisRow._render_chip_strip`. Find the lines that create `_add_ctx_btn`:

```python
        if not hasattr(self, "_add_ctx_btn") or self._add_ctx_btn is None:
            self._add_ctx_btn = QPushButton("+ Add context")
            self._add_ctx_btn.setStyleSheet(
                "QPushButton { font-size: 11px; padding: 2px 8px; }"
            )
```

Replace with:

```python
        if not hasattr(self, "_add_ctx_btn") or self._add_ctx_btn is None:
            self._add_ctx_btn = QPushButton("+ Add context")
            self._add_ctx_btn.setStyleSheet(
                "QPushButton { font-size: 11px; padding: 2px 8px; }"
            )
            self._add_ctx_btn.clicked.connect(self._on_add_context_clicked)
```

And add a new method to `CustomAnalysisRow` (place it right after `add_context_files`):

```python
    def _on_add_context_clicked(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select context document(s)",
            "",
            "Context documents (*.pdf *.docx *.txt)",
        )
        if paths:
            self.add_context_files(list(paths))
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "add_context_button" -v`

Expected: Both tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/ui/med_chron_config_form.py tests/test_wizard/test_med_chron_context_docs.py
git -c commit.gpgsign=false commit -m @'
feat(med-chron-ui): wire + Add context to QFileDialog

Clicking the button opens QFileDialog.getOpenFileNames filtered to
.pdf/.docx/.txt and appends the picked paths via add_context_files.
Cancelling is a no-op.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 7: Verify warning label updates after attach

**Files:**
- (No code change — the wiring is already in place from Task 5.)
- Test: `tests/test_wizard/test_med_chron_context_docs.py`

- [ ] **Step 1: Write the test**

Append to `tests/test_wizard/test_med_chron_context_docs.py`:

```python
# -----------------------------
# Task 7: warning label
# -----------------------------

def test_warning_label_visible_when_file_lacks_text_layer(qtbot, tmp_path):
    # An empty .txt fails sniff_text_layer.
    p = tmp_path / "blank.txt"
    p.write_text("", encoding="utf-8")

    row = _make_row(qtbot)
    row.add_context_files([str(p)])
    assert row._context_warning_label.isVisible() is True
    assert "blank.txt" in row._context_warning_label.text()


def test_warning_label_hidden_when_all_files_have_text(qtbot, tmp_path):
    p = tmp_path / "good.txt"
    p.write_text("real content", encoding="utf-8")

    row = _make_row(qtbot)
    row.add_context_files([str(p)])
    assert row._context_warning_label.isVisible() is False


def test_warning_label_hides_after_bad_file_removed(qtbot, tmp_path):
    bad = tmp_path / "blank.txt"
    bad.write_text("", encoding="utf-8")
    good = tmp_path / "good.txt"
    good.write_text("real content", encoding="utf-8")

    row = _make_row(qtbot)
    row.add_context_files([str(bad), str(good)])
    assert row._context_warning_label.isVisible() is True

    row._remove_context_file(str(bad))
    assert row._context_warning_label.isVisible() is False
```

- [ ] **Step 2: Run the tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "warning_label" -v`

Expected: All three tests PASS (because Task 5 already wired `_refresh_context_warning` into `add_context_files` and `_remove_context_file`).

If any test fails, double-check that in Task 5 both `add_context_files` and `_remove_context_file` call `self._refresh_context_warning()` at the end.

- [ ] **Step 3: Commit**

```powershell
git add tests/test_wizard/test_med_chron_context_docs.py
git -c commit.gpgsign=false commit -m @'
test(med-chron-ui): verify OCR-warning label visibility logic

Three tests covering: visible when an attached file has no text layer,
hidden when all attached files do, and re-hides when the offending file
is removed. No production code changes — the wiring is already in place
from Task 5; these tests document and lock the behavior.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 8: Dual-shape commit (persisted vs run)

**Files:**
- Modify: `icharlotte_core/ui/med_chron_config_form.py` — change `_validated_custom_rows` return shape and `commit_user_config` body.
- Test: `tests/test_wizard/test_med_chron_context_docs.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_wizard/test_med_chron_context_docs.py`:

```python
# -----------------------------
# Task 8: dual-shape commit
# -----------------------------

def _write_session(tmp_path, *, narrative_missing=False):
    cache = tmp_path / ".med_chron" / "abc123"
    cache.mkdir(parents=True)
    session_path = cache / "session.json"
    session_path.write_text(json.dumps({
        "version": 1,
        "phase": "awaiting_input",
        "input_path": str(tmp_path / "rec.docx"),
        "narrative_text_path": str(cache / "narrative.txt"),
        "full_text_path": str(cache / "full.txt"),
        "narrative_missing": narrative_missing,
        "provider_name": "Acme PT",
        "file_number": "1234.567",
        "catalog": [
            {"id": "rewrite_chronology", "title": "Rewrite Chronology",
             "description": "...", "uses_tables": False,
             "default_selected": True},
        ],
        "user_config": None,
    }, indent=2), encoding="utf-8")
    return session_path


def test_commit_writes_context_files_to_session(qtbot, tmp_path):
    """When a custom analysis has attached context files, commit_user_config
    must include them in the session JSON's user_config.custom_analyses."""
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session_path = _write_session(tmp_path)
    ctx = tmp_path / "status.txt"
    ctx.write_text("ctx", encoding="utf-8")

    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    row = form.add_custom_row()
    row.label_edit.setText("Defense targets")
    row.instruction_edit.setPlainText("Identify providers.")
    row.add_context_files([str(ctx)])

    assert form.commit_user_config() is True

    written = json.loads(session_path.read_text(encoding="utf-8"))
    customs = written["user_config"]["custom_analyses"]
    assert len(customs) == 1
    assert customs[0]["context_files"] == [str(ctx)]
    assert customs[0]["label"] == "Defense targets"


def test_commit_does_not_persist_context_files_to_global_store(qtbot, tmp_path):
    """The global custom_analyses_store must continue to hold only
    {label, instruction} — never context_files."""
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    from icharlotte_core.med_chron import custom_analyses_store
    session_path = _write_session(tmp_path)
    ctx = tmp_path / "status.txt"
    ctx.write_text("ctx", encoding="utf-8")

    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    row = form.add_custom_row()
    row.label_edit.setText("Defense targets")
    row.instruction_edit.setPlainText("Identify providers.")
    row.add_context_files([str(ctx)])

    form.commit_user_config()

    saved = custom_analyses_store.load()
    assert len(saved) == 1
    assert "context_files" not in saved[0]
    assert saved[0] == {"label": "Defense targets", "instruction": "Identify providers."}


def test_reopening_form_loads_saved_analysis_with_empty_context(qtbot, tmp_path):
    """After a commit, opening a NEW form against a NEW session shows the
    persisted label/instruction but starts with no attached context files."""
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    session1 = _write_session(tmp_path)
    ctx = tmp_path / "status.txt"
    ctx.write_text("ctx", encoding="utf-8")

    form1 = MedChronConfigForm(session1)
    qtbot.addWidget(form1)
    row = form1.add_custom_row()
    row.label_edit.setText("Defense targets")
    row.instruction_edit.setPlainText("Identify providers.")
    row.add_context_files([str(ctx)])
    form1.commit_user_config()

    # Build a second session in a different tmp subdir.
    session2_dir = tmp_path / "session2"
    session2_dir.mkdir()
    session2 = _write_session(session2_dir)

    form2 = MedChronConfigForm(session2)
    qtbot.addWidget(form2)

    # Pre-populated row should exist with persisted text, but no context files.
    assert len(form2.custom_rows) == 1
    row2 = form2.custom_rows[0]
    assert row2.label() == "Defense targets"
    assert row2.instruction() == "Identify providers."
    assert row2.context_files() == []


def test_commit_omits_unchecked_rows_from_session(qtbot, tmp_path):
    """If the include checkbox is unchecked, the row is persisted globally
    but NOT included in session.json's run-shape list."""
    from icharlotte_core.ui.med_chron_config_form import MedChronConfigForm
    from icharlotte_core.med_chron import custom_analyses_store
    session_path = _write_session(tmp_path)
    ctx = tmp_path / "status.txt"
    ctx.write_text("ctx", encoding="utf-8")

    form = MedChronConfigForm(session_path)
    qtbot.addWidget(form)
    row = form.add_custom_row()
    row.label_edit.setText("Defense targets")
    row.instruction_edit.setPlainText("Identify providers.")
    row.add_context_files([str(ctx)])
    row.include_cb.setChecked(False)

    # Need at least ONE thing checked, or commit fails validation.
    # The default rewrite_chronology checkbox is already checked.
    form.commit_user_config()

    written = json.loads(session_path.read_text(encoding="utf-8"))
    assert written["user_config"]["custom_analyses"] == []

    # But persisted to global store.
    saved = custom_analyses_store.load()
    assert len(saved) == 1
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -k "commit" or "reopening" -v`

Expected: `test_commit_writes_context_files_to_session` and `test_reopening_form_loads_saved_analysis_with_empty_context` FAIL because the form doesn't yet write `context_files` to the session. The other two may pass already if luck-of-the-draw.

- [ ] **Step 3: Update `_validated_custom_rows` to return dual shapes**

In `icharlotte_core/ui/med_chron_config_form.py`, find `_validated_custom_rows`:

```python
    def _validated_custom_rows(self) -> tuple[list[dict], list[dict], str]:
        """Return ``(all_valid_rows, included_rows, error_msg)``.
        ...
        """
        all_valid: list[dict] = []
        included: list[dict] = []
        for r in self.custom_rows:
            if r.is_empty():
                continue
            lbl, instr = r.label(), r.instruction()
            if not lbl or not instr:
                return [], [], (
                    "Custom analyses need both a label and an instruction. "
                    "Fill in (or remove) the partially-completed row."
                )
            entry = {"label": lbl, "instruction": instr}
            all_valid.append(entry)
            if r.is_included():
                included.append(entry)
        return all_valid, included, ""
```

Replace it with:

```python
    def _validated_custom_rows(self) -> tuple[list[dict], list[dict], str]:
        """Return ``(persisted_rows, run_rows, error_msg)``.

        - ``persisted_rows`` — ``{label, instruction}`` only, for the global
          custom_analyses_store. Context files are intentionally excluded so
          they do not leak between sessions / cases.
        - ``run_rows`` — ``{label, instruction, context_files}`` for the
          session JSON; only rows whose include checkbox is checked.
        - ``error_msg`` — non-empty if a row is partially filled.
        """
        persisted: list[dict] = []
        run_rows: list[dict] = []
        for r in self.custom_rows:
            if r.is_empty():
                continue
            lbl, instr = r.label(), r.instruction()
            if not lbl or not instr:
                return [], [], (
                    "Custom analyses need both a label and an instruction. "
                    "Fill in (or remove) the partially-completed row."
                )
            persisted.append({"label": lbl, "instruction": instr})
            if r.is_included():
                run_rows.append({
                    "label": lbl,
                    "instruction": instr,
                    "context_files": r.context_files(),
                })
        return persisted, run_rows, ""
```

- [ ] **Step 4: Update `commit_user_config` to use the renamed shapes**

In `icharlotte_core/ui/med_chron_config_form.py`, find `commit_user_config`. Inside it, the only thing that changes is the variable names (no behavioral change). Locate this block:

```python
        selected = self._selected_catalog_ids()
        all_valid_custom, included_custom, err = self._validated_custom_rows()
        if err:
            self._error_label.setText(err)
            self._error_label.setVisible(True)
            return False
        if not selected and not included_custom:
            self._error_label.setText(
                "Select at least one analysis, or add a custom analysis."
            )
            self._error_label.setVisible(True)
            return False

        # Auto-save the full list of valid custom rows (including unchecked
        # ones) to the global store so they reappear next time.
        try:
            custom_analyses_store.save(all_valid_custom)
        except OSError:
            # Persisting globally is best-effort; the run can still proceed.
            pass

        session_manager.update_user_config(
            self.session_path,
            {
                "selected_catalog_ids": selected,
                "custom_analyses": included_custom,
            },
        )
        return True
```

Replace with:

```python
        selected = self._selected_catalog_ids()
        persisted_custom, run_custom, err = self._validated_custom_rows()
        if err:
            self._error_label.setText(err)
            self._error_label.setVisible(True)
            return False
        if not selected and not run_custom:
            self._error_label.setText(
                "Select at least one analysis, or add a custom analysis."
            )
            self._error_label.setVisible(True)
            return False

        # Persist label + instruction only to the global store. Context
        # files are intentionally NOT persisted — they belong to this
        # session / case only.
        try:
            custom_analyses_store.save(persisted_custom)
        except OSError:
            # Persisting globally is best-effort; the run can still proceed.
            pass

        session_manager.update_user_config(
            self.session_path,
            {
                "selected_catalog_ids": selected,
                "custom_analyses": run_custom,
            },
        )
        return True
```

- [ ] **Step 5: Run the new tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_med_chron_context_docs.py -v`

Expected: All Task 8 tests PASS. All earlier tests still pass.

- [ ] **Step 6: Run the full wizard test suite to catch regressions**

Run: `python -m pytest tests/test_wizard/ -v`

Expected: All wizard tests pass.

- [ ] **Step 7: Run the full med_chron suite as well**

Run: `python -m pytest tests/test_med_chron/ -v`

Expected: All med_chron tests pass.

- [ ] **Step 8: Commit**

```powershell
git add icharlotte_core/ui/med_chron_config_form.py tests/test_wizard/test_med_chron_context_docs.py
git -c commit.gpgsign=false commit -m @'
feat(med-chron-ui): dual-shape commit — persist text, run with context

_validated_custom_rows now returns two parallel lists: persisted_rows
(label + instruction only, written to the global custom_analyses_store)
and run_rows (label + instruction + context_files, written to the
session JSON). This is what guarantees context files stay scoped to a
single run — when the form reopens against a new session, saved
analyses come back with empty file lists.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
'@
```

---

## Task 9: Manual smoke test

**Files:** None.

Per the global CLAUDE.md "Always test after developing or changing a feature" rule, manually verify the feature end-to-end in the running iCharlotte UI.

- [ ] **Step 1: Launch the app**

Run: `python iCharlotte.py`

- [ ] **Step 2: Open the wizard for a case**

In the app:
1. Pick any case from the Master Case Tab.
2. Switch to the Wizard tab.
3. Add the **Med Chron Analysis** task.
4. On the inline file picker for that task, select an existing medical chronology `.docx` (one that has BRIEF SYNOPSIS pre/post-injury sections).
5. Wait for Phase 1 (the "Preparing chronology…" spinner) to finish.

- [ ] **Step 3: Add a custom analysis with context**

1. Click **+ Add custom analysis**.
2. Type a label, e.g., `Defense deposition targets`.
3. Type an instruction, e.g., `Identify treatment providers worth deposing for the defense, using the uploaded status report as context.`
4. **Drag a small `.pdf` or `.docx` from Explorer onto the instruction textbox** — confirm a chip appears below with the filename.
5. Click **+ Add context** — confirm `QFileDialog` opens, filtered to `*.pdf *.docx *.txt`. Pick another file.
6. **Drop an unsupported file** (e.g., `.png`) onto the textbox — confirm NO chip is added and that the textbox doesn't visibly accept the drop (or that text-drop behavior is unaffected).
7. **Attach an image-only PDF** if you have one — confirm the yellow warning label appears listing it.
8. Click **✕** on one of the chips — confirm it disappears.

- [ ] **Step 4: Run Phase 2**

1. Click **Proceed**.
2. Wait for Phase 2 to finish.
3. Open the generated `med_chron_custom_1_<slug>_<filename>.docx` from the case's `NOTES/AI OUTPUT` folder.
4. Spot-check the output: does it reference content from the context document? It should — even briefly. (If the LLM ignored the context entirely, double-check the prompt was wired by reading the most recent run's log: `Med_Chron_activity.log` in the project root.)

- [ ] **Step 5: Confirm persistence semantics**

1. Close the wizard task tab.
2. Re-open it (re-add the Med Chron Analysis task; pick a chronology again).
3. Once Phase 1 completes and the picker appears, confirm:
   - The custom analysis label + instruction text **are pre-filled** from the global store.
   - The context-files chip strip **is empty** — none of the previously-attached files come back.

- [ ] **Step 6: Confirm cross-case isolation**

1. Switch to a different case.
2. Add the Med Chron Analysis task, pick a chronology.
3. Confirm the custom analysis text reappears (it's global) but the chip strip is empty.

- [ ] **Step 7: Final commit (if any tweaks were needed)**

If the smoke test surfaced bugs, fix them with targeted patches + tests, then commit. If everything worked, no commit needed.

---

## Done

After Task 9 passes:

- Per-row context documents work for custom Med-Cron analyses.
- Files attach via drag-drop on the instruction textbox or via the **+ Add context** button.
- Instruction text persists across runs/cases; context files do NOT.
- A warning label appears for attached files that likely need OCR at run time.
- Phase 2 truncates oversized files at 120k chars and silently skips missing/unreadable ones.

**Files touched at end:**
- `Scripts/MED_CHRON_ANALYSES/prompts/_custom_wrapper.txt`
- `Scripts/med_chron.py`
- `icharlotte_core/ui/med_chron_config_form.py`
- `tests/test_med_chron/test_context_block_rendering.py` (new)
- `tests/test_wizard/test_med_chron_context_docs.py` (new)

**Spec:** [docs/superpowers/specs/2026-05-19-med-chron-context-docs-design.md](../specs/2026-05-19-med-chron-context-docs-design.md)
