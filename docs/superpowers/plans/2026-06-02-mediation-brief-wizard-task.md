# Mediation Brief — Wizard Task Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Mediation Brief" Wizard task card that drives the existing, unmodified `MediationBriefGenerator` through a Settings → Status → Output flow, saving the brief to `NOTES/AI OUTPUT/MEDIATION` with a Save-a-Copy-As escape hatch.

**Architecture:** A new in-process Wizard task. One new module (`pages/mediation_brief_page.py`) holds a bespoke settings page, a `QThread` worker that runs the generator pipeline, a bespoke read-only output page, and a `WizardTaskContainer` tab (modeled on `GenerateMotionTaskTab`). Four small edits register and route the task. The generator, the Chat-tab feature, and the Word AI Assistant are untouched.

**Tech Stack:** Python 3, PySide6, python-docx, PyMuPDF (fitz), pytest + pytest-qt. Qt binding is **PySide6** (the app uses PySide6, not PyQt6).

**Spec:** `docs/superpowers/specs/2026-06-02-mediation-brief-wizard-task-design.md`

**Branch:** `feature/mediation-brief-wizard-task` (already created; the spec is committed there).

---

## File Structure

- **Create** `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` — settings page, worker, save helper, document-reader helpers, output page, tab, builder. Single responsibility: everything specific to the Mediation Brief Wizard task.
- **Create** `tests/test_wizard/test_mediation_brief_page.py` — all tests for the task.
- **Modify** `icharlotte_core/ui/wizard/registry.py` — add the `mediation_brief` `TaskSpec`.
- **Modify** `icharlotte_core/ui/wizard/task_routing.py` — add one entry to `_IN_PROCESS_TASK_BUILDERS`.
- **Modify** `icharlotte_core/ui/wizard/in_process_task_tab.py` — add the thin `build_mediation_brief_tab` wrapper.
- **Modify** `iCharlotte.py` — add `"MediationBriefTaskTab"` to the close-guard class tuple.

**Reused (do NOT modify):** `icharlotte_core/mediation_brief.py` (`MediationBriefGenerator`, `GENERATION_ORDER`, `SECTION_HEADINGS`), `ui/wizard/task_scaffold.WizardTaskContainer`, `ui/wizard/pages/status_page.StatusPage`, `ui/wizard/theme`, `ui/wizard/docx_io.load_docx_as_html`, `ui/context_files_dialog.ContextFilesDialog`, `document_processor.extract_docx_text`.

> **Testing note:** Running iCharlotte interferes with PySide6 imports during pytest collection. Run the test suite with the app **closed**. This repo is a concurrent multi-session checkout — `git add` only the specific files listed in each commit step.

---

### Task 1: Register the task in the registry

**Files:**
- Modify: `icharlotte_core/ui/wizard/registry.py` (insert into `TASK_REGISTRY`, before the `"chat"` entry ~line 228)
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_wizard/test_mediation_brief_page.py` with:

```python
"""Tests for the Mediation Brief Wizard task."""
import os

import pytest


# ---- Registry / routing (pure logic, no Qt) ----

def test_mediation_brief_registered():
    from icharlotte_core.ui.wizard.registry import TASK_REGISTRY

    spec = TASK_REGISTRY["mediation_brief"]
    assert spec.title == "Mediation Brief"
    assert spec.category == "Motions & Drafting"
    assert spec.script_name == ""
    assert "mediation" in spec.keywords


def test_mediation_brief_has_valid_category():
    from icharlotte_core.ui.wizard.registry import CATEGORY_ORDER, TASK_REGISTRY

    assert TASK_REGISTRY["mediation_brief"].category in CATEGORY_ORDER
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -v`
Expected: FAIL with `KeyError: 'mediation_brief'`.

- [ ] **Step 3: Add the registry entry**

In `icharlotte_core/ui/wizard/registry.py`, insert this entry into the `TASK_REGISTRY` dict immediately before the `"chat"` entry:

```python
    "mediation_brief": TaskSpec(
        task_id="mediation_brief",
        title="Mediation Brief",
        description="Generate a defense-side mediation brief from case documents.",
        icon_glyph="\U0001F91D",  # 🤝
        script_name="",  # in-process worker
        category="Motions & Drafting",
        keywords=[
            "mediation", "brief", "settlement", "mediator",
            "defense", "confidential",
        ],
        default_folders=[],
    ),
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Run the category tests to confirm no regression**

Run: `python -m pytest tests/test_wizard/test_task_categories.py -v`
Expected: PASS (all). (Assertions are dynamic against `list_tasks()`; Discovery's count is unaffected.)

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/registry.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): register Mediation Brief task in registry"
```

---

### Task 2: Route the task as an in-process builder

**Files:**
- Modify: `icharlotte_core/ui/wizard/task_routing.py:4-10` (the `_IN_PROCESS_TASK_BUILDERS` dict)
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py` (append wrapper at end of file)
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_mediation_brief_page.py`:

```python
def test_mediation_brief_is_in_process():
    from icharlotte_core.ui.wizard.task_routing import (
        get_in_process_task_builder_name,
        is_in_process_task,
        requires_initial_file_picker,
    )

    assert get_in_process_task_builder_name("mediation_brief") == "build_mediation_brief_tab"
    assert is_in_process_task("mediation_brief") is True
    # In-process tasks own their source selection — no pre-Settings picker.
    assert requires_initial_file_picker("mediation_brief") is False


def test_build_mediation_brief_tab_attribute_exists():
    from icharlotte_core.ui.wizard import in_process_task_tab

    assert hasattr(in_process_task_tab, "build_mediation_brief_tab")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py::test_mediation_brief_is_in_process -v`
Expected: FAIL (`get_in_process_task_builder_name` returns `None`).

- [ ] **Step 3: Add the routing entry**

In `icharlotte_core/ui/wizard/task_routing.py`, add the `mediation_brief` line to `_IN_PROCESS_TASK_BUILDERS`:

```python
_IN_PROCESS_TASK_BUILDERS = {
    "subpoena_tracker": "build_subpoena_tab",
    "respond_to_discovery": "build_respond_to_discovery_tab",
    "oppose_motion": "build_oppose_motion_tab",
    "separate": "build_separate_tab",
    "generate_motion": "build_generate_motion_tab",
    "mediation_brief": "build_mediation_brief_tab",
}
```

- [ ] **Step 4: Add the thin builder wrapper**

Append to the end of `icharlotte_core/ui/wizard/in_process_task_tab.py` (lazy import inside the function so this attribute exists even before the page module is written):

```python
def build_mediation_brief_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None,
):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import (
        build_mediation_brief_tab as _build,
    )

    return _build(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -v`
Expected: PASS (4 tests). `test_build_mediation_brief_tab_attribute_exists` passes because `hasattr` does not call the function (the page module need not exist yet).

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/task_routing.py icharlotte_core/ui/wizard/in_process_task_tab.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): route mediation_brief to build_mediation_brief_tab"
```

---

### Task 3: Create the page module with the document reader

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/mediation_brief_page.py`
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_mediation_brief_page.py`:

```python
# ---- Document reader ----

pytestqt = pytest.importorskip("pytestqt")  # ensures PySide6/Qt present for module import


def test_read_documents_txt_and_docx(tmp_path):
    from docx import Document

    from icharlotte_core.ui.wizard.pages.mediation_brief_page import _read_documents

    txt = tmp_path / "note.txt"
    txt.write_text("Plain text body", encoding="utf-8")

    docx_path = tmp_path / "summary.docx"
    doc = Document()
    doc.add_paragraph("Intro paragraph text")
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "CellA"
    table.rows[0].cells[1].text = "CellB"
    doc.save(str(docx_path))

    content, warnings = _read_documents([str(txt), str(docx_path)])

    assert "Plain text body" in content
    assert "Intro paragraph text" in content
    # Table-aware extraction must surface cell text (no silent data loss).
    assert "CellA" in content and "CellB" in content
    assert warnings == []


def test_read_documents_reports_missing_file(tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import _read_documents

    content, warnings = _read_documents([str(tmp_path / "nope.pdf")])

    assert content == ""
    assert any("not found" in w.lower() for w in warnings)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py::test_read_documents_txt_and_docx -v`
Expected: FAIL (`ModuleNotFoundError: ...mediation_brief_page`).

- [ ] **Step 3: Create the module with imports and the reader helpers**

Create `icharlotte_core/ui/wizard/pages/mediation_brief_page.py`:

```python
"""Wizard task: defense-side Mediation Brief generation.

Drives the existing ``MediationBriefGenerator`` (unchanged) through a
Settings -> Status -> Output flow. Generation only; section refinement and
deposition-quote insertion are provided by the Word AI Assistant (Win+V) on
the generated document.
"""
from __future__ import annotations

import logging
import os
import re
import shutil

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.mediation_brief import (
    GENERATION_ORDER,
    SECTION_HEADINGS,
    MediationBriefGenerator,
)
from icharlotte_core.ui.context_files_dialog import ContextFilesDialog
from icharlotte_core.ui.wizard import theme
from icharlotte_core.ui.wizard.docx_io import load_docx_as_html
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer

logger = logging.getLogger(__name__)

TASK_PAGE_SETTINGS = 0
TASK_PAGE_STATUS = 1
TASK_PAGE_OUTPUT = 2

DEFAULT_BRIEF_FILENAME = "Defendant's Confidential Mediation Brief.docx"
MEDIATION_SUBDIR = ("NOTES", "AI OUTPUT", "MEDIATION")


# --------------------------- document extraction ---------------------------

def _read_pdf(path: str) -> str:
    import fitz

    parts = []
    doc = fitz.open(path)
    try:
        for page in doc:
            parts.append(page.get_text())
    finally:
        doc.close()
    return "\n".join(parts)


def _read_doc_via_word(path: str) -> str:
    """Read a legacy .doc via read-only Word COM. Never Quit/Visible the user's
    Word (global safety rule); only close the document we open."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = None
        try:
            doc = word.Documents.Open(
                FileName=os.path.abspath(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            return doc.Content.Text or ""
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
    finally:
        pythoncom.CoUninitialize()


def _read_msg(path: str) -> str:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        item = namespace.OpenSharedItem(os.path.abspath(path))
        try:
            subject = item.Subject or ""
            sender = item.SenderName or ""
            body = item.Body or ""
            if not body.strip() and item.HTMLBody:
                body = re.sub(r"<[^>]+>", "", item.HTMLBody).strip()
            return f"From: {sender}\nSubject: {subject}\n\n{body}"
        finally:
            try:
                item.Close(0)
            except Exception:
                pass
    finally:
        pythoncom.CoUninitialize()


def _read_documents(paths) -> tuple[str, list[str]]:
    """Extract text from the given files into a single string + warning list.

    Table-aware for .docx (uses ``extract_docx_text`` — legal chronologies and
    depo summaries live inside tables; a ``doc.paragraphs`` loop loses them).
    Mirrors the Chat-tab reader so the Wizard feeds the engine identical input.
    """
    content_parts: list[str] = []
    warnings: list[str] = []
    for path in paths or []:
        if not path or not os.path.isfile(path):
            warnings.append(f"File not found: {path}")
            continue
        ext = os.path.splitext(path)[1].lower()
        name = os.path.basename(path)
        try:
            if ext == ".txt":
                with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                    text = fh.read()
            elif ext == ".docx":
                from icharlotte_core.document_processor import extract_docx_text

                text = extract_docx_text(path)
            elif ext == ".doc":
                text = _read_doc_via_word(path)
            elif ext == ".pdf":
                text = _read_pdf(path)
            elif ext == ".msg":
                text = _read_msg(path)
            else:
                warnings.append(f"Unsupported file type skipped: {name}")
                continue
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"Could not read {name}: {exc}")
            continue
        if text and text.strip():
            content_parts.append(f"--- FILE: {name} ---\n{text}")
        else:
            warnings.append(f"No text extracted from {name}")
    return "\n\n".join(content_parts), warnings
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k read_documents -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/mediation_brief_page.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): mediation brief page module + document reader"
```

---

### Task 4: Add the lock-safe save helper

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` (append `_save_brief`)
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_mediation_brief_page.py`:

```python
# ---- Save helper ----

def test_save_brief_writes_to_target(tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import _save_brief

    class FakeGen:
        def assemble_document(self, caption, out):
            with open(out, "w", encoding="utf-8") as fh:
                fh.write("brief")

    dest = str(tmp_path / "MEDIATION")
    path = _save_brief(FakeGen(), "caption.docx", dest, "Brief.docx")

    assert os.path.basename(path) == "Brief.docx"
    assert os.path.isfile(path)


def test_save_brief_falls_back_when_destination_locked(tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import _save_brief

    class LockGen:
        def assemble_document(self, caption, out):
            if os.path.basename(out) == "Brief.docx":
                raise PermissionError("file is open in Word")
            with open(out, "w", encoding="utf-8") as fh:
                fh.write("brief")

    dest = str(tmp_path / "MEDIATION")
    path = _save_brief(LockGen(), "caption.docx", dest, "Brief.docx")

    assert os.path.basename(path) == "Brief (2).docx"
    assert os.path.isfile(path)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k save_brief -v`
Expected: FAIL (`ImportError: cannot import name '_save_brief'`).

- [ ] **Step 3: Implement `_save_brief`**

Append to `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` (after `_read_documents`):

```python
# ------------------------------- assembly ---------------------------------

def _save_brief(generator, caption_path: str, dest_dir: str, filename: str) -> str:
    """Assemble the brief into ``dest_dir/filename``. If the destination is
    locked (open in Word), fall back to a counter-suffixed name. Returns the
    path actually written. Never closes the user's Word (global safety rule)."""
    os.makedirs(dest_dir, exist_ok=True)
    base, ext = os.path.splitext(filename)
    candidates = [filename] + [f"{base} ({i}){ext}" for i in range(2, 21)]
    last_error: Exception | None = None
    for candidate in candidates:
        dest = os.path.join(dest_dir, candidate)
        try:
            generator.assemble_document(caption_path, dest)
            return dest
        except PermissionError as exc:  # destination locked — try next name
            last_error = exc
            continue
    raise last_error or PermissionError("Could not save the mediation brief.")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k save_brief -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/mediation_brief_page.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): lock-safe save helper for mediation brief"
```

---

### Task 5: Implement the generation worker

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` (append `MediationBriefWizardWorker`)
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_mediation_brief_page.py`:

```python
# ---- Worker ----

class _FakeGen:
    """Stand-in for MediationBriefGenerator that records the pipeline calls."""

    def __init__(self):
        self.sections = {}
        self.is_active = False
        self.caption_template_path = None
        self.document_content = ""
        self.calls = []

    def get_style_excerpts(self):
        self.calls.append("style")
        return {}

    def run_planning_pass(self):
        self.calls.append("planning")
        return "plan"

    def generate_section(self, name):
        self.calls.append(f"section:{name}")
        return f"{name} body"

    def assemble_document(self, caption, out):
        self.calls.append("assemble")
        with open(out, "w", encoding="utf-8") as fh:
            fh.write("brief")


def _run_worker(qtbot, monkeypatch, tmp_path, *, cancel_before=False, gen=None):
    from icharlotte_core.ui.wizard.pages import mediation_brief_page as mod
    from icharlotte_core.mediation_brief import GENERATION_ORDER

    fake = gen or _FakeGen()
    monkeypatch.setattr(mod, "MediationBriefGenerator", lambda: fake)
    monkeypatch.setattr(mod, "_read_documents", lambda paths: ("document text", []))

    caption = tmp_path / "Caption.docx"
    caption.write_text("caption", encoding="utf-8")
    case_path = str(tmp_path)
    settings = {"files": ["x.pdf"], "caption_path": str(caption)}

    worker = mod.MediationBriefWizardWorker(case_path, "001", settings)
    progress = []
    results = []
    worker.progress.connect(progress.append)
    worker.finished_result.connect(lambda ok, payload: results.append((ok, payload)))
    if cancel_before:
        worker.cancel()
    worker.run()  # run synchronously in the test thread
    return fake, progress, results, GENERATION_ORDER


def test_worker_success_runs_full_pipeline(qtbot, monkeypatch, tmp_path):
    fake, progress, results, order = _run_worker(qtbot, monkeypatch, tmp_path)

    assert fake.calls[0] == "style"
    assert fake.calls[1] == "planning"
    for name in order:
        assert f"section:{name}" in fake.calls
    assert fake.is_active is True
    assert len(results) == 1
    ok, path = results[0]
    assert ok is True
    assert os.path.join("NOTES", "AI OUTPUT", "MEDIATION") in path
    assert os.path.isfile(path)


def test_worker_cancel_stops_before_sections(qtbot, monkeypatch, tmp_path):
    fake, progress, results, order = _run_worker(
        qtbot, monkeypatch, tmp_path, cancel_before=True
    )

    assert results == [(False, "Generation cancelled.")]
    assert not any(c.startswith("section:") for c in fake.calls)


def test_worker_reports_engine_error(qtbot, monkeypatch, tmp_path):
    class BoomGen(_FakeGen):
        def run_planning_pass(self):
            raise RuntimeError("boom")

    fake, progress, results, order = _run_worker(
        qtbot, monkeypatch, tmp_path, gen=BoomGen()
    )

    assert len(results) == 1
    ok, payload = results[0]
    assert ok is False
    assert "boom" in payload


def test_worker_fails_on_empty_content(qtbot, monkeypatch, tmp_path):
    from icharlotte_core.ui.wizard.pages import mediation_brief_page as mod

    monkeypatch.setattr(mod, "MediationBriefGenerator", lambda: _FakeGen())
    monkeypatch.setattr(mod, "_read_documents", lambda paths: ("", ["nothing"]))
    caption = tmp_path / "Caption.docx"
    caption.write_text("caption", encoding="utf-8")

    worker = mod.MediationBriefWizardWorker(
        str(tmp_path), "001",
        {"files": ["x.pdf"], "caption_path": str(caption)},
    )
    results = []
    worker.finished_result.connect(lambda ok, payload: results.append((ok, payload)))
    worker.run()

    assert results and results[0][0] is False
    assert "could not read" in results[0][1].lower()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k worker -v`
Expected: FAIL (`AttributeError: ... MediationBriefWizardWorker`).

- [ ] **Step 3: Implement the worker**

Append to `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` (after `_save_brief`):

```python
# -------------------------------- worker ----------------------------------

class MediationBriefWizardWorker(QThread):
    progress = Signal(str)
    finished_result = Signal(bool, object)  # (success, output_path | error_msg)

    def __init__(self, case_path: str, file_number: str, settings: dict, parent=None):
        super().__init__(parent)
        self.case_path = case_path or ""
        self.file_number = file_number or ""
        self.settings = dict(settings or {})
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def _emit_cancelled(self) -> bool:
        if self._cancelled:
            self.finished_result.emit(False, "Generation cancelled.")
            return True
        return False

    def run(self) -> None:
        try:
            files = list(self.settings.get("files", []))
            caption_path = self.settings.get("caption_path", "")

            self.progress.emit("Reading source documents…")
            content, warnings = _read_documents(files)
            for warning in warnings:
                self.progress.emit(f"WARNING: {warning}")
            if not content.strip():
                self.finished_result.emit(
                    False, "Could not read any text from the selected documents."
                )
                return
            if not caption_path or not os.path.isfile(caption_path):
                self.finished_result.emit(
                    False, "No caption template selected (or it no longer exists)."
                )
                return
            if self._emit_cancelled():
                return

            generator = MediationBriefGenerator()
            generator.caption_template_path = caption_path
            generator.document_content = content

            self.progress.emit("Loading style reference…")
            generator.get_style_excerpts()

            if self._emit_cancelled():
                return
            self.progress.emit("Analyzing documents…")
            generator.run_planning_pass()

            total = len(GENERATION_ORDER)
            for index, section_name in enumerate(GENERATION_ORDER, start=1):
                if self._emit_cancelled():
                    return
                _, title = SECTION_HEADINGS.get(section_name, ("", section_name))
                label = title or section_name
                self.progress.emit(f"Generating {label} ({index} of {total})…")
                result = generator.generate_section(section_name)
                if result:
                    generator.sections[section_name] = result
            generator.is_active = True

            if self._emit_cancelled():
                return
            self.progress.emit("Assembling document…")
            dest_dir = os.path.join(self.case_path, *MEDIATION_SUBDIR)
            filename = self.settings.get("suggested_filename") or DEFAULT_BRIEF_FILENAME
            saved_path = _save_brief(generator, caption_path, dest_dir, filename)
            self.progress.emit(f"Saved to {saved_path}")
            self.finished_result.emit(True, saved_path)
        except Exception as exc:  # noqa: BLE001
            logger.exception("MediationBriefWizardWorker failed")
            self.finished_result.emit(False, str(exc))
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k worker -v`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/mediation_brief_page.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): mediation brief generation worker"
```

---

### Task 6: Implement the settings page

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` (append `MediationBriefSettingsPage`)
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_mediation_brief_page.py`:

```python
# ---- Settings page ----

def test_settings_page_autodetects_caption(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import (
        MediationBriefSettingsPage,
    )

    (tmp_path / "Caption Page.docx").write_text("x", encoding="utf-8")
    page = MediationBriefSettingsPage(str(tmp_path), "001")
    qtbot.addWidget(page)

    assert page.caption_edit.text().endswith("Caption Page.docx")


def test_settings_page_add_dedupes_and_roundtrips(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import (
        MediationBriefSettingsPage,
    )

    page = MediationBriefSettingsPage(str(tmp_path), "001")
    qtbot.addWidget(page)

    page.files_list.addItem("a.pdf")
    page.files_list.addItem("a.pdf")  # duplicate added manually
    page.from_dict({"files": ["b.pdf", "c.pdf"], "caption_path": "cap.docx"})

    assert page.current_files() == ["b.pdf", "c.pdf"]
    assert page.to_dict() == {"files": ["b.pdf", "c.pdf"], "caption_path": "cap.docx"}


def test_settings_page_validation_blocks_without_files(qtbot, tmp_path, monkeypatch):
    from icharlotte_core.ui.wizard.pages import mediation_brief_page as mod

    page = mod.MediationBriefSettingsPage(str(tmp_path), "001")
    qtbot.addWidget(page)
    monkeypatch.setattr(mod.QMessageBox, "warning", lambda *a, **k: None)

    emitted = []
    page.run_requested.connect(emitted.append)
    page._on_generate()  # no files, no caption

    assert emitted == []


def test_settings_page_emits_run_requested(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import (
        MediationBriefSettingsPage,
    )

    caption = tmp_path / "Caption.docx"
    caption.write_text("x", encoding="utf-8")
    page = MediationBriefSettingsPage(str(tmp_path), "001")
    qtbot.addWidget(page)
    page.files_list.addItem("doc1.pdf")
    page.caption_edit.setText(str(caption))

    with qtbot.waitSignal(page.run_requested, timeout=500) as blocker:
        page._on_generate()

    payload = blocker.args[0]
    assert payload["files"] == ["doc1.pdf"]
    assert payload["caption_path"] == str(caption)
    assert payload["suggested_filename"].endswith(".docx")
    assert payload["save_default_dir"].endswith(os.path.join("AI OUTPUT", "MEDIATION"))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k settings_page -v`
Expected: FAIL (`ImportError: cannot import name 'MediationBriefSettingsPage'`).

- [ ] **Step 3: Implement the settings page**

Append to `icharlotte_core/ui/wizard/pages/mediation_brief_page.py`:

```python
# ----------------------------- settings page ------------------------------

class MediationBriefSettingsPage(QWidget):
    """Source documents + caption template + Generate."""

    run_requested = Signal(dict)

    def __init__(self, case_path: str, file_number: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.case_path = case_path or ""
        self.file_number = file_number or ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.helper_text(
            "Generate a defense-side mediation brief from the selected case "
            "documents. The brief is built section by section and saved to "
            "NOTES/AI OUTPUT/MEDIATION. Refine sections or add deposition quotes "
            "afterward in Word (Win+V → Mediation Brief)."
        ))

        files_row = QHBoxLayout()
        files_row.addWidget(theme.section_header("Source documents"))
        files_row.addStretch()
        self.add_files_btn = theme.secondary_button("Add Files…")
        self.add_files_btn.clicked.connect(self._on_add_files)
        files_row.addWidget(self.add_files_btn)
        self.remove_files_btn = theme.secondary_button("Remove")
        self.remove_files_btn.clicked.connect(self._on_remove_files)
        files_row.addWidget(self.remove_files_btn)
        layout.addLayout(files_row)

        self.files_list = QListWidget()
        self.files_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        layout.addWidget(self.files_list, 1)

        layout.addWidget(theme.section_header("Caption template"))
        caption_row = QHBoxLayout()
        self.caption_edit = QLineEdit()
        self.caption_edit.setPlaceholderText("Path to the case caption .docx…")
        self.caption_edit.textChanged.connect(self._update_caption_hint)
        caption_row.addWidget(self.caption_edit, 1)
        self.browse_caption_btn = theme.secondary_button("Browse…")
        self.browse_caption_btn.clicked.connect(self._on_browse_caption)
        caption_row.addWidget(self.browse_caption_btn)
        layout.addLayout(caption_row)
        self.caption_hint = theme.error_text("")
        layout.addWidget(self.caption_hint)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.generate_btn = theme.primary_button("Generate Mediation Brief")
        self.generate_btn.clicked.connect(self._on_generate)
        btn_row.addWidget(self.generate_btn)
        layout.addLayout(btn_row)

        self._autodetect_caption()

    # ---- helpers ----

    def _autodetect_caption(self) -> None:
        if self.case_path:
            try:
                found = MediationBriefGenerator().find_caption_template(self.case_path)
            except Exception:  # noqa: BLE001
                found = None
            if found:
                self.caption_edit.setText(found)
        self._update_caption_hint()

    def _update_caption_hint(self, *_args) -> None:
        path = self.caption_edit.text().strip()
        if not path:
            self.caption_hint.setText(
                "No caption template found — click Browse… to select one."
            )
        elif not os.path.isfile(path):
            self.caption_hint.setText("Caption template not found at that path.")
        else:
            self.caption_hint.setText("")

    def current_files(self) -> list[str]:
        return [self.files_list.item(i).text() for i in range(self.files_list.count())]

    def _on_add_files(self) -> None:
        picked = ContextFilesDialog.get_files(
            self,
            title="Select source document(s) for the mediation brief",
            start_dir=self.case_path or "",
            file_filter="Documents (*.pdf *.docx *.doc *.txt *.msg);;All files (*.*)",
        )
        existing = set(self.current_files())
        for path in picked or []:
            if path not in existing:
                self.files_list.addItem(path)
                existing.add(path)

    def _on_remove_files(self) -> None:
        for item in self.files_list.selectedItems():
            self.files_list.takeItem(self.files_list.row(item))

    def _on_browse_caption(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select caption template", self.case_path or "",
            "Word Documents (*.docx)",
        )
        if path:
            self.caption_edit.setText(path)

    def to_dict(self) -> dict:
        return {
            "files": self.current_files(),
            "caption_path": self.caption_edit.text().strip(),
        }

    def from_dict(self, data: dict | None) -> None:
        data = data or {}
        self.files_list.clear()
        for path in data.get("files", []) or []:
            self.files_list.addItem(path)
        caption = data.get("caption_path", "")
        if caption:
            self.caption_edit.setText(caption)
        self._update_caption_hint()

    def _on_generate(self) -> None:
        files = self.current_files()
        caption_path = self.caption_edit.text().strip()
        if not files:
            QMessageBox.warning(self, "No documents", "Add at least one source document.")
            return
        if not caption_path or not os.path.isfile(caption_path):
            QMessageBox.warning(
                self, "No caption template", "Select a caption template (.docx)."
            )
            return
        self.run_requested.emit({
            "files": files,
            "caption_path": caption_path,
            "save_default_dir": os.path.join(self.case_path, *MEDIATION_SUBDIR),
            "suggested_filename": DEFAULT_BRIEF_FILENAME,
        })
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k settings_page -v`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/mediation_brief_page.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): mediation brief settings page"
```

---

### Task 7: Implement the output page

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` (append `MediationBriefOutputPage`)
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_mediation_brief_page.py`:

```python
# ---- Output page ----

def _make_real_docx(path):
    from docx import Document

    doc = Document()
    doc.add_paragraph("Brief preview body")
    doc.save(str(path))


def test_output_page_show_result_enables_actions(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import (
        MediationBriefOutputPage,
    )

    out = tmp_path / "Brief.docx"
    _make_real_docx(out)
    page = MediationBriefOutputPage()
    qtbot.addWidget(page)
    page.show_result(str(out))

    assert page.output_path == str(out)
    assert page.open_btn.isEnabled() is True
    assert page.save_as_btn.isEnabled() is True
    assert "Saved to" in page.saved_banner.text()


def test_output_page_save_a_copy(qtbot, tmp_path, monkeypatch):
    from icharlotte_core.ui.wizard.pages import mediation_brief_page as mod

    out = tmp_path / "Brief.docx"
    _make_real_docx(out)
    target = tmp_path / "copy" / "MyBrief.docx"
    target.parent.mkdir()

    page = mod.MediationBriefOutputPage()
    qtbot.addWidget(page)
    page.show_result(str(out))

    monkeypatch.setattr(
        mod.QFileDialog, "getSaveFileName",
        staticmethod(lambda *a, **k: (str(target), "")),
    )
    monkeypatch.setattr(mod.QMessageBox, "information", lambda *a, **k: None)
    page._on_save_as()

    assert target.is_file()
    # Original stays in place.
    assert out.is_file()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k output_page -v`
Expected: FAIL (`ImportError: cannot import name 'MediationBriefOutputPage'`).

- [ ] **Step 3: Implement the output page**

Append to `icharlotte_core/ui/wizard/pages/mediation_brief_page.py`:

```python
# ------------------------------ output page -------------------------------

class MediationBriefOutputPage(QWidget):
    """Read-only preview + Open in Word + Save a Copy As…."""

    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._output_path = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        layout.setSpacing(theme.SPACE_MD)

        self.saved_banner = QLabel("")
        self.saved_banner.setWordWrap(True)
        self.saved_banner.setStyleSheet(
            f"background-color: {theme.SUCCESS_BG}; color: {theme.SUCCESS};"
            f" border-radius: {theme.RADIUS_SM}px; padding: 8px 12px;"
            f" font-weight: 600;"
        )
        self.saved_banner.setVisible(False)
        layout.addWidget(self.saved_banner)

        self.preview = QTextBrowser()
        layout.addWidget(self.preview, 1)

        btn_row = QHBoxLayout()
        self.open_btn = theme.secondary_button("Open in Word")
        self.open_btn.clicked.connect(self._on_open)
        self.open_btn.setEnabled(False)
        btn_row.addWidget(self.open_btn)
        self.save_as_btn = theme.secondary_button("Save a Copy As…")
        self.save_as_btn.clicked.connect(self._on_save_as)
        self.save_as_btn.setEnabled(False)
        btn_row.addWidget(self.save_as_btn)
        btn_row.addStretch()
        self.rerun_btn = theme.secondary_button("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_btn = theme.secondary_button("Edit Settings & Re-run")
        self.edit_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_btn)
        layout.addLayout(btn_row)

    @property
    def output_path(self) -> str:
        return self._output_path

    def show_result(self, path: str) -> None:
        self._output_path = path or ""
        ok = bool(self._output_path) and os.path.isfile(self._output_path)
        self.open_btn.setEnabled(ok)
        self.save_as_btn.setEnabled(ok)
        if ok:
            self.saved_banner.setText(f"Saved to: {self._output_path}")
            self.saved_banner.setVisible(True)
            try:
                self.preview.setHtml(load_docx_as_html(self._output_path))
            except Exception as exc:  # noqa: BLE001
                self.preview.setPlainText(f"(Could not render preview: {exc})")
        else:
            self.saved_banner.setVisible(False)
            self.preview.setPlainText("(No output produced.)")

    def load_output(self, path: str) -> None:
        self.show_result(path)

    def _on_open(self) -> None:
        if self._output_path and os.path.isfile(self._output_path):
            try:
                os.startfile(self._output_path)  # Windows
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "Open failed", f"Could not open in Word:\n{exc}")

    def _on_save_as(self) -> None:
        if not self._output_path or not os.path.isfile(self._output_path):
            return
        default = os.path.join(
            os.path.dirname(self._output_path), os.path.basename(self._output_path)
        )
        target, _ = QFileDialog.getSaveFileName(
            self, "Save a copy of the mediation brief", default,
            "Word Documents (*.docx);;All files (*.*)",
        )
        if not target:
            return
        if not target.lower().endswith(".docx"):
            target += ".docx"
        try:
            shutil.copyfile(self._output_path, target)
            QMessageBox.information(self, "Saved", f"Saved a copy to:\n{target}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Save failed", f"Could not save a copy:\n{exc}")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k output_page -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/mediation_brief_page.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): mediation brief output page"
```

---

### Task 8: Implement the task tab and builder

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/mediation_brief_page.py` (append `MediationBriefTaskTab` + `build_mediation_brief_tab`)
- Test: `tests/test_wizard/test_mediation_brief_page.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_wizard/test_mediation_brief_page.py`:

```python
# ---- Task tab ----

def test_build_tab_returns_tab_with_three_pages(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import (
        MediationBriefTaskTab,
        build_mediation_brief_tab,
    )
    from icharlotte_core.ui.wizard.registry import get_task

    spec = get_task("mediation_brief")
    tab = build_mediation_brief_tab(spec, str(tmp_path), "001", None)
    qtbot.addWidget(tab)

    assert isinstance(tab, MediationBriefTaskTab)
    assert tab.count() == 3
    assert tab.spec.task_id == "mediation_brief"
    assert tab.files == []


def test_tab_run_switches_to_status_and_starts_worker(qtbot, tmp_path, monkeypatch):
    from icharlotte_core.ui.wizard.pages import mediation_brief_page as mod
    from icharlotte_core.ui.wizard.registry import get_task

    spec = get_task("mediation_brief")
    tab = mod.MediationBriefTaskTab(spec, str(tmp_path), "001", None)
    qtbot.addWidget(tab)

    started = {}

    class _StubWorker:
        def __init__(self, *a, **k):
            started["created"] = True
        progress = None
        finished_result = None
        finished = None

        # Minimal QThread-like surface used by the tab.
        def start(self):
            started["started"] = True

    # Patch the worker so no real generation runs; just confirm wiring.
    real = mod.MediationBriefWizardWorker

    class _Recording(real):
        def start(self):  # do not actually launch the thread
            started["started"] = True

    monkeypatch.setattr(mod, "MediationBriefWizardWorker", _Recording)

    tab._on_run({"files": ["x.pdf"], "caption_path": str(tmp_path / "c.docx")})

    assert tab.currentIndex() == mod.TASK_PAGE_STATUS
    assert started.get("started") is True
    # Clean up the created (but never started) worker reference.
    tab._worker = None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k "build_tab or tab_run" -v`
Expected: FAIL (`ImportError: cannot import name 'MediationBriefTaskTab'`).

- [ ] **Step 3: Implement the tab and builder**

Append to `icharlotte_core/ui/wizard/pages/mediation_brief_page.py`:

```python
# --------------------------------- tab ------------------------------------

class MediationBriefTaskTab(WizardTaskContainer):
    """Settings → Status → Output container for the Mediation Brief task."""

    task_completed = Signal(dict)

    def __init__(self, spec, case_path: str, file_number: str, parent: QWidget | None = None):
        super().__init__(spec, parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._worker = None
        self._finishing_worker = None
        self._last_settings: dict = {}

        self.settings_page = MediationBriefSettingsPage(case_path, file_number)
        self.status_page = StatusPage()
        self.output_page = MediationBriefOutputPage()

        self.addWidget(self.settings_page)   # TASK_PAGE_SETTINGS
        self.addWidget(self.status_page)     # TASK_PAGE_STATUS
        self.addWidget(self.output_page)     # TASK_PAGE_OUTPUT

        self.settings_page.run_requested.connect(self._on_run)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.output_page.rerun_requested.connect(self._on_rerun)
        self.output_page.edit_settings_requested.connect(
            lambda: self.setCurrentIndex(TASK_PAGE_SETTINGS)
        )

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list[str]:
        return list(self.settings_page.current_files())

    def _on_run(self, settings: dict) -> None:
        if self._worker is not None or self._finishing_worker is not None:
            self.status_page.on_status("A mediation brief is already generating.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        self._last_settings = dict(settings or {})
        self.status_page.reset()
        self.status_page.on_status("Starting mediation brief…")
        self.status_page.progress_bar.setRange(0, 0)  # indeterminate
        self.setCurrentIndex(TASK_PAGE_STATUS)

        worker = MediationBriefWizardWorker(
            self._case_path, self._file_number, self._last_settings, parent=None
        )
        worker.progress.connect(self.status_page.on_status)
        worker.finished_result.connect(self._on_worker_finished)
        worker.finished.connect(lambda w=worker: self._on_worker_thread_finished(w))
        worker.finished.connect(worker.deleteLater)
        self._worker = worker
        worker.start()

    def _on_cancel(self) -> None:
        if self._worker is not None:
            self.status_page.on_status("Cancelling… (finishing the current section)")
            try:
                self._worker.cancel()
            except Exception:  # noqa: BLE001
                pass
        else:
            self.setCurrentIndex(TASK_PAGE_SETTINGS)

    def _on_rerun(self) -> None:
        self._on_run(self._last_settings)

    def _on_worker_finished(self, success: bool, payload: object) -> None:
        from datetime import datetime

        if self.sender() is not None and self.sender() is not self._worker:
            return
        if self.sender() is self._worker:
            self._finishing_worker = self._worker
            self._worker = None

        if not success:
            self.status_page.on_status(f"FAILED: {payload}")
            # Re-enable the status button so the user can go back (routes through
            # cancel_requested → _on_cancel, which returns to Settings when no
            # worker is running). No disconnect/reconnect — keeps cancel wiring
            # intact across re-runs.
            self.status_page.cancel_btn.setEnabled(True)
            self.status_page.cancel_btn.setText("Back to Settings")
            return

        path = payload if isinstance(payload, str) else ""
        self.output_page.show_result(path)
        self.setCurrentIndex(TASK_PAGE_OUTPUT)
        self.task_completed.emit({
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": list(self.settings_page.current_files()),
            "settings": self.settings_page.to_dict(),
            "output_path": path,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        })

    def _on_worker_thread_finished(self, worker) -> None:
        if self._worker is worker:
            self._worker = None
        if self._finishing_worker is worker:
            self._finishing_worker = None

    def closeEvent(self, event) -> None:
        for worker in (self._worker, self._finishing_worker):
            if worker is not None and worker.isRunning():
                QMessageBox.information(
                    self, "Task running",
                    "The mediation brief is still generating. Wait for it to "
                    "finish before closing this tab.",
                )
                event.ignore()
                return
        super().closeEvent(event)


def build_mediation_brief_tab(spec, case_path: str, file_number: str, parent: QWidget | None):
    """Open the Mediation Brief task on its Settings page."""
    return MediationBriefTaskTab(
        spec=spec, case_path=case_path, file_number=file_number, parent=parent
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -k "build_tab or tab_run" -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Run the whole new test file**

Run: `python -m pytest tests/test_wizard/test_mediation_brief_page.py -v`
Expected: PASS (all tests, ~18).

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/ui/wizard/pages/mediation_brief_page.py tests/test_wizard/test_mediation_brief_page.py
git commit -m "feat(wizard): mediation brief task tab + builder"
```

---

### Task 9: Wire the close-guard in iCharlotte.py

**Files:**
- Modify: `iCharlotte.py:1719-1721` (the class-name tuple in `_on_tab_close_requested`)

- [ ] **Step 1: Read the current close guard**

Run: `python -m pytest -q` is not applicable here; instead open `iCharlotte.py` and locate `_on_tab_close_requested`. The current tuple reads:

```python
            if widget.__class__.__name__ in (
                "OpposeMotionTaskTab", "GenerateMotionTaskTab", "SeparateTaskTab"
            ) and worker.isRunning():
```

- [ ] **Step 2: Add `MediationBriefTaskTab` to the tuple**

Replace it with:

```python
            if widget.__class__.__name__ in (
                "OpposeMotionTaskTab", "GenerateMotionTaskTab", "SeparateTaskTab",
                "MediationBriefTaskTab",
            ) and worker.isRunning():
```

This makes closing a tab mid-generation show "wait for it to finish" instead of tearing down the running thread. (The handler uses `removeTab` + `deleteLater`, not `tab.close()`, so the tab's own `closeEvent` does not fire on the X click — this tuple is the real protection. The worker is parented to `None`, so even an unguarded close would not orphan it, but the message is the correct UX.)

- [ ] **Step 3: Byte-compile to confirm no syntax error**

Run: `python -c "import ast; ast.parse(open('iCharlotte.py', encoding='utf-8').read()); print('OK')"`
Expected: `OK`

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py
git commit -m "feat(wizard): guard close-during-run for Mediation Brief tab"
```

---

### Task 10: Full suite + manual end-to-end verification

**Files:** none (verification only)

- [ ] **Step 1: Run the full wizard test suite (app must be closed)**

Run: `python -m pytest tests/test_wizard/ -q`
Expected: all pass (the new file added; existing tests unaffected — `test_wizard_tab.py` and `test_task_categories.py` are dynamic against `list_tasks()`).

- [ ] **Step 2: Run the mediation-brief engine tests (regression — engine unchanged)**

Run: `python -m pytest tests/test_mediation_brief_live.py tests/test_mediation_brief_quote_search.py -q`
Expected: all pass (no engine changes were made).

- [ ] **Step 3: Launch the app and verify the card**

Run: `python iCharlotte.py`
Verify: load a case → Wizard mode → the **Mediation Brief** card (🤝) appears under **Motions & Drafting**; typing "mediation" in the launcher search filters to it.

- [ ] **Step 4: Run a generation end-to-end**

In the app: open the Mediation Brief card → add a few real source documents (at least one deposition transcript PDF + pleadings) → confirm the caption template auto-detected (or Browse to one) → click **Generate Mediation Brief**. Watch the Status page show per-section progress. On completion, confirm:
- The Output page renders a preview and shows **"Saved to: …\NOTES\AI OUTPUT\MEDIATION\Defendant's Confidential Mediation Brief.docx"**.
- The file exists at that path on disk.
- **Open in Word** opens it; **Save a Copy As…** writes a copy to a chosen location while the original remains in MEDIATION.

- [ ] **Step 5: Verify the Word document is valid**

The worker's `assemble_document` already runs `validate_report` (printed to the log). Confirm the log shows the validation summary with no errors. If errors appear, debug before considering the task complete (mandatory Word-validation rule).

- [ ] **Step 6: Verify the Chat-tab feature is untouched**

In Advanced mode → Chat tab → Templates dropdown still shows **Mediation Brief**, and it still generates as before (engine shared, unchanged).

- [ ] **Step 7: Final commit (if any verification fixes were needed)**

```bash
git add -A
git commit -m "test(wizard): verify Mediation Brief task end-to-end"
```

---

## Self-Review

**Spec coverage:**
- Card under "Motions & Drafting" → Task 1. ✓
- Routing as in-process builder, no pre-Settings picker → Task 2. ✓
- Settings page: multi-folder picker + caption auto-detect/Browse + Generate + validation + to/from_dict → Task 6. ✓
- Worker drives the unmodified generator (read → style → planning → sections → assemble), with progress + cancel + error → Task 5; table-aware extraction → Task 3. ✓
- Default save to `NOTES/AI OUTPUT/MEDIATION` with lock fallback → Task 4 + Task 5. ✓
- Output page: read-only preview, "Saved to" banner, Open in Word, Save a Copy As…, Re-run/Edit → Task 7. ✓
- Bespoke tab mirroring GenerateMotionTaskTab; worker parented to None; close guard → Task 8 + Task 9. ✓
- `output_path` + `load_output` for snapshot/reopen → Task 7 (output page) / Task 8 (tab exposes `output_page`). ✓
- Open/reopen/restore via the existing generic in-process branch (fresh tab) → no code needed; verified in Task 10. ✓
- Tests (registry/routing pure-logic + settings/worker/output/tab pytest-qt) → Tasks 1–8. ✓
- Chat tab + engine untouched → verified Task 10 Step 6. ✓

**Placeholder scan:** No "TBD"/"add error handling"/"similar to" — every code step shows complete code.

**Type consistency:** `run_requested(dict)` payload keys (`files`, `caption_path`, `save_default_dir`, `suggested_filename`) are produced in Task 6 and consumed in Task 5. `MediationBriefWizardWorker(case_path, file_number, settings, parent)` signature matches its construction in Task 8. `show_result`/`load_output`/`output_path` defined in Task 7 and used in Task 8. `MEDIATION_SUBDIR`/`DEFAULT_BRIEF_FILENAME` defined in Task 3 and used in Tasks 5–6. `cancel()` defined in Task 5 and called in Task 8. Page-index constants `TASK_PAGE_*` defined in Task 3 and used in Tasks 5–8. Consistent.
