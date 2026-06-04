# Generate Motion — Citation Review Parity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give the Draft-a-Motion (Generate Motion) wizard task the same citation-review surface Oppose-a-Motion has — verdict summary banner, color-coded clickable citation underlines, and a right-side detail panel — and bring its research to the same granularity.

**Architecture:** Extract the side-agnostic citation-review UI out of `oppose_motion_page.py` into a new `citation_review.py` (render/panel helpers, `CitationDetailPanel`, `CitationDetailDialog`, plus a `CitationReviewOutputPage` base class). Refactor `OpposeMotionOutputPage` into a thin subclass (re-exporting the moved symbols so existing tests/imports keep working) and build `GenerateMotionOutputPage` as a parallel subclass. Then upgrade `GenerateMotionWorker` to research each selected outline subsection (reusing oppose's `_research_targets`) with a cache dir and `max_workers=2`.

**Tech Stack:** Python 3, PySide6 (QWidget/QTextBrowser/QLabel), pytest + pytest-qt. Reuses `icharlotte_core.opposition` models/verifier/research and `icharlotte_core.motion_generation`.

**Environment notes:**
- Run tests with the venv interpreter: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest ...` (system Python lacks `bs4`/`PySide6`).
- **Stop the running iCharlotte before a full pytest collection** — a launched app instance breaks PySide6 import in collection (`ImportError: cannot import name 'QApplication'`).
- This is a concurrent multi-session checkout: `git add` only the files listed per task; never `git add -A`.
- The app runs from the main checkout `C:\geminiterminal2\`; restart iCharlotte to see edits live.
- Work happens on branch `feature/generate-motion-citation-review` (already created).

---

## File Structure

| File | Responsibility | Change |
| --- | --- | --- |
| `icharlotte_core/ui/wizard/pages/citation_review.py` | Side-agnostic citation-review UI: render helpers, panel helpers, `CitationDetailPanel`, `CitationDetailDialog`, and `CitationReviewOutputPage` base class. | **Create** |
| `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` | Opposition task page. `OpposeMotionOutputPage` becomes a thin subclass; re-exports moved symbols. `_make_local_corpus`/`_research_targets`/`_ReverifyWorker` stay. | **Modify** |
| `icharlotte_core/ui/wizard/pages/generate_motion_page.py` | Generate-motion task page. `GenerateMotionOutputPage` becomes a `CitationReviewOutputPage` subclass; worker researches subsection leaves. | **Modify** |
| `tests/test_wizard/test_citation_review.py` | Smoke test: shared module exposes symbols; oppose re-exports them. | **Create** |
| `tests/test_wizard/test_generate_motion_output_page.py` | Verdict colors, summary banner, panel-on-click, save-warns-on-red, empty-state for generate output page. | **Create** |
| `tests/test_wizard/test_generate_motion_worker.py` | Worker researches subsection leaves (not just top-level grounds). | **Create** |
| `.gitignore` | Ignore prompt-dir research caches. | **Modify** |

---

## Task 1: Extract `citation_review.py` and refactor `OpposeMotionOutputPage`

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/citation_review.py`
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
- Test: `tests/test_wizard/test_citation_review.py`

- [ ] **Step 1: Write the failing smoke test**

Create `tests/test_wizard/test_citation_review.py`:

```python
"""The shared citation-review toolkit exposes its symbols, and oppose
re-exports them for backward compatibility with existing imports/tests."""
from __future__ import annotations

import pytest

pytest.importorskip("PySide6")


def test_citation_review_exposes_symbols():
    from icharlotte_core.ui.wizard.pages import citation_review as cr

    for name in (
        "CitationReviewOutputPage",
        "CitationDetailPanel",
        "CitationDetailDialog",
        "_render_draft_html",
        "_build_citation_index",
        "_format_inline_html",
        "_color_for_verdict",
        "_citation_header_html",
        "_citation_body_html",
        "_run_find_replacement",
        "_VERDICT_COLORS",
        "_VERDICT_HEADER_COLORS",
        "_VERDICT_LABELS",
    ):
        assert hasattr(cr, name), f"citation_review missing {name}"


def test_oppose_motion_page_reexports_shared_symbols():
    # Existing tests import these names from oppose_motion_page; they must
    # keep resolving after the extraction.
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import (  # noqa: F401
        CitationDetailDialog,
        CitationDetailPanel,
        OpposeMotionOutputPage,
        _render_draft_html,
    )

    from icharlotte_core.ui.wizard.pages.citation_review import (
        CitationReviewOutputPage,
    )

    assert issubclass(OpposeMotionOutputPage, CitationReviewOutputPage)
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_citation_review.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'icharlotte_core.ui.wizard.pages.citation_review'`.

- [ ] **Step 3: Create `citation_review.py`**

Create `icharlotte_core/ui/wizard/pages/citation_review.py`. It has three parts: (A) module header + imports, (B) the **verbatim move** of the render/panel helpers and the two widgets out of `oppose_motion_page.py`, and (C) the **new** `CitationReviewOutputPage` base class.

**(A) Header + imports** — type exactly:

```python
"""Shared citation-review UI for wizard tasks that draft a brief and verify
its citations (Oppose a Motion, Generate Motion).

Holds the verdict-colored body renderer, the per-citation detail helpers, the
right-side ``CitationDetailPanel`` (and legacy ``CitationDetailDialog``), and a
``CitationReviewOutputPage`` base class. All of it is side-agnostic: it
operates only on ``DraftDocument`` / ``CitationVerification``.
"""
from __future__ import annotations

import html
import os
import re
import shutil

from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.opposition.models import DraftDocument
```

**(B) Verbatim move from `oppose_motion_page.py`.** Copy the following spans **exactly as they are today** into `citation_review.py`, in this order. These are pure relocations — do not change logic:

1. Lines `612-709` — the body-render block: `_HORIZONTAL_RULE_RE`, `_MD_HEADING_RE`, `_MD_ITALIC_RE`, `_VERDICT_COLORS`, `_color_for_verdict`, `_render_draft_html`, `_build_citation_index`, `_format_inline_html`.
2. Lines `712-960` — the panel block: `_VERDICT_HEADER_COLORS`, `_VERDICT_LABELS`, `_citation_header_html`, `_citation_body_html`, `_run_find_replacement`, and the `CitationDetailPanel` class.
3. Lines `963-1019` — the `CitationDetailDialog` class.

**One change after the move:** in the relocated `_render_draft_html`, change the title fallback so it is not opposition-specific:

```python
    title = html.escape(draft.title or "Memorandum")
```

(was `draft.title or "Opposition Memorandum"`).

**(C) Append the new base class** to `citation_review.py`:

```python
class CitationReviewOutputPage(QWidget):
    """Output page for tasks that draft a brief then verify its citations.

    Renders the draft body with color-coded clickable citation anchors, shows
    a verdict summary banner, and drives a right-side ``CitationDetailPanel``
    that updates when an anchor is clicked. Subclasses customize the title,
    the empty-citations message, and the bottom action buttons.
    """

    #: Fallback title when ``draft.title`` is empty; also used for the
    #: suggested save filename and the save dialog caption.
    default_title = "Memorandum"

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.draft = DraftDocument()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        self.summary_banner = QLabel("")
        self.summary_banner.setTextFormat(Qt.TextFormat.RichText)
        self.summary_banner.setWordWrap(True)
        self.summary_banner.setVisible(False)
        outer.addWidget(self.summary_banner)

        layout = QHBoxLayout()
        layout.setSpacing(12)
        self.editor = QTextBrowser()
        self.editor.setOpenLinks(False)
        self.editor.setOpenExternalLinks(False)
        self.editor.anchorClicked.connect(self._on_anchor_clicked)
        layout.addWidget(self.editor, 2)

        self.detail_panel = CitationDetailPanel()
        layout.addWidget(self.detail_panel, 1)
        outer.addLayout(layout, 1)

        row = QHBoxLayout()
        row.addStretch()
        self._build_action_buttons(row)
        outer.addLayout(row)

    # -- overridable seams ------------------------------------------------

    def empty_citations_message(self) -> str:
        return (
            "No citations were detected. If California case-law research "
            "returned no results, the draft was written without case "
            "citations. Review for any factual or statutory support that may "
            "need strengthening."
        )

    def _build_action_buttons(self, row: QHBoxLayout) -> None:
        """Populate the bottom button row. Base adds only a Save button;
        subclasses override and call ``_add_save_button`` where they want it."""
        self._add_save_button(row)

    def _add_save_button(self, row: QHBoxLayout) -> None:
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_as)
        row.addWidget(self.save_btn)

    # -- rendering --------------------------------------------------------

    def show_result(self, draft: DraftDocument) -> None:
        self.draft = draft
        self.editor.setHtml(_render_draft_html(draft))
        self._refresh_summary_banner()
        if draft.citations:
            self.show_citation(0)
        else:
            self.detail_panel.clear(self.empty_citations_message())

    def show_citation(self, index: int) -> None:
        if index < 0 or index >= len(self.draft.citations):
            return
        self.detail_panel.set_citation(self.draft.citations[index])

    def _refresh_summary_banner(self) -> None:
        if not self.draft.citations:
            self.summary_banner.setVisible(False)
            return
        counts: dict[str, int] = {}
        for cv in self.draft.citations:
            verdict = (cv.verdict or "UNVERIFIED").upper()
            counts[verdict] = counts.get(verdict, 0) + 1
        total = sum(counts.values())
        red = counts.get("NOT_SUPPORTED", 0) + counts.get("NOT_FOUND", 0)
        parts = [
            f"<b>Verification:</b> {total} citation(s) checked &mdash; ",
            f"\U0001F7E2 {counts.get('SUPPORTED', 0)} supported, ",
            f"\U0001F7E1 {counts.get('PARTIAL', 0)} partial, ",
            f"\U0001F534 {red} flagged, ",
            f"⚪ {counts.get('UNVERIFIED', 0)} unverified.",
        ]
        warning = ""
        if red > 0:
            warning = (
                f"<br><span style='color:#c5221f;'>⚠ {red} citation(s) don't "
                "support what the brief claims. Review the red-flagged cites "
                "before filing.</span>"
            )
        self.summary_banner.setText("".join(parts) + warning)
        self.summary_banner.setVisible(True)

    def _on_anchor_clicked(self, url: QUrl) -> None:
        scheme = url.scheme()
        if scheme == "citation":
            try:
                index = int(url.path().lstrip("/") or url.host() or "0")
            except (TypeError, ValueError):
                return
            self.show_citation(index)
            return
        QDesktopServices.openUrl(url)

    @property
    def output_path(self) -> str:
        return self.draft.preview_path

    def load_output(self, output_path: str) -> None:
        body_text = ""
        if output_path and output_path.lower().endswith(".docx") and os.path.isfile(output_path):
            try:
                from docx import Document

                doc = Document(output_path)
                body_text = "\n".join(p.text for p in doc.paragraphs if p.text)
            except Exception:
                body_text = f"(Could not render generated document: {output_path})"
        self.show_result(
            DraftDocument(
                title=os.path.splitext(os.path.basename(output_path or ""))[0] or self.default_title,
                body_text=body_text,
                preview_path=output_path or "",
            )
        )

    @staticmethod
    def default_save_dir(preview_path: str) -> str:
        marker = os.path.join(".icharlotte", "wizard_previews")
        before_marker, _, _after_marker = preview_path.partition(marker)
        if before_marker:
            return os.path.dirname(os.path.normpath(before_marker))
        return os.path.dirname(preview_path)

    def save_as(self) -> None:
        if not self.draft.preview_path:
            QMessageBox.warning(
                self, "No preview", "No generated preview is available."
            )
            return

        red = sum(
            1 for cv in self.draft.citations
            if (cv.verdict or "").upper() in {"NOT_SUPPORTED", "NOT_FOUND"}
        )
        if red > 0:
            choice = QMessageBox.question(
                self,
                "Citations flagged",
                (
                    f"This document has {red} citation(s) flagged as "
                    "NOT_SUPPORTED or NOT_FOUND.\n\nSave anyway?"
                ),
                QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel,
            )
            if choice != QMessageBox.StandardButton.Save:
                return

        suggested = os.path.join(
            self.default_save_dir(self.draft.preview_path),
            f"{self.draft.title or self.default_title}.docx",
        )
        target, _ = QFileDialog.getSaveFileName(
            self,
            f"Save {self.default_title}",
            suggested,
            "Word Documents (*.docx);;All files (*.*)",
        )
        if not target:
            return
        if not target.lower().endswith(".docx"):
            target += ".docx"
        if os.path.abspath(target) == os.path.abspath(self.draft.preview_path):
            QMessageBox.warning(
                self,
                "Choose another location",
                "Select a location outside the internal preview file.",
            )
            return
        try:
            shutil.copyfile(self.draft.preview_path, target)
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", f"Could not save file:\n{exc}")
            return
        QMessageBox.information(self, "Saved", f"Saved:\n{target}")
```

- [ ] **Step 4: Refactor `oppose_motion_page.py` — add re-export imports**

In `oppose_motion_page.py`, immediately after the existing import block (right after the `from icharlotte_core.word_validator import validate_opposition_docx` line near the top), add:

```python
from icharlotte_core.ui.wizard.pages.citation_review import (  # noqa: F401  (re-exported for tests/back-compat)
    CitationDetailDialog,
    CitationDetailPanel,
    CitationReviewOutputPage,
    _build_citation_index,
    _citation_body_html,
    _citation_header_html,
    _color_for_verdict,
    _format_inline_html,
    _render_draft_html,
    _run_find_replacement,
    _VERDICT_COLORS,
    _VERDICT_HEADER_COLORS,
    _VERDICT_LABELS,
)
```

- [ ] **Step 5: Refactor `oppose_motion_page.py` — delete the moved definitions**

Delete these now-relocated spans from `oppose_motion_page.py` (they live in `citation_review.py` now):
- the body-render block (`_HORIZONTAL_RULE_RE` … `_format_inline_html`),
- the panel block (`_VERDICT_HEADER_COLORS` … end of `CitationDetailPanel`),
- the `CitationDetailDialog` class.

(In the version this plan was written against these were lines `612-1019`. Match by symbol name, not line number, since earlier edits shift lines.)

**Keep** `_ReverifyWorker` (the DEV-only worker), `_make_local_corpus`, `_research_targets`, `_corpus_*`, both workers, the settings page, and `build_oppose_motion_tab`.

- [ ] **Step 6: Refactor `oppose_motion_page.py` — slim `OpposeMotionOutputPage` to a subclass**

Replace the entire `class OpposeMotionOutputPage(QWidget):` definition (its `__init__`, `show_result`, `_refresh_summary_banner`, `_on_anchor_clicked`, `output_path`, `load_output`, `default_save_dir`, `show_citation`, `save_as`, and the DEV re-verify handlers) with this subclass. The DEV-only re-verify button + its two handlers stay here (they are opposition-only dev tooling):

```python
class OpposeMotionOutputPage(CitationReviewOutputPage):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    default_title = "Opposition Memorandum"

    def empty_citations_message(self) -> str:
        return (
            "No citations were detected in this opposition. If California "
            "case-law research returned no results, the draft was written "
            "without case citations. Review the brief for any factual or "
            "statutory support that may need strengthening."
        )

    def _build_action_buttons(self, row) -> None:
        self._reverify_worker = None  # DEV-only worker handle
        # ── DEV-ONLY: re-verify button. Remove when no longer needed. ──
        self.reverify_btn = QPushButton("Re-verify Citations (DEV)")
        self.reverify_btn.setToolTip(
            "Temporary developer button. Re-extracts citations from the "
            "current body and re-runs the verifier so you can test "
            "parser / verifier changes without re-running the full "
            "draft pipeline."
        )
        self.reverify_btn.setStyleSheet("QPushButton { color: #b06000; }")
        self.reverify_btn.clicked.connect(self._on_reverify_clicked)
        row.addWidget(self.reverify_btn)
        # ── END DEV-ONLY ──
        self._add_save_button(row)

    # ── DEV-ONLY: re-verify handlers. Remove when no longer needed. ──

    def _on_reverify_clicked(self) -> None:
        if self._reverify_worker is not None:
            return  # already running
        body_text = (self.draft.body_text or "").strip()
        if not body_text:
            QMessageBox.warning(
                self,
                "Nothing to verify",
                "No draft body is loaded. Re-open or re-run the task first.",
            )
            return
        self.reverify_btn.setEnabled(False)
        self.reverify_btn.setText("Re-verifying… (DEV)")
        self.summary_banner.setText(
            "<b>Re-verifying citations…</b> (this runs the verifier on the "
            "current body without re-drafting)"
        )
        self.summary_banner.setVisible(True)
        worker = _ReverifyWorker(body_text=body_text, parent=None)
        worker.finished_result.connect(self._on_reverify_finished)
        worker.finished.connect(worker.deleteLater)
        self._reverify_worker = worker
        worker.start()

    def _on_reverify_finished(self, success: bool, payload: object) -> None:
        self._reverify_worker = None
        self.reverify_btn.setEnabled(True)
        self.reverify_btn.setText("Re-verify Citations (DEV)")
        if not success:
            QMessageBox.critical(
                self,
                "Re-verify failed",
                f"Re-verification failed:\n\n{payload}",
            )
            self._refresh_summary_banner()
            return
        citations = payload if isinstance(payload, list) else []
        self.draft.citations = citations
        self.editor.setHtml(_render_draft_html(self.draft))
        self._refresh_summary_banner()
        if citations:
            self.show_citation(0)
        else:
            self.detail_panel.clear(
                "Re-verify found no citations in the current body text."
            )

    # ── END DEV-ONLY ──
```

- [ ] **Step 7: Remove now-unused imports from `oppose_motion_page.py`**

After the move, these top-of-file imports are no longer referenced in `oppose_motion_page.py` (all their users moved to `citation_review.py`). Remove them:
- `import html`
- `import shutil`
- from the `PySide6.QtWidgets` import list: `QDialog`, `QTextBrowser`
- from `PySide6.QtGui`: the whole `from PySide6.QtGui import QDesktopServices` line
- from `PySide6.QtCore`: drop `QUrl` (keep `QThread`, `Qt`, `Signal`)

**Keep** `QLabel` (still used by `OpposeMotionSettingsPage`) and `QFileDialog` (used by `build_oppose_motion_tab`). The next step's smoke import catches any over-removal.

- [ ] **Step 8: Smoke-import + run the new test and the full oppose suite**

Run:
```
C:\geminiterminal2\.venv\Scripts\python.exe -c "import icharlotte_core.ui.wizard.pages.oppose_motion_page, icharlotte_core.ui.wizard.pages.citation_review; print('import ok')"
C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_citation_review.py tests/test_wizard/test_oppose_motion_output_page_verdicts.py tests/test_wizard/test_oppose_motion_page.py -v
```
Expected: `import ok`, then all tests PASS (new smoke test + every existing oppose test — proving the extraction + re-exports broke nothing).

- [ ] **Step 9: Commit**

```
git add icharlotte_core/ui/wizard/pages/citation_review.py icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_wizard/test_citation_review.py
git commit -m "refactor(wizard): extract shared citation_review toolkit + base page"
```

---

## Task 2: Rebuild `GenerateMotionOutputPage` as a `CitationReviewOutputPage` subclass

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/generate_motion_page.py`
- Test: `tests/test_wizard/test_generate_motion_output_page.py`

- [ ] **Step 1: Write the failing test file**

Create `tests/test_wizard/test_generate_motion_output_page.py`:

```python
"""Citation-review surface on the Generate Motion output page (parity with
the oppose output page)."""
from __future__ import annotations

import pytest

pytest.importorskip("PySide6")
pytest.importorskip("pytestqt")

from PySide6.QtCore import QUrl  # noqa: E402

from icharlotte_core.opposition.models import (  # noqa: E402
    CitationVerification,
    DraftDocument,
)
from icharlotte_core.ui.wizard.pages.generate_motion_page import (  # noqa: E402
    GenerateMotionOutputPage,
)


def _draft(citations):
    return DraftDocument(
        title="Motion to Compel",
        body_text="See *Smith v. Jones* (2010) 50 Cal.4th 100 for support.",
        citations=citations,
    )


def test_supported_citation_renders_green(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(_draft([
        CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="SUPPORTED", kind="case",
        )
    ]))
    html = page.editor.toHtml().lower()
    assert "#1e8e3e" in html  # green underline color


def test_not_supported_citation_renders_red(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(_draft([
        CitationVerification(
            citation_text="Smith v. Jones (2010) 50 Cal.4th 100",
            verdict="NOT_SUPPORTED", kind="case",
        )
    ]))
    assert "#c5221f" in page.editor.toHtml().lower()  # red


def test_summary_banner_counts_per_verdict(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(DraftDocument(
        title="M", body_text="Body.",
        citations=[
            CitationVerification(citation_text="a", verdict="SUPPORTED"),
            CitationVerification(citation_text="b", verdict="SUPPORTED"),
            CitationVerification(citation_text="c", verdict="PARTIAL"),
            CitationVerification(citation_text="d", verdict="NOT_SUPPORTED"),
        ],
    ))
    banner = page.summary_banner.text().lower()
    assert "supported" in banner
    assert "2" in banner


def test_clicking_anchor_updates_detail_panel(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(DraftDocument(
        title="M",
        body_text="First *A v. B* (2010) 1 Cal.5th 1; then *C v. D* (2011) 2 Cal.5th 2.",
        citations=[
            CitationVerification(citation_text="A v. B (2010) 1 Cal.5th 1",
                                 case_name="A v. B", verdict="SUPPORTED", kind="case"),
            CitationVerification(citation_text="C v. D (2011) 2 Cal.5th 2",
                                 case_name="C v. D", verdict="NOT_SUPPORTED", kind="case"),
        ],
    ))
    page._on_anchor_clicked(QUrl("citation:1"))
    assert "C v. D" in page.detail_panel.header_label.text()
    assert "NOT SUPPORTED" in page.detail_panel.header_label.text()


def test_empty_citations_shows_motion_specific_message(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    page.show_result(DraftDocument(title="M", body_text="Body.", citations=[]))
    assert "motion" in page.detail_panel.body_html.lower()


def test_save_warns_on_red_verdicts(qtbot, monkeypatch, tmp_path):
    from PySide6.QtWidgets import QFileDialog, QMessageBox

    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    preview = tmp_path / "preview.docx"
    preview.write_bytes(b"dummy")
    page.show_result(DraftDocument(
        title="M", body_text="b", preview_path=str(preview),
        citations=[CitationVerification(citation_text="x", verdict="NOT_SUPPORTED")],
    ))

    warned = {"yes": False}

    def fake_question(parent, title, text, *args, **kwargs):
        warned["yes"] = True
        return QMessageBox.StandardButton.Cancel  # user cancels

    monkeypatch.setattr(QMessageBox, "question", fake_question)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **k: ("", ""))

    page.save_as()
    assert warned["yes"]


def test_open_in_word_button_present(qtbot):
    page = GenerateMotionOutputPage()
    qtbot.addWidget(page)
    assert hasattr(page, "open_btn")
    assert page.open_btn.text() == "Open in Word"
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_generate_motion_output_page.py -v`
Expected: FAIL — the current `GenerateMotionOutputPage` has no `editor`/`summary_banner`/`detail_panel`/`save_as`/`_on_anchor_clicked` (it uses `self.body.setPlainText`). Tests error with `AttributeError`.

- [ ] **Step 3: Replace `GenerateMotionOutputPage` in `generate_motion_page.py`**

Add the base-class import alongside the other `icharlotte_core.ui.wizard.pages` imports near the top:

```python
from icharlotte_core.ui.wizard.pages.citation_review import CitationReviewOutputPage
```

Then replace the entire existing `class GenerateMotionOutputPage(QWidget):` definition (the heading/body/open-button version) with:

```python
class GenerateMotionOutputPage(CitationReviewOutputPage):
    default_title = "Generated Motion"

    def empty_citations_message(self) -> str:
        return (
            "No citations were detected in this motion. If California "
            "case-law research returned no results, the motion was drafted "
            "without case citations. Review for any statutory support that "
            "may need strengthening."
        )

    def _build_action_buttons(self, row) -> None:
        self._add_save_button(row)
        self.open_btn = QPushButton("Open in Word")
        self.open_btn.clicked.connect(self._open_preview)
        self.open_btn.setEnabled(False)
        row.addWidget(self.open_btn)

    def show_result(self, draft: DraftDocument) -> None:
        super().show_result(draft)
        self._refresh_open_btn()

    def load_output(self, output_path: str) -> None:
        super().load_output(output_path)
        self._refresh_open_btn()

    def _refresh_open_btn(self) -> None:
        path = self.output_path
        self.open_btn.setEnabled(bool(path) and os.path.isfile(path))

    def _open_preview(self) -> None:
        path = self.output_path
        if path and os.path.isfile(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
```

`QPushButton`, `QDesktopServices`, `QUrl`, and `os` are already imported at the top of `generate_motion_page.py`. The `QTextBrowser` import is now unused — remove `QTextBrowser` from the `PySide6.QtWidgets` import list. (`load_output`'s positional `output_path` and the `output_path` property required by `iCharlotte.py` reopen wiring are both inherited/preserved.)

- [ ] **Step 4: Run the test to verify it passes**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_generate_motion_output_page.py -v`
Expected: PASS — all 7 tests.

- [ ] **Step 5: Run the existing generate settings-page tests (no regression)**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_generate_motion_page.py -v`
Expected: PASS (unchanged — they only touch the settings page).

- [ ] **Step 6: Commit**

```
git add icharlotte_core/ui/wizard/pages/generate_motion_page.py tests/test_wizard/test_generate_motion_output_page.py
git commit -m "feat(generate-motion): citation-review output page (banner, colored cites, detail panel)"
```

---

## Task 3: Worker research-granularity parity + cache + gitignore

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/generate_motion_page.py`
- Modify: `.gitignore`
- Test: `tests/test_wizard/test_generate_motion_worker.py`

- [ ] **Step 1: Write the failing worker test**

Create `tests/test_wizard/test_generate_motion_worker.py`:

```python
"""GenerateMotionWorker researches each selected outline subsection (parity
with oppose), not just the top-level grounds."""
from __future__ import annotations

import pytest

pytest.importorskip("PySide6")


def test_worker_researches_subsection_leaves(monkeypatch, tmp_path):
    import icharlotte_core.motion_generation.samples as samples
    import icharlotte_core.ui.wizard.pages.generate_motion_page as gm
    from icharlotte_core.opposition.models import (
        DraftDocument, MotionMetadata, OutlineNode,
    )

    # Stub everything around the research call so run() is fast + offline.
    monkeypatch.setattr(gm, "extract_context_bundle", lambda files: ("ctx", []))
    monkeypatch.setattr(gm, "_make_local_corpus", lambda: object())  # truthy → research runs, no token needed
    monkeypatch.setattr(gm, "draft_motion",
                        lambda *a, **k: DraftDocument(title="M", body_text="Body."))
    monkeypatch.setattr(gm, "extract_citations", lambda body: [])
    monkeypatch.setattr(gm, "assemble_motion_preview", lambda **k: k.get("output_path"))
    monkeypatch.setattr(gm.DiscoveryAssembler, "find_caption_page", staticmethod(lambda p: ""))
    monkeypatch.setattr(samples, "load_exemplars", lambda tid: [])

    captured = {}

    def fake_research(targets, **kwargs):
        captured["targets"] = list(targets)
        captured["max_workers"] = kwargs.get("max_workers")
        captured["cache_dir"] = kwargs.get("cache_dir")
        return []

    monkeypatch.setattr(gm, "research_arguments", fake_research)

    settings = {
        "motion_type_id": "compel",
        "motion_type_name": "Motion to Compel",
        "target_files": [],
        "metadata": MotionMetadata(
            motion_type="Motion to Compel",
            relief_requested="Compel further responses",
            principal_arguments=["Boilerplate objections are improper"],
        ).to_dict(),
        "outline": [
            OutlineNode(text="Argument", selected=True, children=[
                OutlineNode(text="Responses were evasive and incomplete", selected=True),
            ]).to_dict(),
        ],
    }
    worker = gm.GenerateMotionWorker(case_path=str(tmp_path), file_number="123", settings=settings)

    results = {}
    worker.finished_result.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()  # synchronous (not via QThread.start)

    assert results.get("ok") is True
    # Top-level ground AND the selected subsection leaf were both researched.
    assert "Boilerplate objections are improper" in captured["targets"]
    assert any("evasive" in t.lower() for t in captured["targets"])
    # Parity knobs.
    assert captured["max_workers"] == 2
    assert captured["cache_dir"] and "generate_motion" in captured["cache_dir"]
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_generate_motion_worker.py -v`
Expected: FAIL — currently the worker passes `metadata.principal_arguments` (so the "evasive" leaf assertion fails), `max_workers=4`, and no `cache_dir`.

- [ ] **Step 3: Import `_research_targets` in `generate_motion_page.py`**

Change the existing import line:

```python
from icharlotte_core.ui.wizard.pages.oppose_motion_page import _make_local_corpus
```

to:

```python
from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
    _make_local_corpus,
    _research_targets,
)
```

- [ ] **Step 4: Update the research block in `GenerateMotionWorker.run()`**

In `GenerateMotionWorker.run()`, find the research block (it begins `corpus = _make_local_corpus()` and the `if metadata.principal_arguments and (corpus is not None or token):` condition). Replace that whole block — from `corpus = _make_local_corpus()` down to the closing `else:` warning emit — with:

```python
            research_targets = _research_targets(metadata, plan)
            corpus = _make_local_corpus()
            token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()
            retrieved = []
            # Cache opinion text under this task's prompt dir (mirrors oppose).
            repo_root = os.path.dirname(
                os.path.dirname(
                    os.path.dirname(
                        os.path.dirname(os.path.dirname(__file__))
                    )
                )
            )
            opinion_cache = os.path.join(
                repo_root, "Scripts", "prompts", "generate_motion", ".cache", "opinions"
            )
            if research_targets and (corpus is not None or token):
                client = corpus
                if client is None:
                    from icharlotte_core.legal_research.sources.courtlistener import (
                        CourtListenerClient,
                    )
                    client = CourtListenerClient(token)
                    self.progress.emit(
                        "Local corpus not built; using CourtListener API "
                        f"({len(research_targets)} points)..."
                    )
                else:
                    self.progress.emit(
                        f"Researching authorities locally ({len(research_targets)} points)..."
                    )
                retrieved = research_arguments(
                    research_targets,
                    cl_client=client,
                    query_llm=make_pass_llm("research_queries"),
                    rerank_llm=make_pass_llm("rerank_select"),
                    # Keep concurrency low: 4 parallel workers burst the LLM /
                    # CourtListener rate limit on the per-point query-gen +
                    # rerank calls (oppose's hard-won lesson).
                    max_workers=2,
                    on_progress=self.progress.emit,
                    cache_dir=opinion_cache,
                )
                self.progress.emit(f"Retrieved {len(retrieved)} grounded authorities.")
            else:
                self.progress.emit(
                    "WARNING: no grounded research available; drafting from statutes only."
                )
```

(`plan` is already computed earlier in `run()` as `plan = selected_section_plan(outline)`; `make_pass_llm` is already obtained from `_make_llms()`.)

- [ ] **Step 5: Run the worker test to verify it passes**

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/test_generate_motion_worker.py -v`
Expected: PASS.

- [ ] **Step 6: Add the cache dir to `.gitignore`**

In `.gitignore`, under the `# Project Specific` section, add:

```
# Per-task research opinion caches
Scripts/prompts/*/.cache/
```

- [ ] **Step 7: Commit**

```
git add icharlotte_core/ui/wizard/pages/generate_motion_page.py tests/test_wizard/test_generate_motion_worker.py .gitignore
git commit -m "feat(generate-motion): research each outline subsection w/ cache (worker parity)"
```

---

## Task 4: Full-suite regression + live app verification

**Files:** none (verification only).

- [ ] **Step 1: Run the whole wizard test suite**

Ensure iCharlotte is **not** running (a live app instance breaks PySide6 import in pytest collection), then:

Run: `C:\geminiterminal2\.venv\Scripts\python.exe -m pytest tests/test_wizard/ -q`
Expected: all pass (no regressions in oppose, generate, or other wizard tasks). Investigate and fix any failure before proceeding.

- [ ] **Step 2: Live verification in the running app (MANDATORY per CLAUDE.md)**

Launch iCharlotte (`python iCharlotte.py` from `C:\geminiterminal2\`), open a case, run **Draft a Motion** end-to-end on a real motion type with context documents, and confirm on the output page:
- the verdict **summary banner** appears with per-verdict counts,
- citations in the body are **color-coded and clickable**,
- clicking a citation updates the **right-side detail panel** (with Open-source / Find-replacement where applicable),
- **Save** prompts a red-flag warning when any citation is NOT_SUPPORTED/NOT_FOUND, and **Open in Word** opens the preview.

Capture a screenshot for the record:
`powershell -ExecutionPolicy Bypass -File "C:\geminiterminal2\screenshot_util.ps1" -WindowTitle "iCharlotte"`
then read `screenshot.png`. If anything is off, debug, fix, and re-verify (do not close the user's Word windows).

- [ ] **Step 3: Update memory**

Append a topic-file note (and a one-line `MEMORY.md` index entry) recording that the citation-review UI now lives in `citation_review.py` (`CitationReviewOutputPage` base) shared by both motion tasks, and that the generate worker reached research parity (`_research_targets`, `max_workers=2`, cache dir). Link `[[oppose_motion_redesign]]` and `[[wizard_categories_and_generate_motion]]`.

---

## Self-Review

**Spec coverage:**
- Output-page parity (banner, colored cites, detail panel, save-with-warning) → Task 2. ✓
- Shared toolkit + `CitationReviewOutputPage` base, both pages subclass → Tasks 1 & 2. ✓
- Backward-compat re-exports so oppose tests pass → Task 1 (Steps 4, 8). ✓
- Worker research granularity (`_research_targets`), cache dir, `max_workers` 4→2 → Task 3. ✓
- "Stays in parity" goal → shared base class (Task 1). ✓
- Non-goal: validation already internal to assembler → confirmed, no task. ✓
- Non-goal: no reopen verdict persistence; no DEV button on generate → reflected (generate subclass omits reverify; `load_output` inherited). ✓
- Testing + env caveats → Tasks 1–4 use the venv interpreter and the "stop the app" note. ✓

**Placeholder scan:** No TBD/TODO; every code step has complete code; verbatim-move steps cite exact symbols + line spans (precise, not vague). ✓

**Type/name consistency:** `default_title`, `empty_citations_message()`, `_build_action_buttons(row)`, `_add_save_button(row)`, `show_result`, `show_citation`, `_on_anchor_clicked`, `output_path`, `load_output`, `_refresh_open_btn`, `_open_preview` used consistently across base + both subclasses. `_research_targets(metadata, plan)` matches oppose's signature; `research_arguments(..., max_workers=, on_progress=, cache_dir=)` matches its real signature. ✓
