# Depo Config Dialog Extensions — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to execute this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend `DepoSummaryConfigDialog` with drag-and-drop topic reordering, summary bias control, and context-document drop zone. Phase 2 reads dropped docs and concatenates their text into the summary prompt.

**Spec:** `docs/superpowers/specs/2026-05-14-depo-config-extensions-design.md`

**Tech Stack:** PySide6 (`QListWidget` with `InternalMove` drag mode + `setAcceptDrops`), pytest + pytest-qt, existing `DocumentProcessor` / `extract_docx_text` / Word COM helpers.

---

## File map

**Modified:**
- `icharlotte_core/ui/depo_summary_config_dialog.py` — Topics → QListWidget, new bias row, new context-docs panel, updated `accept()` validation
- `Scripts/summarize_deposition.py` — `_extract_context_documents`, `_extract_doc_via_word_com`, `_resolve_bias_directive`, updated `_build_topic_locked_prompt`, updated `process_summary`
- `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` — two new placeholders (`{bias_directive}`, `{context_section}`), two new rules
- `tests/test_deposition/test_depo_summary_config_dialog.py` — 9 new tests
- `tests/test_deposition/test_summarize_deposition_phases.py` — 6 new tests
- `tests/test_deposition/test_full_flow_smoke.py` — minor field additions

---

## Task 1: Topics panel → `QListWidget` with drag-reorder

Replace the existing `QVBoxLayout` of `_TopicRow` widgets with a `QListWidget` configured for internal-move drag-reorder. Update the dialog's API: `self.topic_rows` (list) becomes `self.topics_list` (QListWidget) plus a new helper `self.topic_rows_in_order()` that yields each row widget in current visual order. Update existing tests that iterate `topic_rows` to use the helper.

**Files:**
- Modify: `icharlotte_core/ui/depo_summary_config_dialog.py`
- Modify: `tests/test_deposition/test_depo_summary_config_dialog.py` (update existing tests; add 1 new test)

- [ ] **Step 1: Update the dialog**

In `icharlotte_core/ui/depo_summary_config_dialog.py`:

a) Update imports — add `QListWidget`, `QListWidgetItem` to the `PySide6.QtWidgets` import line.

b) Replace the existing topics-section block (the `topics_container`, `topics_layout`, `self.topic_rows` list, and the `QScrollArea` wrapping) with:

```python
        # Topics list with drag-reorder
        root.addWidget(QLabel("Topics (drag to reorder, uncheck to omit, edit text to rename):"))
        self.topics_list = QListWidget()
        self.topics_list.setDragDropMode(QListWidget.InternalMove)
        self.topics_list.setSelectionMode(QListWidget.SingleSelection)
        self.topics_list.setDefaultDropAction(Qt.MoveAction)
        for t in self._session.get("topics", []):
            row = _TopicRow(t.get("title", ""))
            item = QListWidgetItem()
            item.setSizeHint(row.sizeHint())
            self.topics_list.addItem(item)
            self.topics_list.setItemWidget(item, row)
        root.addWidget(self.topics_list, 1)
```

c) Add a helper method to `DepoSummaryConfigDialog`:

```python
    def topic_rows_in_order(self):
        """Yield each topic _TopicRow widget in current visual order."""
        for i in range(self.topics_list.count()):
            item = self.topics_list.item(i)
            yield self.topics_list.itemWidget(item)
```

d) Update `accept()`'s `selected_topics` comprehension to use the helper instead of `self.topic_rows`:

```python
        selected_topics = [
            row.title_edit.text().strip()
            for row in self.topic_rows_in_order()
            if row.checkbox.isChecked() and row.title_edit.text().strip()
        ]
```

- [ ] **Step 2: Update the existing tests that reference `topic_rows`**

The existing tests in `tests/test_deposition/test_depo_summary_config_dialog.py` use `dlg.topic_rows`. Replace with `list(dlg.topic_rows_in_order())`. Specifically:

- `test_dialog_loads_session_and_populates_topics` (3 occurrences of `dlg.topic_rows`)
- `test_dialog_accept_writes_user_config_back_to_session` (2 occurrences of `dlg.topic_rows[...]`)
- `test_dialog_cancel_does_not_modify_session` (1 occurrence)
- `test_dialog_accept_blocks_when_no_topics_selected` (1 occurrence in the `for row in dlg.topic_rows: row.checkbox.setChecked(False)` line)

For tests that index by integer (e.g., `dlg.topic_rows[1]`), wrap with `list(...)`: `list(dlg.topic_rows_in_order())[1]`.

- [ ] **Step 3: Add the drag-reorder test**

Append to `tests/test_deposition/test_depo_summary_config_dialog.py`:

```python
def test_dialog_topic_drag_reorder_changes_selected_topics_order(qtbot, tmp_path):
    """Programmatically reorder topics via QListWidget and verify selected_topics order."""
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    # Initial order from the fixture: Pre-Accident History, Mechanism Of Injury, Damages.
    # Move "Damages" (row 2) to the top.
    item = dlg.topics_list.takeItem(2)
    # When takeItem is called, the item widget is destroyed; rebuild the row for the test
    # by re-creating the topic row at the new index.
    from icharlotte_core.ui.depo_summary_config_dialog import _TopicRow
    from PySide6.QtWidgets import QListWidgetItem
    new_row = _TopicRow("Damages")
    new_item = QListWidgetItem()
    new_item.setSizeHint(new_row.sizeHint())
    dlg.topics_list.insertItem(0, new_item)
    dlg.topics_list.setItemWidget(new_item, new_row)

    dlg.accept()

    loaded = session_manager.read_session(session_path)
    cfg = loaded["user_config"]
    assert cfg["selected_topics"] == ["Damages", "Pre-Accident History", "Mechanism Of Injury"]
```

- [ ] **Step 4: Run tests**

```
python -m pytest tests/test_deposition/test_depo_summary_config_dialog.py -v
```

Expected: 6 tests pass (5 existing — updated to use helper — + 1 new drag-reorder test).

Then full suite:

```
python -m pytest tests/test_deposition/ -v
```

Expected: 25 pass (24 existing + 1 new).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/depo_summary_config_dialog.py tests/test_deposition/test_depo_summary_config_dialog.py
git commit -m "feat(ui): topics list with drag-reorder in DepoSummaryConfigDialog"
```

---

## Task 2: Bias control (combo + custom field)

Add a "Summary bias" row between the topics list and the additional-topics field. Combobox with four options; a sibling `QLineEdit` that's visible only when *Custom…* is selected. Switching away clears the custom field. Accept writes `bias` and `bias_custom` to `user_config`.

**Files:**
- Modify: `icharlotte_core/ui/depo_summary_config_dialog.py`
- Modify: `tests/test_deposition/test_depo_summary_config_dialog.py` (add 3 new tests)

- [ ] **Step 1: Update the dialog**

In `icharlotte_core/ui/depo_summary_config_dialog.py`:

a) Add `QComboBox` to the `PySide6.QtWidgets` import.

b) Add this block in `__init__` between the topics list (`root.addWidget(self.topics_list, 1)`) and the "Additional topics" label/field:

```python
        # Summary bias row
        bias_row = QHBoxLayout()
        bias_row.addWidget(QLabel("Summary bias:"))
        self.bias_combo = QComboBox()
        # Display text → internal value
        for label, value in (
            ("Neutral", "neutral"),
            ("Most favorable to plaintiff", "pro_plaintiff"),
            ("Most favorable to defense", "pro_defense"),
            ("Custom…", "custom"),
        ):
            self.bias_combo.addItem(label, value)
        bias_row.addWidget(self.bias_combo)
        self.bias_custom_edit = QLineEdit()
        self.bias_custom_edit.setPlaceholderText(
            "Describe the editorial lens (e.g., 'Highlight any inconsistencies in injury testimony')"
        )
        self.bias_custom_edit.setVisible(False)
        bias_row.addWidget(self.bias_custom_edit, 1)
        root.addLayout(bias_row)

        self.bias_combo.currentIndexChanged.connect(self._on_bias_combo_changed)
```

c) Add the slot to `DepoSummaryConfigDialog`:

```python
    def _on_bias_combo_changed(self, _index):
        is_custom = self.bias_combo.currentData() == "custom"
        self.bias_custom_edit.setVisible(is_custom)
        if not is_custom:
            self.bias_custom_edit.clear()
```

d) Update `accept()` to include bias fields in the `cfg` dict. Add these keys to the existing `cfg = {...}`:

```python
            "bias": self.bias_combo.currentData() or "neutral",
            "bias_custom": (self.bias_custom_edit.text().strip()
                            if self.bias_combo.currentData() == "custom" else ""),
```

- [ ] **Step 2: Add the new tests**

Append to `tests/test_deposition/test_depo_summary_config_dialog.py`:

```python
def test_dialog_bias_combo_defaults_to_neutral_and_writes_neutral_to_session(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    assert dlg.bias_combo.currentData() == "neutral"
    assert dlg.bias_custom_edit.isVisible() is False

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["bias"] == "neutral"
    assert cfg["bias_custom"] == ""


def test_dialog_bias_custom_reveals_text_field_and_round_trips(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    # Pick "Custom…" (last item, internal value 'custom')
    custom_idx = next(
        i for i in range(dlg.bias_combo.count())
        if dlg.bias_combo.itemData(i) == "custom"
    )
    dlg.bias_combo.setCurrentIndex(custom_idx)
    assert dlg.bias_custom_edit.isVisible() is True

    dlg.bias_custom_edit.setText("Highlight inconsistencies in injury testimony.")
    dlg.accept()

    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["bias"] == "custom"
    assert cfg["bias_custom"] == "Highlight inconsistencies in injury testimony."


def test_dialog_bias_switching_away_from_custom_clears_custom_field(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    custom_idx = next(
        i for i in range(dlg.bias_combo.count())
        if dlg.bias_combo.itemData(i) == "custom"
    )
    pro_def_idx = next(
        i for i in range(dlg.bias_combo.count())
        if dlg.bias_combo.itemData(i) == "pro_defense"
    )
    dlg.bias_combo.setCurrentIndex(custom_idx)
    dlg.bias_custom_edit.setText("Some custom directive")
    dlg.bias_combo.setCurrentIndex(pro_def_idx)
    assert dlg.bias_custom_edit.isVisible() is False
    assert dlg.bias_custom_edit.text() == ""

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["bias"] == "pro_defense"
    assert cfg["bias_custom"] == ""
```

- [ ] **Step 3: Run tests**

```
python -m pytest tests/test_deposition/test_depo_summary_config_dialog.py -v
python -m pytest tests/test_deposition/ -v
```

Expected: 28 tests total pass (25 + 3 new).

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/ui/depo_summary_config_dialog.py tests/test_deposition/test_depo_summary_config_dialog.py
git commit -m "feat(ui): bias control combo with custom directive field"
```

---

## Task 3: Context documents panel (drop zone + add button + remove buttons)

Add a `QListWidget` styled drop zone between the bias row and the settings row. Accepts `.pdf`, `.doc`, `.docx` via drag-drop and via an "Add files…" button. Each entry shows the basename + an `×` remove button. The dialog tracks paths in `self._context_doc_paths`. Accept writes the list to `user_config.context_doc_paths`, filtered to files still present on disk.

**Files:**
- Modify: `icharlotte_core/ui/depo_summary_config_dialog.py`
- Modify: `tests/test_deposition/test_depo_summary_config_dialog.py` (add 5 new tests)

- [ ] **Step 1: Update the dialog**

In `icharlotte_core/ui/depo_summary_config_dialog.py`:

a) Add to imports: `QFileDialog`, `QPushButton`, `QTimer` (from `PySide6.QtCore`), `QUrl` is not needed since we read `event.mimeData().urls()` directly. Also add `Qt` is already imported.

b) Add this block in `__init__` between the bias row and the settings row:

```python
        # Context documents drop zone
        ctx_header = QHBoxLayout()
        ctx_header.addWidget(QLabel("Context documents (drop .pdf, .doc, .docx here):"))
        ctx_header.addStretch(1)
        self._ctx_add_btn = QPushButton("Add files…")
        self._ctx_add_btn.clicked.connect(self._on_add_context_files)
        ctx_header.addWidget(self._ctx_add_btn)
        root.addLayout(ctx_header)

        self.context_docs_list = QListWidget()
        self.context_docs_list.setFixedHeight(80)
        self.context_docs_list.setAcceptDrops(True)
        # Wire drop handlers
        self.context_docs_list.dragEnterEvent = self._context_drag_enter
        self.context_docs_list.dragMoveEvent = self._context_drag_enter  # same predicate
        self.context_docs_list.dropEvent = self._context_drop
        root.addWidget(self.context_docs_list)

        self._ctx_status_label = QLabel("")
        self._ctx_status_label.setStyleSheet("color: #c62828; font-style: italic; font-size: 11px;")
        root.addWidget(self._ctx_status_label)

        self._context_doc_paths: list[Path] = []
        self._ctx_status_clear_timer = QTimer(self)
        self._ctx_status_clear_timer.setSingleShot(True)
        self._ctx_status_clear_timer.timeout.connect(lambda: self._ctx_status_label.setText(""))
```

c) Add the slots and helpers to `DepoSummaryConfigDialog`:

```python
    _CONTEXT_DOC_EXTS = (".pdf", ".doc", ".docx")

    def _context_drag_enter(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def _context_drop(self, event):
        urls = event.mimeData().urls()
        # Defer to avoid blocking the OS shell (Explorer)
        QTimer.singleShot(0, lambda: self._add_context_paths_from_urls(urls))
        event.acceptProposedAction()

    def _add_context_paths_from_urls(self, urls):
        accepted, rejected = [], []
        for url in urls:
            if not url.isLocalFile():
                rejected.append(url.toString())
                continue
            p = Path(url.toLocalFile())
            if p.suffix.lower() in self._CONTEXT_DOC_EXTS and p.exists():
                accepted.append(p)
            else:
                rejected.append(p.name)
        for p in accepted:
            self._append_context_path(p)
        if rejected:
            self._show_ctx_status(
                f"Unsupported file type — skipped: {', '.join(rejected[:3])}"
                + (" …" if len(rejected) > 3 else "")
            )

    def _append_context_path(self, path: Path):
        if path in self._context_doc_paths:
            return  # de-dupe
        self._context_doc_paths.append(path)
        item = QListWidgetItem()
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(4, 0, 4, 0)
        row_layout.addWidget(QLabel(path.name), 1)
        remove_btn = QPushButton("×")
        remove_btn.setFixedSize(20, 20)
        remove_btn.setStyleSheet("QPushButton { color: #c62828; font-weight: bold; }")
        remove_btn.clicked.connect(lambda _checked=False, p=path: self._remove_context_path(p))
        row_layout.addWidget(remove_btn)
        item.setSizeHint(row.sizeHint())
        self.context_docs_list.addItem(item)
        self.context_docs_list.setItemWidget(item, row)

    def _remove_context_path(self, path: Path):
        if path not in self._context_doc_paths:
            return
        idx = self._context_doc_paths.index(path)
        self._context_doc_paths.pop(idx)
        self.context_docs_list.takeItem(idx)

    def _on_add_context_files(self):
        paths, _filter = QFileDialog.getOpenFileNames(
            self, "Add context documents", "",
            "Documents (*.pdf *.doc *.docx)",
        )
        for p in paths:
            path = Path(p)
            if path.suffix.lower() in self._CONTEXT_DOC_EXTS and path.exists():
                self._append_context_path(path)

    def _show_ctx_status(self, msg: str):
        self._ctx_status_label.setText(msg)
        self._ctx_status_clear_timer.start(3000)
```

d) Update `accept()` to write context paths (filtered to still-extant files). Add this line just before the existing `cfg = {...}`:

```python
        context_doc_paths = [str(p.resolve()) for p in self._context_doc_paths if p.exists()]
```

Then include in the `cfg` dict:

```python
            "context_doc_paths": context_doc_paths,
```

e) Bump the dialog default size:

```python
        self.resize(800, 750)
```

- [ ] **Step 2: Add the new tests**

Append to `tests/test_deposition/test_depo_summary_config_dialog.py`:

```python
def test_dialog_context_docs_accept_via_add_button(qtbot, tmp_path):
    """Programmatically simulate the 'Add files…' flow with two real temp files."""
    session_path = _make_session(tmp_path)
    pdf_path = tmp_path / "complaint.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")
    docx_path = tmp_path / "med_summary.docx"
    docx_path.write_bytes(b"PK\x03\x04")  # arbitrary docx marker

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._append_context_path(pdf_path)
    dlg._append_context_path(docx_path)
    assert dlg.context_docs_list.count() == 2

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    paths = cfg["context_doc_paths"]
    assert any("complaint.pdf" in p for p in paths)
    assert any("med_summary.docx" in p for p in paths)


def test_dialog_context_docs_reject_unsupported_extensions(qtbot, tmp_path):
    """Dropping a .txt file is rejected and surfaces a status message."""
    from PySide6.QtCore import QUrl
    session_path = _make_session(tmp_path)
    txt_path = tmp_path / "notes.txt"
    txt_path.write_text("nope", encoding="utf-8")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._add_context_paths_from_urls([QUrl.fromLocalFile(str(txt_path))])

    assert dlg.context_docs_list.count() == 0
    assert "Unsupported" in dlg._ctx_status_label.text()
    assert "notes.txt" in dlg._ctx_status_label.text()


def test_dialog_context_docs_remove_button_drops_path(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    pdf_a = tmp_path / "a.pdf"
    pdf_a.write_bytes(b"%PDF-1.4\n")
    pdf_b = tmp_path / "b.pdf"
    pdf_b.write_bytes(b"%PDF-1.4\n")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._append_context_path(pdf_a)
    dlg._append_context_path(pdf_b)
    dlg._remove_context_path(pdf_a)
    assert dlg.context_docs_list.count() == 1
    assert pdf_a not in dlg._context_doc_paths
    assert pdf_b in dlg._context_doc_paths

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert len(cfg["context_doc_paths"]) == 1
    assert "b.pdf" in cfg["context_doc_paths"][0]


def test_dialog_context_docs_drop_via_mime(qtbot, tmp_path):
    """Simulate a drop event built from QMimeData."""
    from PySide6.QtCore import QMimeData, QPoint, QUrl
    from PySide6.QtGui import QDropEvent
    from PySide6.QtCore import Qt as _Qt

    session_path = _make_session(tmp_path)
    pdf_path = tmp_path / "ev.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(pdf_path))])
    event = QDropEvent(
        QPoint(0, 0), _Qt.CopyAction, mime, _Qt.LeftButton, _Qt.NoModifier
    )
    dlg._context_drop(event)
    # _context_drop defers via QTimer.singleShot(0, …) — pump the event loop once.
    qtbot.wait(50)

    assert dlg.context_docs_list.count() == 1
    assert pdf_path in dlg._context_doc_paths


def test_dialog_context_docs_missing_at_accept_are_silently_dropped(qtbot, tmp_path):
    session_path = _make_session(tmp_path)
    pdf_path = tmp_path / "gone.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")

    dlg = DepoSummaryConfigDialog(session_path)
    qtbot.addWidget(dlg)
    dlg._append_context_path(pdf_path)
    pdf_path.unlink()  # remove file after it was added

    dlg.accept()
    cfg = session_manager.read_session(session_path)["user_config"]
    assert cfg["context_doc_paths"] == []
```

- [ ] **Step 3: Run tests**

```
python -m pytest tests/test_deposition/test_depo_summary_config_dialog.py -v
python -m pytest tests/test_deposition/ -v
```

Expected: 33 tests total pass (28 + 5 new).

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/ui/depo_summary_config_dialog.py tests/test_deposition/test_depo_summary_config_dialog.py
git commit -m "feat(ui): context documents drop zone in DepoSummaryConfigDialog"
```

---

## Task 4: Bias directive resolver + prompt template update

Add `_resolve_bias_directive(cfg)` helper. Update `_build_topic_locked_prompt` to accept `bias_directive` and `context_documents` keyword args and substitute the two new placeholders. Update the prompt template file with the new placeholders and rules. Update `process_summary` to call the resolver and pass the result.

For this task, **context_documents is always an empty list** when calling `_build_topic_locked_prompt` from `process_summary` — context extraction wiring comes in Task 5. This lets us merge the prompt-template change and the bias change as a single coherent commit and keep Task 5 focused on extraction.

**Files:**
- Modify: `Scripts/summarize_deposition.py`
- Modify: `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` (full rewrite)
- Modify: `tests/test_deposition/test_summarize_deposition_phases.py` (add bias tests; update existing prompt tests if needed)

- [ ] **Step 1: Update the prompt template**

Replace the entire contents of `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` with:

```
You are summarizing a deposition transcript using a fixed set of topic headings supplied by the attorney. You do not choose the topics. You do not add or omit topics. You write bullet points under the headings provided.

DEPONENT LABEL: {deponent_label}
BULLETS PER TOPIC: {bullets_per_topic}

SUMMARY BIAS / EDITORIAL LENS:
{bias_directive}
{context_section}

TOPIC HEADINGS (use exactly these, in this order):
{topic_list}

OUTPUT RULES:
1. Start the summary with this sentence: "The parties took the deposition of [name of deponent] on [date of deposition]. Below are the most salient portions of [name of deponent]'s testimony:"
2. Under each topic heading, write exactly {bullets_per_topic} bullet points summarizing what the deponent testified regarding that topic. If the transcript contains less testimony on a topic than requested, write fewer bullets — never invent testimony.
3. Each topic heading appears on its own line, in bold (use **Heading** markdown), title-cased, with no numbering and no trailing punctuation.
4. Each bullet is at least two complete sentences. Use markdown dashes ("- ") at the start of each bullet.
5. Refer to the deponent as "{deponent_label}" throughout. Do not use the deponent's full name except in the opening sentence.
6. Summarize testimony directly. Do not use introductory clauses like "Regarding," "Concerning," "With respect to," or "When asked about." Write "Plaintiff testified he broke his leg," not "Regarding his injuries, Plaintiff testified he broke his leg."
7. Avoid repeating phrases that flag content as testimony. Write "Joe texted Plaintiff photos," not "Plaintiff testified that Joe texted her photos."
8. Apply the SUMMARY BIAS instruction above when selecting which testimony to emphasize, while never inventing or distorting what the deponent actually said.
9. When ADDITIONAL CASE CONTEXT is provided, use it to inform your selection and phrasing of testimony, especially to identify contradictions or admissions that bear on the case. The deposition transcript itself remains the only source for direct testimony content — do not attribute statements from context documents to the deponent.
10. Do not add introductory or concluding paragraphs beyond rule 1's opening sentence. Do not add an exhibits section, impeachment section, or any section not in the topic list.

CUSTOM RULES FROM THE ATTORNEY (apply in addition to the rules above):
{custom_rules}
```

- [ ] **Step 2: Add bias resolver and update `_build_topic_locked_prompt`**

In `Scripts/summarize_deposition.py`:

a) Add a new helper `_resolve_bias_directive`. Place it just before `_build_topic_locked_prompt`:

```python
_BIAS_DIRECTIVES = {
    "neutral": (
        "Maintain a neutral, balanced tone. Include both favorable and unfavorable "
        "testimony equally."
    ),
    "pro_plaintiff": (
        "Emphasize testimony most favorable to the plaintiff's case. Include testimony "
        "favorable to the defense only when it directly contradicts the plaintiff's claims."
    ),
    "pro_defense": (
        "Emphasize testimony most favorable to the defense. Include testimony favorable "
        "to the plaintiff only when it directly bears on the defense's case."
    ),
}


def _resolve_bias_directive(cfg: dict) -> str:
    """Map cfg.bias to the directive string injected into the prompt."""
    bias = cfg.get("bias", "neutral")
    if bias == "custom":
        return (cfg.get("bias_custom") or "").strip() or _BIAS_DIRECTIVES["neutral"]
    return _BIAS_DIRECTIVES.get(bias, _BIAS_DIRECTIVES["neutral"])
```

b) Replace `_build_topic_locked_prompt` with:

```python
def _build_topic_locked_prompt(base_prompt: str, *, topic_list: list, bullets_per_topic: int,
                                deponent_label: str, custom_rules: str,
                                bias_directive: str = "",
                                context_documents: list = None) -> str:
    """Render the topic-locked summary prompt with user-supplied substitutions.

    User-supplied strings (deponent_label, custom_rules, bias_directive) are stripped
    of literal `{` and `}` characters to prevent placeholder-leak across slots.
    """
    def _strip_braces(s: str) -> str:
        return (s or "").replace("{", "").replace("}", "")

    rendered_topics = "\n".join(f"- {t}" for t in topic_list)

    if context_documents:
        ctx_blocks = []
        for doc in context_documents:
            ctx_blocks.append(
                f"=== CONTEXT DOC: {doc['filename']} ===\n{doc['text']}"
            )
        context_section = (
            "\n\nADDITIONAL CASE CONTEXT (read these in addition to the deposition "
            "transcript to better inform your summary):\n\n" + "\n\n".join(ctx_blocks)
        )
    else:
        context_section = ""

    return (base_prompt
            .replace("{deponent_label}", _strip_braces(deponent_label))
            .replace("{bullets_per_topic}", str(bullets_per_topic))
            .replace("{topic_list}", rendered_topics)
            .replace("{bias_directive}", _strip_braces(bias_directive))
            .replace("{context_section}", context_section)
            .replace("{custom_rules}", _strip_braces(custom_rules) or "(none)"))
```

c) Update `process_summary` to call the bias resolver and pass it to the prompt builder. Find the existing block:

```python
    prompt = _build_topic_locked_prompt(
        base_prompt,
        topic_list=final_topics,
        bullets_per_topic=cfg.get("bullets_per_topic", 5),
        deponent_label=cfg.get("deponent_label") or session.get("deponent_type", "Deponent"),
        custom_rules=cfg.get("custom_rules", ""),
    )
```

Replace with:

```python
    bias_directive = _resolve_bias_directive(cfg)
    prompt = _build_topic_locked_prompt(
        base_prompt,
        topic_list=final_topics,
        bullets_per_topic=cfg.get("bullets_per_topic", 5),
        deponent_label=cfg.get("deponent_label") or session.get("deponent_type", "Deponent"),
        custom_rules=cfg.get("custom_rules", ""),
        bias_directive=bias_directive,
        context_documents=[],  # Task 5 wires real extraction
    )
```

- [ ] **Step 3: Add bias tests**

Append to `tests/test_deposition/test_summarize_deposition_phases.py`:

```python
@pytest.mark.parametrize("bias_value,bias_custom,expected_substring", [
    ("neutral", "", "neutral, balanced tone"),
    ("pro_plaintiff", "", "most favorable to the plaintiff"),
    ("pro_defense", "", "most favorable to the defense"),
    ("custom", "Highlight inconsistencies in injury testimony.",
     "Highlight inconsistencies in injury testimony."),
])
def test_phase2_resolves_bias_directive_for_each_preset(
    tmp_path, monkeypatch, bias_value, bias_custom, expected_substring
):
    """Each bias preset routes the expected directive language into the prompt."""
    session_path = tmp_path / "session.json"
    cached_path = tmp_path / "session.txt"
    cached_path.write_text(FAKE_TRANSCRIPT, encoding="utf-8")
    session_manager.write_session(session_path, {
        "version": 1,
        "phase": "ready_for_summary",
        "input_path": str(tmp_path / "X.pdf"),
        "cached_text_path": str(cached_path),
        "deponent_name": "X",
        "deposition_date": "Jan 1, 2024",
        "deponent_type": "Plaintiff",
        "file_number": "0000.000",
        "topics": [{"id": 1, "title": "T", "rank": 1, "discussion_density": "high"}],
        "user_config": {
            "selected_topics": ["T"],
            "added_topics": [],
            "bullets_per_topic": 5,
            "deponent_label": "Plaintiff",
            "custom_rules": "",
            "cross_check_enabled": False,
            "bias": bias_value,
            "bias_custom": bias_custom,
            "context_doc_paths": [],
        },
    })

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)

    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        return "**T**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    logger = summarize_deposition.AgentLogger("BiasTest", log_to_file=False)
    summarize_deposition.process_summary(str(session_path), logger)
    assert expected_substring in captured["prompt"]
```

Also update the existing `_write_ready_session` helper to include the new fields with default values. Locate the existing helper in the file and add to its `user_config` dict:

```python
            "bias": "neutral",
            "bias_custom": "",
            "context_doc_paths": [],
```

This keeps all existing phase-2 tests working since they call the helper.

- [ ] **Step 4: Run tests**

```
python -m pytest tests/test_deposition/test_summarize_deposition_phases.py -v
python -m pytest tests/test_deposition/ -v
```

Expected: 37 tests total pass (33 + 4 new parametrized).

- [ ] **Step 5: Commit**

```bash
git add Scripts/summarize_deposition.py Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt tests/test_deposition/test_summarize_deposition_phases.py
git commit -m "feat(deposition): bias directive resolver + prompt template placeholders"
```

---

## Task 5: Context document extractor (helpers + orchestration)

Add `_extract_context_documents` and `_extract_doc_via_word_com`. Wire into `process_summary` so the extracted docs flow into `_build_topic_locked_prompt`. Add progress markers.

**Files:**
- Modify: `Scripts/summarize_deposition.py`
- Modify: `tests/test_deposition/test_summarize_deposition_phases.py` (add 5 new tests)

- [ ] **Step 1: Add the helpers**

In `Scripts/summarize_deposition.py`, add this block right before `_resolve_bias_directive`:

```python
_CONTEXT_DOC_PER_DOC_CHAR_CAP = 100_000


def _extract_doc_via_word_com(path: str, logger) -> str:
    """Extract text from a legacy .doc via Word COM. Returns empty string on failure.

    Mirrors the ChatTab._extract_doc_text pattern: attach to the user's running Word
    instance (NEVER call word.Quit() or set word.Visible). Open the doc read-only and
    only close the Document we opened.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError:
        logger.warning("win32com not available; cannot extract .doc files")
        return ""
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(
            FileName=os.path.abspath(path),
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
        )
        text = doc.Content.Text or ""
        return text
    except Exception as e:
        logger.warning(f".doc extraction failed for {path}: {e}")
        return ""
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        # Intentionally do NOT word.Quit() — attached to user's running Word


def _extract_context_documents(paths, logger) -> list:
    """Extract text from each context document. Returns [{filename, text}, ...].

    Per-doc failures (missing, unsupported, empty, extraction error) log a warning
    and skip that doc. Per-doc character cap of 100,000 with a truncation marker.
    """
    from pathlib import Path as _Path
    results = []
    for path in paths:
        p = _Path(path)
        if not p.exists():
            logger.warning(f"Context doc missing, skipping: {path}")
            continue
        ext = p.suffix.lower()
        try:
            if ext == ".pdf":
                processor = DocumentProcessor(ocr_config=OCRConfig(adaptive=True), logger=logger)
                result = processor.extract_with_dynamic_ocr(str(p))
                text = result.text if result.success else ""
            elif ext == ".docx":
                from icharlotte_core.document_processor import extract_docx_text
                text = extract_docx_text(str(p))
            elif ext == ".doc":
                text = _extract_doc_via_word_com(str(p), logger)
            else:
                logger.warning(f"Unsupported context doc extension {ext}, skipping: {path}")
                continue
        except Exception as e:
            logger.warning(f"Context doc extraction failed for {path}: {e}")
            continue
        if not text:
            logger.warning(f"Context doc produced no text, skipping: {path}")
            continue
        if len(text) > _CONTEXT_DOC_PER_DOC_CHAR_CAP:
            text = (text[:_CONTEXT_DOC_PER_DOC_CHAR_CAP]
                    + f"\n\n[...truncated at {_CONTEXT_DOC_PER_DOC_CHAR_CAP} chars]")
        results.append({"filename": p.name, "text": text})
    return results
```

- [ ] **Step 2: Wire into `process_summary`**

In `process_summary`, between the `logger.progress(5, ...)` session-load and the `logger.progress(10, ...)` cached-text-read, add context extraction:

```python
    context_doc_paths = cfg.get("context_doc_paths") or []
    if context_doc_paths:
        for i, p in enumerate(context_doc_paths, 1):
            logger.progress(
                5 + min(4, (i * 4) // max(1, len(context_doc_paths))),
                f"Extracting context: {os.path.basename(p)} ({i}/{len(context_doc_paths)})...",
            )
        context_documents = _extract_context_documents(context_doc_paths, logger)
    else:
        context_documents = []
```

(The progress percentages fill 5–9% across however many docs there are. They're posted before extraction starts; we accept that minor inaccuracy in exchange for simpler code.)

Then update the `_build_topic_locked_prompt` call to pass real `context_documents`:

```python
        context_documents=context_documents,
```

(Replace the `context_documents=[],` placeholder added in Task 4.)

- [ ] **Step 3: Add the new tests**

Append to `tests/test_deposition/test_summarize_deposition_phases.py`:

```python
def test_phase2_concatenates_context_documents_into_prompt(tmp_path, monkeypatch):
    """Two context docs (one PDF, one DOCX) are extracted and injected into the prompt."""
    pdf_path = tmp_path / "complaint.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")
    docx_path = tmp_path / "med.docx"
    docx_path.write_bytes(b"PK\x03\x04")

    # Mock both extractors
    from types import SimpleNamespace
    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        lambda self, p: SimpleNamespace(
            success=True, text="PDF context text", char_count=16, page_count=1,
            ocr_pages=[], ocr_percentage=0.0, error=None,
        ),
    )
    monkeypatch.setattr(
        "icharlotte_core.document_processor.extract_docx_text",
        lambda p: "DOCX context text",
    )

    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(pdf_path), str(docx_path)]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)

    captured = {}

    def fake_call(self, prompt, text, task_type=None, **kw):
        captured["prompt"] = prompt
        return "**T**\n- B."

    monkeypatch.setattr(summarize_deposition.LLMCaller, "call", fake_call)

    logger = summarize_deposition.AgentLogger("CtxTest", log_to_file=False)
    summarize_deposition.process_summary(str(session_path), logger)

    p = captured["prompt"]
    assert "=== CONTEXT DOC: complaint.pdf ===" in p
    assert "PDF context text" in p
    assert "=== CONTEXT DOC: med.docx ===" in p
    assert "DOCX context text" in p


def test_phase2_per_doc_char_cap_truncates_long_docs(tmp_path, monkeypatch):
    pdf_path = tmp_path / "big.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n")

    huge = "X" * 200_000
    from types import SimpleNamespace
    monkeypatch.setattr(
        summarize_deposition.DocumentProcessor,
        "extract_with_dynamic_ocr",
        lambda self, p: SimpleNamespace(
            success=True, text=huge, char_count=len(huge), page_count=100,
            ocr_pages=[], ocr_percentage=0.0, error=None,
        ),
    )

    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(pdf_path)]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)

    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("TruncTest", log_to_file=False),
    )
    assert "[...truncated at 100000 chars]" in captured["prompt"]


def test_phase2_missing_context_doc_logged_and_skipped(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(tmp_path / "ghost.pdf")]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    logger = summarize_deposition.AgentLogger("MissTest", log_to_file=False)
    ok = summarize_deposition.process_summary(str(session_path), logger)
    assert ok is True
    assert "=== CONTEXT DOC:" not in captured["prompt"]


def test_phase2_no_context_docs_leaves_context_section_empty(tmp_path, monkeypatch):
    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("EmptyCtxTest", log_to_file=False),
    )
    assert "ADDITIONAL CASE CONTEXT" not in captured["prompt"]
    assert "=== CONTEXT DOC:" not in captured["prompt"]


def test_phase2_doc_extension_unsupported_skipped(tmp_path, monkeypatch):
    txt_path = tmp_path / "notes.txt"
    txt_path.write_text("not a supported doc", encoding="utf-8")

    session_path = _write_ready_session(
        tmp_path, cross_check=False, selected=["T"], added=[],
    )
    session = session_manager.read_session(session_path)
    session["user_config"]["context_doc_paths"] = [str(txt_path)]
    session_manager.write_session(session_path, session)

    monkeypatch.setattr(summarize_deposition, "save_to_docx", lambda *a, **kw: True)
    monkeypatch.setattr(summarize_deposition, "_register_outputs", lambda *a, **kw: None,
                        raising=False)
    captured = {}
    monkeypatch.setattr(
        summarize_deposition.LLMCaller, "call",
        lambda self, prompt, text, task_type=None, **kw: (captured.update(prompt=prompt) or "**T**\n- B."),
    )

    ok = summarize_deposition.process_summary(
        str(session_path),
        summarize_deposition.AgentLogger("ExtTest", log_to_file=False),
    )
    assert ok is True
    assert "=== CONTEXT DOC:" not in captured["prompt"]
```

- [ ] **Step 4: Update the smoke test**

In `tests/test_deposition/test_full_flow_smoke.py`, locate the `session_manager.update_user_config(...)` call and add the new fields:

```python
    session_manager.update_user_config(session_path, {
        "selected_topics": ["Pre-Accident History", "Mechanism Of Injury"],
        "added_topics": [],
        "bullets_per_topic": 5,
        "deponent_label": "Plaintiff",
        "custom_rules": "",
        "cross_check_enabled": False,
        "bias": "neutral",
        "bias_custom": "",
        "context_doc_paths": [],
    })
```

- [ ] **Step 5: Run tests**

```
python -m pytest tests/test_deposition/ -v
```

Expected: 42 tests total pass (37 + 5 new).

- [ ] **Step 6: Commit**

```bash
git add Scripts/summarize_deposition.py tests/test_deposition/test_summarize_deposition_phases.py tests/test_deposition/test_full_flow_smoke.py
git commit -m "feat(deposition): context document extraction in phase 2"
```

---

## Task 6: Manual verification

Per CLAUDE.md, manually exercise the new UI.

**Files:** (none — manual)

- [ ] **Step 1: Launch the app and run the deposition agent on a real test PDF.**

- [ ] **Step 2: When the READY button appears, click it and verify:**
   - Topics list shows the agent's suggestions.
   - Bias dropdown defaults to "Neutral".
   - Bias custom field is hidden.
   - Context documents panel shows the empty list and an "Add files…" button.

- [ ] **Step 3: Drag-reorder three topics** (e.g., move the bottom topic to position 1). Generate. Open the docx and confirm the bold topic headings match the dragged order.

- [ ] **Step 4: Bias = Most favorable to defense, no context docs.** Generate. Eyeball the resulting bullets for defense-friendly emphasis.

- [ ] **Step 5: Bias = Custom… with "Highlight any inconsistencies in the deponent's testimony about injuries"**. Generate. Confirm the summary reflects the custom directive (look for inconsistency callouts).

- [ ] **Step 6: Drag a real complaint PDF and a med-summary DOCX into the context docs panel.** Generate. Confirm the summary mentions or cross-references the context (e.g., differences between testimony and the complaint's allegations).

- [ ] **Step 7: Drop a `.txt` file into the context panel.** Confirm it's rejected and the status label briefly shows "Unsupported file type — skipped".

- [ ] **Step 8: Add two docs, click the `×` on one to remove it.** Confirm only one path remains. Submit and confirm the agent only references the remaining doc.

If anything is broken, report it back so we can dispatch a fix.
