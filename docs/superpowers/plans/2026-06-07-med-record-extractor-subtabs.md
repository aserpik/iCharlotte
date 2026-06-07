# Med Record Extractor Subtabs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convert the Med Record Extractor chronology viewer into two subtabs with verbatim Brief Synopsis and chronology-row text, single-click toggle selection, auto-fit row heights, and globally persisted chronology table column widths.

**Architecture:** Keep the existing `MedChronologySelectionPage` as the settings page and replace the side-by-side splitter with a `QTabWidget`. Preserve matching and extraction behavior by keeping normalized parsing/matching paths internal while displaying raw Word paragraph/cell text. Persist table column widths with `QSettings("iCharlotte", "iCharlotte")` using one stable global key for the chronology table.

**Tech Stack:** Python, PySide6, python-docx, QSettings, pytest, pytest-qt, existing iCharlotte wizard task scaffolding.

---

## File Structure

- Modify `icharlotte_core/med_record_chronology.py`
  - Preserve verbatim Word paragraph/cell text for display.
  - Continue using normalized text for heading detection, stable IDs, matching, and warning previews.
  - Update `match_synopsis_to_rows()` to normalize paragraph text internally before date/provider extraction.
- Modify `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`
  - Replace the `QSplitter` layout with a `QTabWidget` containing `BriefSynopsisPanel` and `ChronologyTablePanel`.
  - Make single-click anywhere on a synopsis item toggle its checkbox.
  - Make single-click anywhere on an extractable chronology row toggle its checkbox.
  - Display all chronology source columns verbatim: `DATE`, `PAGE NO`, `PROVIDER`, `DESCRIPTION`, `Red Flags/Comments`.
  - Add global column-width persistence and row auto-fit behavior to `ChronologyTablePanel`.
- Modify `tests/test_med_record_extractor.py`
  - Add parser tests proving display text is verbatim while matching still works.
- Modify `tests/test_wizard/test_med_record_extractor_viewer.py`
  - Add subtab, click-toggle, verbatim display, column-width persistence, row-height, and restore regression tests.
- Do not modify `iCharlotte.py` or task routing for this feature unless an existing test exposes a break; the builder already returns `MedChronologySelectionPage`.

The repository currently has unrelated dirty files. Stage and commit only these paths for this feature:

```text
icharlotte_core/med_record_chronology.py
icharlotte_core/ui/wizard/pages/med_record_extractor_page.py
tests/test_med_record_extractor.py
tests/test_wizard/test_med_record_extractor_viewer.py
```

---

## Task 1: Preserve Verbatim Parser Display Text

**Files:**
- Modify: `tests/test_med_record_extractor.py`
- Modify: `icharlotte_core/med_record_chronology.py`

- [ ] **Step 1: Write failing parser display tests**

Append these tests to `TestChronologyDocumentParser` in `tests/test_med_record_extractor.py`:

```python
    def test_parse_chronology_document_preserves_verbatim_synopsis_text(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / "chronology.docx"
            doc = Document()
            doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
            paragraph = doc.add_paragraph()
            paragraph.add_run("On 09/21/2020, Test Plaintiff presented to Kaiser Permanente.")
            paragraph.add_run().add_break()
            paragraph.add_run("She was evaluated by Henry Louis Marr, DO.")
            table = doc.add_table(rows=1, cols=5)
            for index, header in enumerate([
                "DATE",
                "PAGE NO",
                "PROVIDER",
                "DESCRIPTION",
                "Red Flags/Comments",
            ]):
                table.rows[0].cells[index].text = header
            row = table.add_row().cells
            row[0].text = "09/21/2020"
            row[1].text = "Record\n\nPg No: 1/2"
            row[2].text = "Kaiser Permanente\nHenry Louis Marr, DO"
            row[3].text = "Emergency note"
            row[4].text = ""
            doc.save(source)

            parsed = parse_chronology_document(str(source))

        self.assertEqual(
            parsed.synopsis_paragraphs[0].text,
            "On 09/21/2020, Test Plaintiff presented to Kaiser Permanente.\n"
            "She was evaluated by Henry Louis Marr, DO.",
        )

    def test_parse_chronology_document_preserves_verbatim_row_cells(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(Path(td) / "chronology.docx")
            parsed = parse_chronology_document(str(source))

        self.assertEqual(
            parsed.rows[0].page_no,
            "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\n"
            "Pg No: 599-604/2530",
        )
        self.assertEqual(
            parsed.rows[0].provider,
            "Kaiser Permanente\nHenry Louis Marr, DO",
        )
        self.assertEqual(
            parsed.rows[0].description,
            "EMERGENCY DEPARTMENT NOTE\nChief Complaint: Ankle injury.",
        )

    def test_synopsis_matching_still_handles_verbatim_line_breaks(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document, match_synopsis_to_rows

        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / "chronology.docx"
            doc = Document()
            doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
            paragraph = doc.add_paragraph()
            paragraph.add_run("On 09/21/2020, Test Plaintiff presented to Kaiser Permanente.")
            paragraph.add_run().add_break()
            paragraph.add_run("She was evaluated by Henry Louis Marr, DO.")
            table = doc.add_table(rows=1, cols=5)
            for index, header in enumerate([
                "DATE",
                "PAGE NO",
                "PROVIDER",
                "DESCRIPTION",
                "Red Flags/Comments",
            ]):
                table.rows[0].cells[index].text = header
            row = table.add_row().cells
            row[0].text = "09/21/2020"
            row[1].text = "Record\n\nPg No: 1/2"
            row[2].text = "Kaiser Permanente\nHenry Louis Marr, DO"
            row[3].text = "Emergency note"
            row[4].text = ""
            doc.save(source)
            parsed = parse_chronology_document(str(source))

        result = match_synopsis_to_rows(parsed.synopsis_paragraphs[0], parsed.rows)

        self.assertEqual(result.status, "confident")
        self.assertEqual(result.row_ids, (parsed.rows[0].id,))
```

- [ ] **Step 2: Verify parser tests fail**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_med_record_extractor.py -k "verbatim or line_breaks" -q
```

Expected: the verbatim tests fail because `_parse_synopsis()` and `_parse_rows()` collapse display text.

- [ ] **Step 3: Implement verbatim parser display fields**

In `icharlotte_core/med_record_chronology.py`, update `match_synopsis_to_rows()` so matching uses normalized paragraph text internally:

```python
def match_synopsis_to_rows(
    paragraph: SynopsisParagraph,
    rows: list[SelectableChronologyRow],
) -> MatchResult:
    paragraph_text = _collapse(paragraph.text)
    dates = {_normalize_date(match.group(0)) for match in _DATE_RE.finditer(paragraph_text)}
    dates.discard("")
    if not dates:
        return MatchResult(status="none", reason="No date found in synopsis paragraph.")

    same_date_rows = [row for row in rows if _normalize_date(row.date) in dates]
    if not same_date_rows:
        return MatchResult(status="none", reason="No chronology rows share the synopsis date.")

    candidates = _provider_candidates(paragraph_text)
    if not candidates:
        return MatchResult(
            status="ambiguous",
            candidate_row_ids=tuple(row.id for row in same_date_rows),
            reason="No provider or treater name found in synopsis paragraph.",
        )
```

Keep the existing remainder of `match_synopsis_to_rows()` after the `candidates` block unchanged.

In `_parse_synopsis()`, preserve `block.text` for display and use collapsed text only for control flow and stable IDs:

```python
def _parse_synopsis(path: str) -> list[SynopsisParagraph]:
    doc = Document(path)
    in_synopsis = False
    paragraphs: list[SynopsisParagraph] = []
    for block in _iter_body_blocks(doc):
        if isinstance(block, Table):
            if in_synopsis and _is_chronology_table(block):
                break
            continue

        display_text = block.text
        match_text = _collapse(display_text)
        if not match_text:
            continue
        if _SYNOPSIS_HEADING_RE.match(match_text):
            in_synopsis = True
            continue
        if not in_synopsis:
            continue

        order = len(paragraphs)
        paragraphs.append(
            SynopsisParagraph(
                id=_stable_id("syn", order, match_text),
                order=order,
                text=display_text,
            )
        )
    return paragraphs
```

In `_parse_rows()`, keep normalized cells for filtering and IDs, but store raw cell text for display:

```python
def _parse_rows(path: str) -> list[SelectableChronologyRow]:
    doc = Document(path)
    for table in doc.tables:
        if not _is_chronology_table(table):
            continue

        rows: list[SelectableChronologyRow] = []
        for raw_row in table.rows[1:]:
            raw_cells = [cell.text for cell in raw_row.cells]
            cells = [_collapse(text) for text in raw_cells]
            if not cells[0]:
                continue
            if len(set(cells)) == 1:
                continue
            record_filename, page_start, page_end = _parse_page_no(raw_row.cells[1].text)
            warning = ""
            if not _is_extractable_page_range(record_filename, page_start, page_end):
                warning = f"Could not parse record/pages from PAGE NO: {cells[1][:80]}"
            order = len(rows)
            rows.append(
                SelectableChronologyRow(
                    id=_stable_id("row", order, "|".join(cells[:4])),
                    order=order,
                    date=raw_cells[0],
                    page_no=raw_cells[1],
                    provider=raw_cells[2],
                    description=raw_cells[3],
                    flags=raw_cells[4],
                    record_filename=record_filename,
                    page_start=page_start,
                    page_end=page_end,
                    warning=warning,
                )
            )
        return rows
    return []
```

- [ ] **Step 4: Run parser tests**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_med_record_extractor.py -k "ChronologyDocumentParser or SynopsisMatching" -q
```

Expected: all selected parser and matching tests pass.

- [ ] **Step 5: Commit parser change**

Run:

```powershell
git add -- icharlotte_core\med_record_chronology.py tests\test_med_record_extractor.py
git diff --cached --check
git commit -m "fix(med-records): preserve chronology display text"
```

Expected: commit includes only the parser and parser test files.

---

## Task 2: Add Subtabs and Single-Click Toggle Behavior

**Files:**
- Modify: `tests/test_wizard/test_med_record_extractor_viewer.py`
- Modify: `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`

- [ ] **Step 1: Write failing subtab and click-toggle tests**

Append these tests to `tests/test_wizard/test_med_record_extractor_viewer.py`:

```python
def test_selection_page_uses_brief_synopsis_and_chronology_row_tabs(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    source = _build_chronology_docx(tmp_path / "chronology.docx")
    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)

    assert page.tab_widget.count() == 2
    assert page.tab_widget.tabText(0) == "Brief Synopsis"
    assert page.tab_widget.tabText(1) == "Chronology Rows"
    assert page.synopsis_panel.count() == 2
    assert page.table_panel.count() == 2


def test_clicking_synopsis_text_toggles_entry_selection(qtbot, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    source = _build_chronology_docx(tmp_path / "chronology.docx")
    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)
    item = page.synopsis_panel.item(0)
    position = page.synopsis_panel.visualItemRect(item).center()

    qtbot.mouseClick(page.synopsis_panel.viewport(), Qt.MouseButton.LeftButton, pos=position)

    assert item.checkState() == Qt.CheckState.Checked
    assert page.is_row_checked(page.document.rows[0].id)

    qtbot.mouseClick(page.synopsis_panel.viewport(), Qt.MouseButton.LeftButton, pos=position)

    assert item.checkState() == Qt.CheckState.Unchecked
    assert not page.is_row_checked(page.document.rows[0].id)


def test_clicking_chronology_row_text_toggles_row_selection(qtbot, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    source = _build_chronology_docx(tmp_path / "chronology.docx")
    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)
    row_id = page.document.rows[0].id
    position = page.table_panel.visualRect(page.table_panel.model().index(0, 4)).center()

    qtbot.mouseClick(page.table_panel.viewport(), Qt.MouseButton.LeftButton, pos=position)

    assert page.is_row_checked(row_id)
    assert page.selected_count_label.text() == "1 row selected"

    qtbot.mouseClick(page.table_panel.viewport(), Qt.MouseButton.LeftButton, pos=position)

    assert not page.is_row_checked(row_id)
    assert page.selected_count_label.text() == "0 rows selected"


def test_chronology_row_tab_displays_all_source_columns_verbatim(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    source = _build_chronology_docx(tmp_path / "chronology.docx")
    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)

    assert page.table_panel.columnCount() == 6
    assert page.table_panel.horizontalHeaderItem(1).text() == "DATE"
    assert page.table_panel.horizontalHeaderItem(2).text() == "PAGE NO"
    assert page.table_panel.horizontalHeaderItem(3).text() == "PROVIDER"
    assert page.table_panel.horizontalHeaderItem(4).text() == "DESCRIPTION"
    assert page.table_panel.horizontalHeaderItem(5).text() == "Red Flags/Comments"
    assert page.table_panel.item(0, 2).text() == (
        "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\n"
        "Pg No: 599-604/2530"
    )
    assert page.table_panel.item(0, 3).text() == "Kaiser Permanente\nHenry Louis Marr, DO"
    assert page.table_panel.item(0, 4).text() == "EMERGENCY DEPARTMENT NOTE\nChief Complaint: Ankle injury."
```

- [ ] **Step 2: Verify subtab/click tests fail**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -k "subtabs or toggles or source_columns_verbatim" -q
```

Expected: FAIL because `tab_widget` does not exist and row click/text display behavior is still the old side-by-side table.

- [ ] **Step 3: Implement single-click toggles in `BriefSynopsisPanel`**

In `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`, add `QMouseEvent` to the typing imports if needed and add this method to `BriefSynopsisPanel`:

```python
    def mousePressEvent(self, event) -> None:
        item = self.itemAt(event.pos())
        if item is None:
            super().mousePressEvent(event)
            return
        next_state = (
            Qt.CheckState.Unchecked
            if item.checkState() == Qt.CheckState.Checked
            else Qt.CheckState.Checked
        )
        item.setCheckState(next_state)
        event.accept()
```

Keep the existing `itemChanged` signal path so checkbox state changes still call `_on_item_changed()`.

- [ ] **Step 4: Implement row-wide toggles and source columns in `ChronologyTablePanel`**

Change the table setup from 5 columns to 6 columns:

```python
self.setColumnCount(6)
self.setHorizontalHeaderLabels([
    "Select",
    "DATE",
    "PAGE NO",
    "PROVIDER",
    "DESCRIPTION",
    "Red Flags/Comments",
])
self.setWordWrap(True)
self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
```

Update `load_rows()` so it writes verbatim source fields:

```python
self.setItem(index, 0, self._check_item(row))
self.setItem(index, 1, self._text_item(row.date))
self.setItem(index, 2, self._text_item(row.page_no))
self.setItem(index, 3, self._text_item(row.provider))
self.setItem(index, 4, self._text_item(row.description))
self.setItem(index, 5, self._text_item(row.flags))
```

Add row-wide click toggling:

```python
    def mousePressEvent(self, event) -> None:
        row_index = self.rowAt(event.pos().y())
        if row_index < 0:
            super().mousePressEvent(event)
            return
        item = self.item(row_index, 0)
        if item is None:
            super().mousePressEvent(event)
            return
        row_id = item.data(Qt.ItemDataRole.UserRole)
        if not self._extractable_by_id.get(row_id, False):
            event.accept()
            return
        self.set_row_checked(row_id, not self.is_row_checked(row_id))
        event.accept()
```

Update tooltip application loops to use `range(self.columnCount())`, which already supports the new sixth column.

- [ ] **Step 5: Replace splitter with tabs in `MedChronologySelectionPage._setup_ui()`**

Replace the splitter block:

```python
splitter = QSplitter(Qt.Orientation.Horizontal)
splitter.addWidget(self._build_synopsis_pane())
splitter.addWidget(self._build_table_pane())
splitter.setSizes([360, 760])
layout.addWidget(splitter, 1)
```

with:

```python
self.tab_widget = QTabWidget()
self.tab_widget.addTab(self._build_synopsis_pane(), "Brief Synopsis")
self.tab_widget.addTab(self._build_table_pane(), "Chronology Rows")
layout.addWidget(self.tab_widget, 1)
```

Add `QTabWidget` to the imports and remove `QSplitter` if it is no longer used.

- [ ] **Step 6: Run subtab/click tests**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -k "subtabs or toggles or source_columns_verbatim" -q
```

Expected: all selected tests pass.

- [ ] **Step 7: Run existing viewer tests**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -q
```

Expected: all viewer tests pass. If any tests expected old `Pages` display text or the side-by-side layout, update the assertions to the approved subtab/verbatim design.

- [ ] **Step 8: Commit subtab UI change**

Run:

```powershell
git add -- icharlotte_core\ui\wizard\pages\med_record_extractor_page.py tests\test_wizard\test_med_record_extractor_viewer.py
git diff --cached --check
git commit -m "feat(wizard): add med extractor subtabs"
```

Expected: commit includes only the viewer page and viewer tests.

---

## Task 3: Persist Global Chronology Table Column Widths

**Files:**
- Modify: `tests/test_wizard/test_med_record_extractor_viewer.py`
- Modify: `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`

- [ ] **Step 1: Add QSettings isolation fixture for viewer tests**

Near the top of `tests/test_wizard/test_med_record_extractor_viewer.py`, add:

```python
from PySide6.QtCore import QSettings
```

Add this fixture after the imports:

```python
@pytest.fixture(autouse=True)
def isolated_qsettings(tmp_path):
    previous_default_format = QSettings.defaultFormat()
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    yield
    QSettings("iCharlotte", "iCharlotte").sync()
    QSettings.setDefaultFormat(previous_default_format)
```

- [ ] **Step 2: Write failing column-width and row-height tests**

Append these tests to `tests/test_wizard/test_med_record_extractor_viewer.py`:

```python
def test_chronology_column_widths_persist_globally(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    source = _build_chronology_docx(tmp_path / "chronology.docx")
    first = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(first)
    first.table_panel.setColumnWidth(4, 640)
    first.table_panel.save_column_widths()

    second = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(second)

    assert second.table_panel.columnWidth(4) == 640


def test_chronology_row_heights_auto_fit_wrapped_text(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import (
        MedChronologySelectionPage,
    )

    source = _build_chronology_docx(
        tmp_path / "chronology.docx",
        first_page_no="Record\n\nPg No: 1/2",
    )
    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)
    page.table_panel.setColumnWidth(4, 120)
    page.table_panel.resizeRowsToContents()

    assert page.table_panel.rowHeight(0) > page.table_panel.verticalHeader().defaultSectionSize()
```

- [ ] **Step 3: Verify settings tests fail**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -k "column_widths_persist or row_heights_auto_fit" -q
```

Expected: FAIL because `save_column_widths()` does not exist and column widths are not restored.

- [ ] **Step 4: Implement column-width persistence**

In `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`, add `QSettings` to imports:

```python
from PySide6.QtCore import QSettings, Qt, Signal
```

Add constants near the imports:

```python
TABLE_COLUMN_WIDTHS_KEY = "wizard/med_record_extractor/chronology_table_column_widths"
DEFAULT_TABLE_COLUMN_WIDTHS = [56, 96, 240, 260, 520, 180]
```

Update `ChronologyTablePanel.__init__()`:

```python
self.horizontalHeader().setStretchLastSection(False)
for column in range(self.columnCount()):
    self.horizontalHeader().setSectionResizeMode(column, QHeaderView.ResizeMode.Interactive)
self.restore_column_widths()
self.horizontalHeader().sectionResized.connect(self._on_section_resized)
```

Add these methods to `ChronologyTablePanel`:

```python
    def restore_column_widths(self) -> None:
        widths = _coerce_column_widths(
            QSettings("iCharlotte", "iCharlotte").value(TABLE_COLUMN_WIDTHS_KEY),
            self.columnCount(),
        )
        if widths is None:
            widths = DEFAULT_TABLE_COLUMN_WIDTHS[: self.columnCount()]
        for column, width in enumerate(widths):
            self.setColumnWidth(column, width)
        self.resizeRowsToContents()

    def save_column_widths(self) -> None:
        widths = [self.columnWidth(column) for column in range(self.columnCount())]
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue(TABLE_COLUMN_WIDTHS_KEY, widths)
        settings.sync()

    def _on_section_resized(self, logical_index: int, old_size: int, new_size: int) -> None:
        self.save_column_widths()
        self.resizeRowsToContents()
```

Add this module helper near `_saved_sources()`:

```python
def _coerce_column_widths(value, expected_count: int) -> list[int] | None:
    if value is None:
        return None
    if isinstance(value, str):
        parts = [part.strip() for part in value.split(",")]
    elif isinstance(value, (list, tuple)):
        parts = list(value)
    else:
        return None
    widths: list[int] = []
    for part in parts:
        try:
            width = int(part)
        except (TypeError, ValueError):
            return None
        if width < 24:
            return None
        widths.append(width)
    if len(widths) != expected_count:
        return None
    return widths
```

In `load_rows()`, keep `self.resizeRowsToContents()` after rows are populated.

- [ ] **Step 5: Run column-width tests**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -k "column_widths_persist or row_heights_auto_fit" -q
```

Expected: both tests pass.

- [ ] **Step 6: Run full viewer tests**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -q
```

Expected: all viewer tests pass.

- [ ] **Step 7: Commit persistence change**

Run:

```powershell
git add -- icharlotte_core\ui\wizard\pages\med_record_extractor_page.py tests\test_wizard\test_med_record_extractor_viewer.py
git diff --cached --check
git commit -m "feat(wizard): persist med extractor table widths"
```

Expected: commit includes only the viewer page and viewer tests.

---

## Task 4: Regression Verification and Live Chronology Smoke Test

**Files:**
- Modify only if failures expose a required fix:
  - `icharlotte_core/med_record_chronology.py`
  - `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`
  - `tests/test_med_record_extractor.py`
  - `tests/test_wizard/test_med_record_extractor_viewer.py`

- [ ] **Step 1: Run focused parser/viewer/in-process tests**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_med_record_extractor.py tests\test_wizard\test_med_record_extractor_viewer.py tests\test_wizard\test_in_process_task_tab.py -q
```

Expected: all tests pass.

- [ ] **Step 2: Run full wizard tests**

Run:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_wizard -q
```

Expected: all wizard tests pass. Existing skipped tests may remain skipped.

- [ ] **Step 3: Run a no-write live chronology viewer smoke test**

Run this script to parse the latest 5800.013 chronology and instantiate the viewer offscreen without extraction:

```powershell
$env:QT_QPA_PLATFORM='offscreen'; @'
from PySide6.QtWidgets import QApplication
from icharlotte_core.med_record_chronology import parse_chronology_document, match_synopsis_to_rows
from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

case_path = r"Z:\Shared\Current Clients\5800 - AMTRUST\013 - Hall"
chron_path = r"Z:\Shared\Current Clients\5800 - AMTRUST\013 - Hall\RECORDS\Medical Summary - DO NOT PRODUCE\BS FILE NO. 5800.013 2025-09-18 Medical Summary of Rhonda L Hall (Updated Version 2).docx"

doc = parse_chronology_document(chron_path)
app = QApplication.instance() or QApplication([])
page = MedChronologySelectionPage(case_path, "5800.013", chron_path)
confident = sum(
    1
    for paragraph in doc.synopsis_paragraphs
    if match_synopsis_to_rows(paragraph, doc.rows).status == "confident"
)

print(f"blocking_errors={doc.blocking_errors}")
print(f"synopsis={len(doc.synopsis_paragraphs)}")
print(f"rows={len(doc.rows)}")
print(f"confident={confident}")
print(f"tabs={page.tab_widget.count()}:{page.tab_widget.tabText(0)}|{page.tab_widget.tabText(1)}")
print(f"table_columns={page.table_panel.columnCount()}")
print(f"extract_enabled={page.extract_btn.isEnabled()}")
'@ | .venv\Scripts\python.exe -
```

Expected output includes:

```text
blocking_errors=[]
tabs=2:Brief Synopsis|Chronology Rows
table_columns=6
extract_enabled=False
```

- [ ] **Step 4: Check git staged state**

Run:

```powershell
git status --short
git diff --cached --name-only
```

Expected: no staged files. Any remaining dirty files should be unrelated pre-existing work unless Task 4 uncovered a required fix.

- [ ] **Step 5: Commit any regression fixes**

If Task 4 required fixes, run:

```powershell
git add -- icharlotte_core\med_record_chronology.py icharlotte_core\ui\wizard\pages\med_record_extractor_page.py tests\test_med_record_extractor.py tests\test_wizard\test_med_record_extractor_viewer.py
git diff --cached --check
git commit -m "fix(wizard): harden med extractor subtab viewer"
```

Expected: if no fixes were needed, skip this step and do not create an empty commit.

---

## Self-Review

Spec coverage:

- Two subtabs are covered by Task 2.
- Brief Synopsis verbatim display is covered by Task 1 and Task 2 tests.
- Single-click synopsis toggling is covered by Task 2.
- Chronology row verbatim source columns are covered by Task 1 and Task 2 tests.
- Single-click chronology row toggling is covered by Task 2.
- Auto-fit row heights and global column-width persistence are covered by Task 3.
- Existing source-aware selection restore is preserved by Task 4 focused test runs.

Placeholder scan:

- No placeholder markers or incomplete implementation steps are intentionally present.
- Every code-changing task includes concrete tests, implementation snippets, commands, and expected outcomes.

Type consistency:

- `BriefSynopsisPanel` and `ChronologyTablePanel` remain the UI classes under `MedChronologySelectionPage`.
- `QSettings("iCharlotte", "iCharlotte")` matches existing app preference storage patterns.
- The table now has six columns: checkbox plus five source chronology columns.
