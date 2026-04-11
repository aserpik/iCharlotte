# Quote Insertion for Mediation Brief — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a post-generation workflow for inserting deposition quotes into an already-generated mediation brief — users search transcripts via LLM, select matching Q&A passages, and insert them (Quick Insert or Weave In) into specific sections.

**Architecture:** A new `QuoteInsertionDialog` (PyQt6 modal) provides the UI for transcript upload, search, and quote selection. The `MediationBriefGenerator` class gains `search_quotes()` and `insert_quotes_quick()` methods. The LLM search runs in a background `QuoteSearchWorker(QThread)`. Weave In mode reuses the existing `generate_section()` + `RefinementWorker` infrastructure.

**Tech Stack:** Python 3.x, PySide6, PyMuPDF (fitz), python-docx, LLMCaller for LLM calls.

**Spec:** `docs/superpowers/specs/2026-04-10-quote-insertion-design.md`

---

## File Structure

| File | Responsibility |
|------|----------------|
| **Create:** `icharlotte_core/ui/quote_dialog.py` | `QuoteInsertionDialog` (modal dialog UI), `QuoteSearchWorker(QThread)` |
| **Create:** `Scripts/prompts/mediation_brief/quote_search_current.txt` | Editable quote search prompt |
| **Create:** `tests/test_quote_insertion.py` | Unit tests for search parsing, insertion logic |
| **Modify:** `icharlotte_core/mediation_brief.py` | Add `search_quotes()`, `insert_quotes_quick()`, `self.saved_path` attribute |
| **Modify:** `icharlotte_core/ui/tabs.py` | Add "Add Quotes" button, handler, post-insertion save logic |
| **Modify:** `Scripts/prompts/registry.json` | Register `mediation_brief:quote_search` prompt |

---

### Task 1: Create Quote Search Prompt and Register

**Files:**
- Create: `Scripts/prompts/mediation_brief/quote_search_current.txt`
- Modify: `Scripts/prompts/registry.json`

- [ ] **Step 1: Create the quote search prompt**

Write `Scripts/prompts/mediation_brief/quote_search_current.txt`:

```text
You are a legal transcript analyst. Your job is to find specific Q&A passages in deposition transcripts that match the user's description.

CRITICAL RULES:
- Return the testimony EXACTLY as it appears in the transcript — do NOT paraphrase, reword, summarize, or clean up the text in any way
- Preserve all original wording, punctuation, dashes, incomplete sentences, and verbal tics exactly as written
- The only acceptable change is removing line numbers that appear at the start of transcript lines
- If you cannot find an exact match, say so — do NOT fabricate or approximate testimony

For each matching passage found, return it in this exact format:

QUOTE_RESULT_START
DEPONENT: [Last name of the deponent]
SOURCE: [Filename of the transcript]
PAGE_LINE: [StartPage:StartLine-EndPage:EndLine or StartPage:StartLine-EndLine]
RELEVANCE: [One sentence explaining why this passage matches the search]
Q. [Exact question text]
A. [Exact answer text]
[Continue with additional Q. and A. lines if the exchange spans multiple questions]
QUOTE_RESULT_END

If multiple relevant passages are found, include multiple QUOTE_RESULT_START/END blocks.
If no matching testimony is found, respond with: NO_MATCHES_FOUND
```

- [ ] **Step 2: Register in registry.json**

Read `Scripts/prompts/registry.json` and add this entry under the `"prompts"` key:

```json
"mediation_brief:quote_search": {
    "agent": "mediation_brief",
    "pass_name": "quote_search",
    "current_version": "v1",
    "versions": [
        {
            "version": "v1",
            "created": "2026-04-10T18:00:00.000000",
            "description": "Initial quote search prompt",
            "author": "",
            "is_current": true,
            "performance_score": 0.0,
            "usage_count": 0
        }
    ]
}
```

- [ ] **Step 3: Commit**

```bash
git add Scripts/prompts/mediation_brief/quote_search_current.txt Scripts/prompts/registry.json
git commit -m "feat(mediation-brief): add quote search prompt and register in workbench"
```

---

### Task 2: Quote Search and Insertion Logic in MediationBriefGenerator

**Files:**
- Modify: `icharlotte_core/mediation_brief.py`
- Create: `tests/test_quote_insertion.py`

- [ ] **Step 1: Write tests for quote search result parsing**

Create `tests/test_quote_insertion.py`:

```python
"""Tests for quote insertion feature."""
import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock


class TestQuoteResultParsing(unittest.TestCase):
    """Test parsing of LLM quote search results."""

    def test_parse_single_quote_result(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        llm_output = (
            "QUOTE_RESULT_START\n"
            "DEPONENT: Haydel\n"
            "SOURCE: 25.05.29 Depo of Benjamin Haydel.pdf\n"
            "PAGE_LINE: 35:4-8\n"
            "RELEVANCE: Plaintiff admits seeing the plastic sheeting\n"
            "Q. Did you see the plastic on the ground before you fell?\n"
            "A. Yeah, I saw it -- I saw it coming up the stairs.\n"
            "QUOTE_RESULT_END"
        )
        results = gen._parse_quote_results(llm_output)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["deponent"], "Haydel")
        self.assertEqual(results[0]["source"], "25.05.29 Depo of Benjamin Haydel.pdf")
        self.assertEqual(results[0]["page_line"], "35:4-8")
        self.assertIn("Q. Did you see", results[0]["qa_text"])
        self.assertIn("A. Yeah, I saw it", results[0]["qa_text"])

    def test_parse_multiple_quote_results(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        llm_output = (
            "QUOTE_RESULT_START\n"
            "DEPONENT: Haydel\n"
            "SOURCE: depo1.pdf\n"
            "PAGE_LINE: 35:4-8\n"
            "RELEVANCE: First relevant passage\n"
            "Q. First question?\n"
            "A. First answer.\n"
            "QUOTE_RESULT_END\n\n"
            "QUOTE_RESULT_START\n"
            "DEPONENT: Smith\n"
            "SOURCE: depo2.pdf\n"
            "PAGE_LINE: 12:1-5\n"
            "RELEVANCE: Second relevant passage\n"
            "Q. Second question?\n"
            "A. Second answer.\n"
            "QUOTE_RESULT_END"
        )
        results = gen._parse_quote_results(llm_output)
        self.assertEqual(len(results), 2)
        self.assertEqual(results[0]["deponent"], "Haydel")
        self.assertEqual(results[1]["deponent"], "Smith")

    def test_parse_no_matches(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        results = gen._parse_quote_results("NO_MATCHES_FOUND")
        self.assertEqual(results, [])

    def test_parse_empty_response(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        results = gen._parse_quote_results("")
        self.assertEqual(results, [])


class TestQuoteInsertion(unittest.TestCase):
    """Test quick insert of quotes into section text."""

    def test_insert_quote_at_end_of_section(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.sections = {
            "liability": "Existing liability argument text.\n\nSUBSECTION: No Duty\nDuty argument here."
        }
        quote = {
            "deponent": "Haydel",
            "source": "depo.pdf",
            "page_line": "35:4-8",
            "qa_text": "Q. Did you see it?\nA. Yes, I saw it.",
            "relevance": "Admits seeing hazard",
        }
        gen.insert_quotes_quick([quote], "liability", None)
        updated = gen.sections["liability"]
        self.assertIn("DEPO_QUOTE_START", updated)
        self.assertIn("Q. Did you see it?", updated)
        self.assertIn("(Haydel Depo Trns., at p. 35:4-8.)", updated)

    def test_insert_quote_into_specific_subsection(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.sections = {
            "liability": (
                "Intro paragraph.\n\n"
                "SUBSECTION: No Duty\n"
                "Duty argument here.\n\n"
                "SUBSECTION: No Breach\n"
                "Breach argument here."
            )
        }
        quote = {
            "deponent": "Haydel",
            "source": "depo.pdf",
            "page_line": "35:4-8",
            "qa_text": "Q. Did you see it?\nA. Yes.",
            "relevance": "Relevant to duty",
        }
        gen.insert_quotes_quick([quote], "liability", "No Duty")
        updated = gen.sections["liability"]
        # Quote should appear between "Duty argument" and "SUBSECTION: No Breach"
        duty_pos = updated.index("Duty argument")
        quote_pos = updated.index("DEPO_QUOTE_START")
        breach_pos = updated.index("SUBSECTION: No Breach")
        self.assertGreater(quote_pos, duty_pos)
        self.assertLess(quote_pos, breach_pos)


class TestSavedPath(unittest.TestCase):
    """Test saved_path attribute for overwrite behavior."""

    def test_saved_path_initialized_to_none(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        self.assertIsNone(gen.saved_path)

    def test_saved_path_cleared_on_reset(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.saved_path = "/some/path.docx"
        gen.reset()
        self.assertIsNone(gen.saved_path)


if __name__ == '__main__':
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_quote_insertion.py -v
```

Expected: FAIL — `_parse_quote_results`, `insert_quotes_quick`, and `saved_path` don't exist.

- [ ] **Step 3: Add `saved_path` attribute to `__init__` and `reset()`**

In `icharlotte_core/mediation_brief.py`, in `__init__()` (~line 201), add after `self.is_active`:

```python
        self.saved_path: Optional[str] = None  # Last saved file path for overwrite
```

In `reset()` (~line 1056), add `self.saved_path = None` alongside the other attribute resets.

- [ ] **Step 4: Implement `_parse_quote_results()`**

Add this method to `MediationBriefGenerator`:

```python
    # ------------------------------------------------------------------
    # Quote search and insertion
    # ------------------------------------------------------------------

    _QUOTE_RESULT_RE = re.compile(
        r'QUOTE_RESULT_START\s*\n'
        r'DEPONENT:\s*(.+)\n'
        r'SOURCE:\s*(.+)\n'
        r'PAGE_LINE:\s*(.+)\n'
        r'RELEVANCE:\s*(.+)\n'
        r'((?:(?:Q\.|A\.).*\n?)+)'
        r'QUOTE_RESULT_END',
        re.MULTILINE
    )

    def _parse_quote_results(self, llm_output: str) -> List[Dict]:
        """Parse LLM quote search output into structured results.

        Returns list of dicts with keys: deponent, source, page_line,
        relevance, qa_text.
        """
        if not llm_output or "NO_MATCHES_FOUND" in llm_output:
            return []

        results = []
        for match in self._QUOTE_RESULT_RE.finditer(llm_output):
            results.append({
                "deponent": match.group(1).strip(),
                "source": match.group(2).strip(),
                "page_line": match.group(3).strip(),
                "relevance": match.group(4).strip(),
                "qa_text": match.group(5).strip(),
            })
        return results
```

- [ ] **Step 5: Implement `search_quotes()`**

Add this method to `MediationBriefGenerator`:

```python
    def search_quotes(self, transcript_paths: List[str], description: str) -> List[Dict]:
        """Search deposition transcripts for Q&A passages matching description.

        Reads all transcripts, sends them to the LLM with the search prompt,
        and returns parsed quote results.

        Args:
            transcript_paths: List of file paths to deposition transcripts (PDF/DOCX)
            description: What testimony the user is looking for

        Returns:
            List of quote result dicts (see _parse_quote_results)
        """
        # Read transcripts
        transcript_text = ""
        for path in transcript_paths:
            fname = os.path.basename(path)
            ext = os.path.splitext(path)[1].lower()
            transcript_text += f"\n--- TRANSCRIPT: {fname} ---\n"
            if ext == ".pdf":
                try:
                    doc = fitz.open(path)
                    for page in doc:
                        transcript_text += page.get_text() + "\n"
                    doc.close()
                except Exception as e:
                    log.warning("Failed to read transcript %s: %s", fname, e)
                    transcript_text += f"[Error reading file: {e}]\n"
            elif ext == ".docx":
                try:
                    doc = DocxDocument(path)
                    for para in doc.paragraphs:
                        transcript_text += para.text + "\n"
                except Exception as e:
                    log.warning("Failed to read transcript %s: %s", fname, e)
                    transcript_text += f"[Error reading file: {e}]\n"
            transcript_text += "\n"

        if not transcript_text.strip():
            return []

        # Build prompt
        search_prompt = self._get_section_prompt("quote_search")
        full_prompt = f"{search_prompt}\n\nUSER'S SEARCH DESCRIPTION:\n{description}"

        system = (
            "You are a legal transcript analyst. Find Q&A passages in deposition "
            "transcripts that match the user's description. Return testimony EXACTLY "
            "as it appears — do NOT paraphrase, reword, or change any text."
        )

        caller = LLMCaller()
        result = caller.call(
            prompt=full_prompt,
            text=transcript_text,
            agent_id="agent_mediation_brief",
        )
        if not result:
            return []

        return self._parse_quote_results(result)
```

- [ ] **Step 6: Implement `insert_quotes_quick()`**

Add this method to `MediationBriefGenerator`:

```python
    def insert_quotes_quick(
        self,
        quotes: List[Dict],
        section_name: str,
        subsection_title: Optional[str] = None,
    ) -> None:
        """Insert quote blocks into a section's text (Quick Insert mode).

        Appends formatted DEPO_QUOTE_START/END blocks to the section text,
        either at the end of a specific subsection or at the end of the section.

        Args:
            quotes: List of quote dicts from _parse_quote_results
            section_name: Target section (e.g., "liability")
            subsection_title: If provided, insert after this subsection.
                If None, append at the end of the section.
        """
        if section_name not in self.sections:
            return

        # Build the quote blocks
        quote_blocks = []
        for q in quotes:
            block = (
                f"\n\nDEPO_QUOTE_START\n"
                f"{q['qa_text']}\n"
                f"DEPO_QUOTE_END\n"
                f"({q['deponent']} Depo Trns., at p. {q['page_line']}.)"
            )
            quote_blocks.append(block)

        insert_text = "".join(quote_blocks)

        section_text = self.sections[section_name]

        if subsection_title:
            # Find the subsection and insert after its content
            # Look for the next SUBSECTION: marker after this one
            pattern = re.compile(
                rf'^SUBSECTION:\s*{re.escape(subsection_title)}\s*$',
                re.MULTILINE
            )
            match = pattern.search(section_text)
            if match:
                # Find the next SUBSECTION: or end of text
                next_sub = self._SUBSECTION_RE.search(section_text, match.end())
                if next_sub:
                    insert_pos = next_sub.start()
                    # Insert before the next subsection
                    section_text = (
                        section_text[:insert_pos].rstrip()
                        + insert_text + "\n\n"
                        + section_text[insert_pos:]
                    )
                else:
                    # Last subsection — append at end
                    section_text = section_text.rstrip() + insert_text
            else:
                # Subsection not found — append at end
                section_text = section_text.rstrip() + insert_text
        else:
            # Append at end of section
            section_text = section_text.rstrip() + insert_text

        self.sections[section_name] = section_text
```

- [ ] **Step 7: Run tests to verify they pass**

```bash
python -m pytest tests/test_quote_insertion.py -v
```

Expected: All tests PASS.

- [ ] **Step 8: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_quote_insertion.py
git commit -m "feat(mediation-brief): quote search, result parsing, and quick insertion logic"
```

---

### Task 3: Quote Insertion Dialog UI

**Files:**
- Create: `icharlotte_core/ui/quote_dialog.py`

- [ ] **Step 1: Create the dialog**

Create `icharlotte_core/ui/quote_dialog.py`:

```python
"""Quote Insertion Dialog for Mediation Brief.

Provides a modal dialog for searching deposition transcripts and inserting
selected Q&A passages into the brief.
"""

import os
from typing import List, Dict, Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QPlainTextEdit, QComboBox, QRadioButton, QButtonGroup,
    QListWidget, QListWidgetItem, QScrollArea, QWidget,
    QCheckBox, QFileDialog, QGroupBox, QProgressBar,
    QMessageBox, QSizePolicy, QFrame, QTextEdit,
)
from PySide6.QtCore import Qt, Signal, QThread

from ..mediation_brief import (
    MediationBriefGenerator, SECTION_ORDER, SECTION_HEADINGS,
)

import logging
log = logging.getLogger(__name__)


class QuoteSearchWorker(QThread):
    """Background worker for LLM quote search."""
    results_ready = Signal(list)  # List of quote result dicts
    error = Signal(str)

    def __init__(self, generator, transcript_paths, description, parent=None):
        super().__init__(parent)
        self.generator = generator
        self.transcript_paths = transcript_paths
        self.description = description

    def run(self):
        try:
            results = self.generator.search_quotes(
                self.transcript_paths, self.description
            )
            self.results_ready.emit(results)
        except Exception as e:
            log.exception("Quote search failed")
            self.error.emit(str(e))


class QuoteResultWidget(QFrame):
    """A single quote result card with checkbox, text, and edit button."""

    def __init__(self, quote_data: Dict, index: int, parent=None):
        super().__init__(parent)
        self.quote_data = quote_data
        self.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        self.setLineWidth(1)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        # Top row: checkbox + deponent info
        top_row = QHBoxLayout()
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        top_row.addWidget(self.checkbox)

        info_label = QLabel(
            f"<b>{quote_data['deponent']}</b> — {quote_data['source']} "
            f"(p. {quote_data['page_line']})"
        )
        info_label.setWordWrap(True)
        top_row.addWidget(info_label, 1)

        edit_btn = QPushButton("Edit")
        edit_btn.setFixedWidth(50)
        edit_btn.clicked.connect(self._toggle_edit)
        top_row.addWidget(edit_btn)
        layout.addLayout(top_row)

        # Relevance note
        relevance = QLabel(f"<i>{quote_data['relevance']}</i>")
        relevance.setWordWrap(True)
        layout.addWidget(relevance)

        # Q&A text display
        self.qa_display = QLabel(quote_data["qa_text"])
        self.qa_display.setWordWrap(True)
        self.qa_display.setStyleSheet("padding: 8px; background: #f5f5f5;")
        layout.addWidget(self.qa_display)

        # Editable text area (hidden by default)
        self.qa_editor = QPlainTextEdit(quote_data["qa_text"])
        self.qa_editor.setMaximumHeight(150)
        self.qa_editor.hide()
        layout.addWidget(self.qa_editor)

        self._editing = False

    def _toggle_edit(self):
        if self._editing:
            # Save edits
            self.quote_data["qa_text"] = self.qa_editor.toPlainText()
            self.qa_display.setText(self.quote_data["qa_text"])
            self.qa_display.show()
            self.qa_editor.hide()
        else:
            self.qa_editor.setPlainText(self.quote_data["qa_text"])
            self.qa_display.hide()
            self.qa_editor.show()
        self._editing = not self._editing

    def is_selected(self) -> bool:
        return self.checkbox.isChecked()

    def get_quote_data(self) -> Dict:
        return self.quote_data


class QuoteInsertionDialog(QDialog):
    """Modal dialog for searching transcripts and inserting quotes."""

    quotes_to_insert = Signal(list, str, str, str)  # quotes, section, subsection, mode

    def __init__(self, generator: MediationBriefGenerator, parent=None):
        super().__init__(parent)
        self.generator = generator
        self._worker = None
        self._result_widgets: List[QuoteResultWidget] = []

        self.setWindowTitle("Add Deposition Quotes")
        self.setMinimumSize(700, 600)
        self.resize(800, 700)

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # 1. Transcript upload area
        transcript_group = QGroupBox("Transcripts")
        transcript_layout = QVBoxLayout(transcript_group)

        self.transcript_list = QListWidget()
        self.transcript_list.setMaximumHeight(100)
        transcript_layout.addWidget(self.transcript_list)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Transcript(s)")
        add_btn.clicked.connect(self._add_transcripts)
        btn_row.addWidget(add_btn)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_transcript)
        btn_row.addWidget(remove_btn)
        btn_row.addStretch()
        transcript_layout.addLayout(btn_row)

        layout.addWidget(transcript_group)

        # 2. Search description
        search_group = QGroupBox("Search Description")
        search_layout = QVBoxLayout(search_group)
        self.search_input = QPlainTextEdit()
        self.search_input.setPlaceholderText(
            "Describe what testimony you're looking for (e.g., "
            "'where plaintiff admits he saw the plastic sheeting before the fall')"
        )
        self.search_input.setMaximumHeight(80)
        search_layout.addWidget(self.search_input)
        layout.addWidget(search_group)

        # 3. Placement controls
        placement_group = QGroupBox("Placement")
        placement_layout = QHBoxLayout(placement_group)

        placement_layout.addWidget(QLabel("Section:"))
        self.section_combo = QComboBox()
        for name in SECTION_ORDER:
            if name in ("introduction", "procedural_status", "settlement_position", "conclusion"):
                continue  # Quotes typically go in liability/damages/facts
            roman, title = SECTION_HEADINGS[name]
            self.section_combo.addItem(f"{roman}. {title}", name)
        self.section_combo.currentIndexChanged.connect(self._update_subsections)
        placement_layout.addWidget(self.section_combo)

        placement_layout.addWidget(QLabel("Subsection:"))
        self.subsection_combo = QComboBox()
        self.subsection_combo.addItem("Auto (LLM decides)", "auto")
        placement_layout.addWidget(self.subsection_combo)

        placement_layout.addSpacing(20)
        self.mode_group = QButtonGroup(self)
        self.quick_radio = QRadioButton("Quick Insert")
        self.quick_radio.setChecked(True)
        self.weave_radio = QRadioButton("Weave In")
        self.mode_group.addButton(self.quick_radio)
        self.mode_group.addButton(self.weave_radio)
        placement_layout.addWidget(self.quick_radio)
        placement_layout.addWidget(self.weave_radio)

        layout.addWidget(placement_group)

        # 4. Search button + progress
        search_row = QHBoxLayout()
        self.search_btn = QPushButton("Search")
        self.search_btn.clicked.connect(self._run_search)
        search_row.addWidget(self.search_btn)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # Indeterminate
        self.progress.hide()
        search_row.addWidget(self.progress)
        search_row.addStretch()
        layout.addLayout(search_row)

        # 5. Results panel
        results_group = QGroupBox("Results")
        results_layout = QVBoxLayout(results_group)

        self.results_scroll = QScrollArea()
        self.results_scroll.setWidgetResizable(True)
        self.results_container = QWidget()
        self.results_container_layout = QVBoxLayout(self.results_container)
        self.results_container_layout.addStretch()
        self.results_scroll.setWidget(self.results_container)
        results_layout.addWidget(self.results_scroll)

        self.no_results_label = QLabel("")
        self.no_results_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        results_layout.addWidget(self.no_results_label)

        layout.addWidget(results_group, 1)

        # 6. Bottom buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.insert_btn = QPushButton("Insert Selected")
        self.insert_btn.setEnabled(False)
        self.insert_btn.clicked.connect(self._insert_selected)
        btn_layout.addWidget(self.insert_btn)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def _add_transcripts(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Transcript(s)", "",
            "Documents (*.pdf *.docx)"
        )
        for f in files:
            # Avoid duplicates
            existing = [
                self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.transcript_list.count())
            ]
            if f not in existing:
                item = QListWidgetItem(os.path.basename(f))
                item.setData(Qt.ItemDataRole.UserRole, f)
                self.transcript_list.addItem(item)

    def _remove_transcript(self):
        for item in self.transcript_list.selectedItems():
            self.transcript_list.takeItem(self.transcript_list.row(item))

    def _update_subsections(self):
        """Populate subsection dropdown from the selected section's content."""
        self.subsection_combo.clear()
        self.subsection_combo.addItem("Auto (LLM decides)", "auto")

        section_name = self.section_combo.currentData()
        if not section_name or section_name not in self.generator.sections:
            return

        import re
        section_text = self.generator.sections[section_name]
        for match in re.finditer(r'^SUBSECTION:\s*(.+)$', section_text, re.MULTILINE):
            title = match.group(1).strip()
            self.subsection_combo.addItem(title, title)

    def _run_search(self):
        if self.transcript_list.count() == 0:
            QMessageBox.warning(self, "No Transcripts", "Please add at least one transcript.")
            return

        description = self.search_input.toPlainText().strip()
        if not description:
            QMessageBox.warning(self, "No Description", "Please describe what testimony to search for.")
            return

        # Collect transcript paths
        paths = []
        for i in range(self.transcript_list.count()):
            paths.append(self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole))

        # Clear previous results
        self._clear_results()
        self.search_btn.setEnabled(False)
        self.progress.show()
        self.no_results_label.setText("Searching...")

        # Run search in background
        self._worker = QuoteSearchWorker(self.generator, paths, description, parent=self)
        self._worker.results_ready.connect(self._on_search_results)
        self._worker.error.connect(self._on_search_error)
        self._worker.start()

    def _clear_results(self):
        self._result_widgets.clear()
        # Remove all widgets from container layout except the stretch
        while self.results_container_layout.count() > 1:
            item = self.results_container_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.insert_btn.setEnabled(False)

    def _on_search_results(self, results: list):
        self.search_btn.setEnabled(True)
        self.progress.hide()

        if not results:
            self.no_results_label.setText(
                "No matching testimony found. Try broadening your search description."
            )
            return

        self.no_results_label.setText(f"Found {len(results)} matching passage(s)")

        for i, quote in enumerate(results):
            widget = QuoteResultWidget(quote, i)
            self._result_widgets.append(widget)
            # Insert before the stretch
            self.results_container_layout.insertWidget(
                self.results_container_layout.count() - 1, widget
            )

        self.insert_btn.setEnabled(True)

    def _on_search_error(self, error_msg: str):
        self.search_btn.setEnabled(True)
        self.progress.hide()
        self.no_results_label.setText(f"Search error: {error_msg}")

    def _insert_selected(self):
        selected = [
            w.get_quote_data() for w in self._result_widgets if w.is_selected()
        ]
        if not selected:
            QMessageBox.warning(self, "No Quotes Selected", "Please select at least one quote.")
            return

        section_name = self.section_combo.currentData()
        subsection = self.subsection_combo.currentData()
        if subsection == "auto":
            subsection = None
        mode = "quick" if self.quick_radio.isChecked() else "weave"

        self.quotes_to_insert.emit(selected, section_name, subsection or "", mode)
        self.accept()
```

- [ ] **Step 2: Verify the file compiles**

```bash
python -c "from icharlotte_core.ui.quote_dialog import QuoteInsertionDialog; print('OK')"
```

Expected: OK

- [ ] **Step 3: Commit**

```bash
git add icharlotte_core/ui/quote_dialog.py
git commit -m "feat(mediation-brief): QuoteInsertionDialog UI and QuoteSearchWorker"
```

---

### Task 4: ChatTab Integration — Add Quotes Button and Handlers

**Files:**
- Modify: `icharlotte_core/ui/tabs.py`

- [ ] **Step 1: Add import**

At the top of `icharlotte_core/ui/tabs.py`, add to the existing mediation brief import block (~line 32):

```python
from ..mediation_brief import (
    MediationBriefGenerator, MediationBriefWorker, RefinementWorker,
    RoutingWorker, SECTION_HEADINGS,
)
from ..ui.quote_dialog import QuoteInsertionDialog
```

Note: the `from ..ui.quote_dialog` import may cause a circular import since tabs.py is in the same package. If so, move the import inside the method that uses it (lazy import in `_on_add_quotes_clicked`).

- [ ] **Step 2: Add "Add Quotes" button to `_on_brief_all_complete()`**

In `_on_brief_all_complete()` (~line 2014), replace the refinement instruction text at the end of the method:

Find this block:
```python
        self.chat_history.append("")
        self.chat_history.append(
            "<i>You can now refine the brief by typing instructions "
            "(e.g., 'make the Damages section more aggressive'). "
            "Or send a normal message to exit brief mode.</i>"
        )
```

Replace with:
```python
        self.chat_history.append("")
        self.chat_history.append(
            "<i>You can now refine the brief by typing instructions "
            "(e.g., 'make the Damages section more aggressive'), "
            "or use the Add Quotes button to insert deposition testimony. "
            "Send a normal message to exit brief mode.</i>"
        )

        # Show "Add Quotes" button
        if not hasattr(self, 'add_quotes_btn') or self.add_quotes_btn is None:
            self.add_quotes_btn = QPushButton("Add Quotes")
            self.add_quotes_btn.clicked.connect(self._on_add_quotes_clicked)
            # Insert button in the layout near the send button
            send_layout = self.send_btn.parent().layout()
            if send_layout:
                send_layout.insertWidget(send_layout.indexOf(self.send_btn), self.add_quotes_btn)
        self.add_quotes_btn.setVisible(True)
        self.add_quotes_btn.setEnabled(True)
```

Also update the `_on_brief_all_complete` method to store the save path. After the `shutil.copy2(temp_output, save_path)` line, add:

```python
                gen.saved_path = save_path
```

- [ ] **Step 3: Add the "Add Quotes" click handler**

Add this method to `ChatTab`:

```python
    def _on_add_quotes_clicked(self):
        """Open the Quote Insertion dialog."""
        from .quote_dialog import QuoteInsertionDialog

        if not self.med_brief_generator or not self.med_brief_generator.is_active:
            return

        dlg = QuoteInsertionDialog(self.med_brief_generator, parent=self)
        dlg.quotes_to_insert.connect(self._on_quotes_confirmed)
        dlg.exec()
```

- [ ] **Step 4: Add the quote confirmation handler**

Add this method to `ChatTab`:

```python
    def _on_quotes_confirmed(self, quotes: list, section_name: str,
                              subsection: str, mode: str):
        """Handle confirmed quote insertion from the dialog."""
        gen = self.med_brief_generator
        count = len(quotes)

        self.chat_history.append(
            f"<b>Inserting {count} quote(s) into "
            f"{SECTION_HEADINGS.get(section_name, ('', section_name))[1]}...</b>"
        )

        if mode == "quick":
            # Quick Insert — synchronous
            sub_title = subsection if subsection else None
            gen.insert_quotes_quick(quotes, section_name, sub_title)
            self.chat_history.append(f"<i>{count} quote(s) inserted.</i>")
            self._reassemble_and_save()
        else:
            # Weave In — regenerate section with LLM
            self.send_btn.setEnabled(False)
            if hasattr(self, 'add_quotes_btn') and self.add_quotes_btn:
                self.add_quotes_btn.setEnabled(False)

            # Build the quote text for the refinement instruction
            quote_text_parts = []
            for q in quotes:
                quote_text_parts.append(
                    f"DEPO_QUOTE_START\n{q['qa_text']}\nDEPO_QUOTE_END\n"
                    f"({q['deponent']} Depo Trns., at p. {q['page_line']}.)"
                )
            all_quotes = "\n\n".join(quote_text_parts)

            instruction = (
                f"Incorporate the following deposition testimony into this section "
                f"at the most appropriate location. Weave it into the argument "
                f"naturally with proper context and transitions. Include the "
                f"testimony verbatim — do not change any wording.\n\n{all_quotes}"
            )

            # Reuse existing refinement infrastructure
            worker = RefinementWorker(
                gen, [section_name], instruction, parent=self
            )
            worker.section_complete.connect(self._on_brief_section_complete)
            worker.all_complete.connect(
                lambda regenerated: self._on_quote_weave_complete()
            )
            worker.error.connect(self._on_brief_error)
            self.med_brief_worker = worker
            worker.start()
```

- [ ] **Step 5: Add the reassemble and save helper**

Add these methods to `ChatTab`:

```python
    def _reassemble_and_save(self):
        """Reassemble the Word document and save (overwrite or Save As)."""
        import tempfile
        gen = self.med_brief_generator

        self.chat_history.append("<i>Assembling document...</i>")

        with tempfile.TemporaryDirectory() as tmpdir:
            temp_output = os.path.join(tmpdir, "mediation_brief.docx")
            try:
                gen.assemble_document(gen.caption_template_path, temp_output)
            except Exception as e:
                self.chat_history.append(f"<b style='color:red'>Assembly error: {e}</b>")
                return

            if gen.saved_path and os.path.exists(os.path.dirname(gen.saved_path)):
                # Overwrite existing file
                shutil.copy2(temp_output, gen.saved_path)
                self.chat_history.append(f"<b>Document updated:</b> {gen.saved_path}")
                self.chat_history.append(
                    "<i><a href='#save_as'>Save As...</a> to save a copy elsewhere.</i>"
                )
            else:
                # No prior save — open Save As dialog
                main_win = self.window()
                case_path = getattr(main_win, 'case_path', None)
                default_dir = os.path.dirname(case_path) if case_path else ""
                default_name = os.path.join(
                    default_dir, "Defendant's Confidential Mediation Brief.docx"
                )
                save_path, _ = QFileDialog.getSaveFileName(
                    self, "Save Mediation Brief", default_name,
                    "Word Documents (*.docx)"
                )
                if save_path:
                    shutil.copy2(temp_output, save_path)
                    gen.saved_path = save_path
                    self.chat_history.append(f"<b>Document saved:</b> {save_path}")
                else:
                    self.chat_history.append("<i>Save cancelled.</i>")

    def _on_quote_weave_complete(self):
        """Handle completion of Weave In mode — reassemble and save."""
        self.send_btn.setEnabled(True)
        if hasattr(self, 'add_quotes_btn') and self.add_quotes_btn:
            self.add_quotes_btn.setEnabled(True)
        self.chat_history.append("<i>Quotes woven into section.</i>")
        self._reassemble_and_save()
```

- [ ] **Step 6: Hide Add Quotes button on case switch**

In `load_case()` (~line 497), in the mediation brief cleanup block, add:

```python
        if hasattr(self, 'add_quotes_btn') and self.add_quotes_btn:
            self.add_quotes_btn.setVisible(False)
```

- [ ] **Step 7: Verify imports work**

```bash
python -c "from icharlotte_core.ui.tabs import ChatTab; print('OK')"
```

Expected: OK

- [ ] **Step 8: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(mediation-brief): Add Quotes button, handlers, reassemble and save logic"
```

---

### Task 5: End-to-End Testing and Polish

**Files:**
- All files from previous tasks (fix issues as needed)

- [ ] **Step 1: Run the full test suite**

```bash
python -m pytest tests/test_mediation_brief.py tests/test_quote_insertion.py -v
```

Expected: All tests pass.

- [ ] **Step 2: Run the existing test suite for regressions**

```bash
python -m pytest tests/ -v --timeout=60
```

Expected: No new failures.

- [ ] **Step 3: Manual end-to-end test**

1. Launch the app, load a case, generate a mediation brief
2. Verify "Add Quotes" button appears after generation
3. Click "Add Quotes" — verify dialog opens
4. Add a deposition transcript PDF
5. Type a search description and click Search
6. Verify results appear with checkboxes, Q&A text, deponent info
7. Test the Edit button on a quote
8. Select Quick Insert mode, pick a section, click "Insert Selected"
9. Verify the document is updated and overwritten
10. Open the document in Word and verify the quote is present with correct formatting
11. Repeat with Weave In mode and verify the section is regenerated with the quote woven in

```bash
python iCharlotte.py
```

- [ ] **Step 4: Fix any issues found**

Address formatting, dialog layout, or logic issues discovered during testing.

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "feat(mediation-brief): complete quote insertion with search, dialog, and save"
```
