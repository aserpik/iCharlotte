# Med Record Extractor Chronology Viewer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace pasted prose entry extraction with a `.docx` chronology viewer where users select Brief Synopsis paragraphs and/or chronology table rows, then extract the selected source-record pages.

**Architecture:** Add a focused backend parser/matcher module that converts chronology `.docx` files into selectable data and row matches. Keep extraction helpers in `icharlotte_core/med_record_extractor.py`, but add a selected-row worker path and move the UI into a new wizard page module wired through the existing in-process task pattern.

**Tech Stack:** Python, PySide6, python-docx, PyMuPDF, pytest, pytest-qt, existing iCharlotte wizard task scaffolding.

---

## File Structure

- Create `icharlotte_core/med_record_chronology.py`
  - Dataclasses for parsed chronology documents, synopsis paragraphs, selectable chronology rows, selection state, and match results.
  - `parse_chronology_document(path)` for `.docx` parsing.
  - `match_synopsis_to_rows(paragraph, rows)` for date/provider matching.
  - Selection-state helpers used by UI tests and the page.
- Modify `icharlotte_core/med_record_extractor.py`
  - Keep existing helper functions: `_parse_page_no`, `_build_file_index`, `_lookup_file`, `_extract_pages`, `_ocr_pdf_if_needed`, `_update_index_document`.
  - Add a selected-row worker path that accepts already parsed chronology rows.
  - Validate `Index of Extracted Records.docx` after saving.
- Create `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`
  - `MedChronologySelectionPage`.
  - `BriefSynopsisPanel`.
  - `ChronologyTablePanel`.
  - `AmbiguousMatchDialog`.
- Modify `icharlotte_core/ui/wizard/in_process_task_tab.py`
  - Change `build_med_extractor_tab` to open a `.docx` picker first.
  - Default picker folder to `find_medical_summary_folder(case_path)` when available.
  - Return an `InProcessTaskTab` backed by `MedChronologySelectionPage` and the selected-row worker.
- Modify tests:
  - `tests/test_med_record_extractor.py`
  - Create `tests/test_wizard/test_med_record_extractor_viewer.py`
  - Update `tests/test_wizard/test_in_process_task_tab.py`
  - Update `tests/test_wizard/test_task_routing.py` only if route expectations need clearer wording.

---

## Task 1: Chronology Parser

**Files:**
- Create: `icharlotte_core/med_record_chronology.py`
- Modify: `tests/test_med_record_extractor.py`

- [ ] **Step 1: Add parser fixture helpers and failing parser tests**

Append these imports and helper functions to `tests/test_med_record_extractor.py` after the existing imports:

```python
import tempfile
from pathlib import Path

from docx import Document
```

Append these helper functions and tests below `TestParsePageNo`:

```python
def _build_chronology_docx(path: Path, *, include_synopsis: bool = True) -> Path:
    doc = Document()
    doc.add_paragraph("CHRONOLOGICAL MEDICAL SUMMARY")
    doc.add_paragraph("Client Name: Test Plaintiff")
    if include_synopsis:
        doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
        doc.add_paragraph(
            "On 09/21/2020, Test Plaintiff presented to Kaiser Permanente. "
            "She was evaluated by Henry Louis Marr, DO, for a right ankle injury."
        )
        doc.add_paragraph(
            "On 09/21/2020, she had an X-ray performed by James Michael Erskine, MD."
        )
    table = doc.add_table(rows=1, cols=5)
    headers = ["DATE", "PAGE NO", "PROVIDER", "DESCRIPTION", "Red Flags/Comments"]
    for index, header in enumerate(headers):
        table.rows[0].cells[index].text = header
    section = table.add_row().cells
    for cell in section:
        cell.text = "POST-INJURY MEDICAL RECORDS"
    row = table.add_row().cells
    row[0].text = "09/21/2020"
    row[1].text = "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 599-604/2530"
    row[2].text = "Kaiser Permanente\nHenry Louis Marr, DO"
    row[3].text = "EMERGENCY DEPARTMENT NOTE\nChief Complaint: Ankle injury."
    row[4].text = ""
    row = table.add_row().cells
    row[0].text = "09/21/2020"
    row[1].text = "Hall - Doc Produced HALL 000001 to 002530 7-21-2023\n\nPg No: 598/2530"
    row[2].text = "Kaiser Permanente\nJames Michael Erskine, MD"
    row[3].text = "RADIOLOGY REPORT\nX-ray right ankle."
    row[4].text = ""
    doc.save(path)
    return path


class TestChronologyDocumentParser(unittest.TestCase):
    def test_parse_chronology_document_extracts_synopsis_and_rows(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(Path(td) / "chronology.docx")
            parsed = parse_chronology_document(str(source))

        self.assertEqual(len(parsed.synopsis_paragraphs), 2)
        self.assertIn("Henry Louis Marr", parsed.synopsis_paragraphs[0].text)
        self.assertEqual(len(parsed.rows), 2)
        self.assertEqual(parsed.rows[0].date, "09/21/2020")
        self.assertEqual(parsed.rows[0].page_start, 599)
        self.assertEqual(parsed.rows[0].page_end, 604)
        self.assertEqual(parsed.rows[0].record_filename, "Hall - Doc Produced HALL 000001 to 002530 7-21-2023")
        self.assertFalse(parsed.blocking_errors)

    def test_parse_chronology_document_warns_when_synopsis_missing(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = _build_chronology_docx(Path(td) / "chronology.docx", include_synopsis=False)
            parsed = parse_chronology_document(str(source))

        self.assertEqual(parsed.synopsis_paragraphs, [])
        self.assertEqual(len(parsed.rows), 2)
        self.assertIn("No Brief Synopsis section found.", parsed.warnings)

    def test_parse_chronology_document_blocks_when_table_missing(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / "no_table.docx"
            doc = Document()
            doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
            doc.add_paragraph("On 09/21/2020, Test Plaintiff saw Kaiser Permanente.")
            doc.save(source)
            parsed = parse_chronology_document(str(source))

        self.assertEqual(parsed.rows, [])
        self.assertIn("No usable 5-column chronology table found.", parsed.blocking_errors)
```

- [ ] **Step 2: Run parser tests to verify they fail**

Run:

```powershell
python -m pytest tests\test_med_record_extractor.py -k "ChronologyDocumentParser" -q
```

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.med_record_chronology'`.

- [ ] **Step 3: Create parser dataclasses and parser implementation**

Create `icharlotte_core/med_record_chronology.py`:

```python
"""Parse medical chronology summary documents for Med Record Extractor."""
from __future__ import annotations

import hashlib
import os
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Iterable, Literal

from docx import Document

from icharlotte_core.med_record_extractor import _normalize_date, _parse_page_no


MatchStatus = Literal["confident", "ambiguous", "none"]


@dataclass(frozen=True)
class SynopsisParagraph:
    id: str
    order: int
    text: str
    warning: str = ""


@dataclass(frozen=True)
class SelectableChronologyRow:
    id: str
    order: int
    date: str
    page_no: str
    provider: str
    description: str
    flags: str
    record_filename: str = ""
    page_start: int = 0
    page_end: int = 0
    warning: str = ""

    @property
    def extractable(self) -> bool:
        return bool(self.record_filename and self.page_start > 0 and self.page_end >= self.page_start)


@dataclass(frozen=True)
class ChronologyDocument:
    source_path: str
    synopsis_paragraphs: list[SynopsisParagraph] = field(default_factory=list)
    rows: list[SelectableChronologyRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    blocking_errors: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class MatchResult:
    status: MatchStatus
    row_ids: tuple[str, ...] = ()
    candidate_row_ids: tuple[str, ...] = ()
    reason: str = ""


_SYNOPSIS_HEADING_RE = re.compile(r"^BRIEF\s+SYNOPSIS\s+OF\s+POST[-\s]INJURY\s+MEDICAL\s+RECORD:?\s*$", re.I)
_MAJOR_HEADING_RE = re.compile(r"^[A-Z0-9][A-Z0-9\s/&.,'()-]{4,}:?$")
_DATE_RE = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b")
_PROVIDER_PATTERNS = (
    re.compile(r"\b(?:evaluated|seen|treated|diagnosed|examined|performed)\s+by\s+([^,.]+(?:,\s*(?:MD|M\.D\.|DO|D\.O\.|PA|P\.A\.|NP|RN))?)", re.I),
    re.compile(r"\bat\s+([A-Z][A-Za-z&.\s]+?)(?:\s+for|\s+with|\.|,)", re.I),
)
_CHRON_HEADERS = ("date", "page no", "provider", "description")


def parse_chronology_document(path: str) -> ChronologyDocument:
    warnings: list[str] = []
    blocking_errors: list[str] = []
    synopsis = _parse_synopsis(path, warnings)
    rows = _parse_rows(path)
    if not synopsis:
        warnings.append("No Brief Synopsis section found.")
    if not rows:
        blocking_errors.append("No usable 5-column chronology table found.")
    return ChronologyDocument(
        source_path=os.path.normpath(path),
        synopsis_paragraphs=synopsis,
        rows=rows,
        warnings=warnings,
        blocking_errors=blocking_errors,
    )


def _parse_synopsis(path: str, warnings: list[str]) -> list[SynopsisParagraph]:
    doc = Document(path)
    in_synopsis = False
    paragraphs: list[SynopsisParagraph] = []
    for para in doc.paragraphs:
        text = _collapse(para.text)
        if not text:
            continue
        if _SYNOPSIS_HEADING_RE.match(text):
            in_synopsis = True
            continue
        if not in_synopsis:
            continue
        if _MAJOR_HEADING_RE.match(text) and not _DATE_RE.search(text):
            break
        if _DATE_RE.search(text):
            order = len(paragraphs)
            paragraphs.append(
                SynopsisParagraph(
                    id=_stable_id("syn", order, text),
                    order=order,
                    text=text,
                )
            )
    return paragraphs


def _parse_rows(path: str) -> list[SelectableChronologyRow]:
    doc = Document(path)
    for table in doc.tables:
        if not table.rows or len(table.rows[0].cells) != 5:
            continue
        headers = [_collapse(cell.text).lower() for cell in table.rows[0].cells]
        if not all(expected in headers[index] for index, expected in enumerate(_CHRON_HEADERS)):
            continue
        rows: list[SelectableChronologyRow] = []
        for raw_row in table.rows[1:]:
            cells = [_collapse(cell.text) for cell in raw_row.cells]
            if not cells[0]:
                continue
            if len(set(cells)) == 1:
                continue
            record_filename, page_start, page_end = _parse_page_no(raw_row.cells[1].text)
            warning = ""
            if not record_filename or page_start <= 0:
                warning = f"Could not parse record/pages from PAGE NO: {cells[1][:80]}"
            order = len(rows)
            rows.append(
                SelectableChronologyRow(
                    id=_stable_id("row", order, "|".join(cells[:4])),
                    order=order,
                    date=cells[0],
                    page_no=raw_row.cells[1].text.strip(),
                    provider=cells[2],
                    description=cells[3],
                    flags=cells[4],
                    record_filename=record_filename,
                    page_start=page_start,
                    page_end=page_end,
                    warning=warning,
                )
            )
        return rows
    return []


def _collapse(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _stable_id(prefix: str, order: int, text: str) -> str:
    digest = hashlib.sha1(text.encode("utf-8", errors="ignore")).hexdigest()[:12]
    return f"{prefix}-{order}-{digest}"
```

- [ ] **Step 4: Run parser tests to verify they pass**

Run:

```powershell
python -m pytest tests\test_med_record_extractor.py -k "ParsePageNo or ChronologyDocumentParser" -q
```

Expected: PASS.

- [ ] **Step 5: Commit parser task**

Run:

```powershell
git add icharlotte_core/med_record_chronology.py tests/test_med_record_extractor.py
git diff --cached --check
git commit -m "feat(med-records): parse chronology documents"
```

---

## Task 2: Synopsis Matching And Selection State

**Files:**
- Modify: `icharlotte_core/med_record_chronology.py`
- Modify: `tests/test_med_record_extractor.py`

- [ ] **Step 1: Write failing matching tests**

Append these tests to `tests/test_med_record_extractor.py`:

```python
class TestSynopsisMatching(unittest.TestCase):
    def _parsed(self):
        from icharlotte_core.med_record_chronology import parse_chronology_document

        td = tempfile.TemporaryDirectory()
        self.addCleanup(td.cleanup)
        source = _build_chronology_docx(Path(td.name) / "chronology.docx")
        return parse_chronology_document(str(source))

    def test_synopsis_match_confident_by_date_and_provider(self):
        from icharlotte_core.med_record_chronology import match_synopsis_to_rows

        parsed = self._parsed()
        result = match_synopsis_to_rows(parsed.synopsis_paragraphs[0], parsed.rows)

        self.assertEqual(result.status, "confident")
        self.assertEqual(result.row_ids, (parsed.rows[0].id,))
        self.assertEqual(result.candidate_row_ids, ())

    def test_synopsis_match_ambiguous_when_provider_is_not_distinct(self):
        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On 09/21/2020, she presented to Kaiser Permanente for ankle care.",
        )
        result = match_synopsis_to_rows(paragraph, parsed.rows)

        self.assertEqual(result.status, "ambiguous")
        self.assertEqual(set(result.candidate_row_ids), {row.id for row in parsed.rows})

    def test_synopsis_match_none_without_same_date(self):
        from icharlotte_core.med_record_chronology import SynopsisParagraph, match_synopsis_to_rows

        parsed = self._parsed()
        paragraph = SynopsisParagraph(
            id="syn-x",
            order=0,
            text="On 12/31/2021, she saw Kaiser Permanente for unrelated care.",
        )
        result = match_synopsis_to_rows(paragraph, parsed.rows)

        self.assertEqual(result.status, "none")
        self.assertEqual(result.row_ids, ())
```

- [ ] **Step 2: Run matching tests to verify they fail**

Run:

```powershell
python -m pytest tests\test_med_record_extractor.py -k "SynopsisMatching" -q
```

Expected: FAIL with `ImportError` for `match_synopsis_to_rows` or assertion failures until matching exists.

- [ ] **Step 3: Implement deterministic matching**

Append these functions to `icharlotte_core/med_record_chronology.py`:

```python
def match_synopsis_to_rows(
    paragraph: SynopsisParagraph,
    rows: Iterable[SelectableChronologyRow],
    *,
    confident_threshold: float = 0.74,
    ambiguous_threshold: float = 0.45,
    gap_threshold: float = 0.18,
) -> MatchResult:
    dates = {_normalize_date(value) for value in _DATE_RE.findall(paragraph.text)}
    if not dates:
        return MatchResult(status="none", reason="No date found in synopsis paragraph.")

    same_date = [row for row in rows if _normalize_date(row.date) in dates]
    if not same_date:
        return MatchResult(status="none", reason="No chronology rows share the paragraph date.")

    provider_candidates = _provider_candidates(paragraph.text)
    if not provider_candidates:
        return MatchResult(
            status="ambiguous",
            candidate_row_ids=tuple(row.id for row in same_date),
            reason="No distinct provider name found in synopsis paragraph.",
        )

    scored = sorted(
        ((_provider_score(provider_candidates, row.provider), row) for row in same_date),
        key=lambda item: item[0],
        reverse=True,
    )
    best_score, best_row = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0
    plausible = [row.id for score, row in scored if score >= ambiguous_threshold]

    if best_score >= confident_threshold and best_score - second_score >= gap_threshold:
        return MatchResult(
            status="confident",
            row_ids=(best_row.id,),
            reason=f"Matched by date and provider score {best_score:.2f}.",
        )
    if plausible:
        return MatchResult(
            status="ambiguous",
            candidate_row_ids=tuple(plausible),
            reason="Multiple same-date provider candidates are plausible.",
        )
    return MatchResult(status="none", reason="No provider candidate matched same-date rows.")


def _provider_candidates(text: str) -> list[str]:
    candidates: list[str] = []
    for pattern in _PROVIDER_PATTERNS:
        for match in pattern.finditer(text):
            candidate = _clean_provider_candidate(match.group(1))
            if candidate and candidate.lower() not in {c.lower() for c in candidates}:
                candidates.append(candidate)
    return candidates


def _clean_provider_candidate(value: str) -> str:
    value = re.sub(r"\s+", " ", value or "").strip(" .,;:")
    value = re.sub(r"\bfor\s+.*$", "", value, flags=re.I).strip(" .,;:")
    return value


def _provider_score(candidates: Iterable[str], provider_cell: str) -> float:
    haystack = _collapse(provider_cell).lower()
    best = 0.0
    for candidate in candidates:
        needle = candidate.lower()
        if needle and needle in haystack:
            best = max(best, 1.0)
        else:
            best = max(best, SequenceMatcher(None, needle, haystack).ratio())
            for part in re.split(r"\s{2,}|\\n|,\\s*", provider_cell):
                best = max(best, SequenceMatcher(None, needle, _collapse(part).lower()).ratio())
    return best
```

- [ ] **Step 4: Add selection-state tests**

Append these tests to `tests/test_med_record_extractor.py`:

```python
class TestSelectionState(unittest.TestCase):
    def test_selection_state_removes_only_paragraph_owned_rows(self):
        from icharlotte_core.med_record_chronology import SelectionState

        state = SelectionState()
        state.select_row("row-manual", source="manual")
        state.select_row("row-shared", source="syn-a")
        state.select_row("row-shared", source="syn-b")
        state.select_row("row-only-a", source="syn-a")

        state.clear_source("syn-a")

        self.assertTrue(state.is_row_selected("row-manual"))
        self.assertTrue(state.is_row_selected("row-shared"))
        self.assertFalse(state.is_row_selected("row-only-a"))
        self.assertEqual(state.selected_row_ids(), ["row-manual", "row-shared"])
```

- [ ] **Step 5: Implement `SelectionState`**

Add this dataclass before `MatchResult` in `icharlotte_core/med_record_chronology.py`:

```python
@dataclass
class SelectionState:
    selected_paragraph_ids: set[str] = field(default_factory=set)
    _row_sources: dict[str, set[str]] = field(default_factory=dict)

    def select_paragraph(self, paragraph_id: str) -> None:
        self.selected_paragraph_ids.add(paragraph_id)

    def deselect_paragraph(self, paragraph_id: str) -> None:
        self.selected_paragraph_ids.discard(paragraph_id)
        self.clear_source(paragraph_id)

    def select_row(self, row_id: str, *, source: str = "manual") -> None:
        self._row_sources.setdefault(row_id, set()).add(source)

    def deselect_row(self, row_id: str, *, source: str = "manual") -> None:
        sources = self._row_sources.get(row_id)
        if not sources:
            return
        sources.discard(source)
        if not sources:
            self._row_sources.pop(row_id, None)

    def clear_source(self, source: str) -> None:
        for row_id in list(self._row_sources):
            self._row_sources[row_id].discard(source)
            if not self._row_sources[row_id]:
                self._row_sources.pop(row_id, None)

    def is_row_selected(self, row_id: str) -> bool:
        return row_id in self._row_sources

    def selected_row_ids(self) -> list[str]:
        return list(self._row_sources.keys())
```

- [ ] **Step 6: Run matching and selection tests**

Run:

```powershell
python -m pytest tests\test_med_record_extractor.py -k "SynopsisMatching or SelectionState" -q
```

Expected: PASS.

- [ ] **Step 7: Commit matching task**

Run:

```powershell
git add icharlotte_core/med_record_chronology.py tests/test_med_record_extractor.py
git diff --cached --check
git commit -m "feat(med-records): match synopsis to chronology rows"
```

---

## Task 3: Selected-Row Extraction Worker

**Files:**
- Modify: `icharlotte_core/med_record_extractor.py`
- Modify: `tests/test_med_record_extractor.py`

- [ ] **Step 1: Write failing selected-row worker tests**

Append these tests to `tests/test_med_record_extractor.py`:

```python
class TestSelectedRowExtractionWorker(unittest.TestCase):
    def test_worker_uses_selected_rows_without_llm_or_auto_chron_lookup(self):
        from unittest.mock import patch
        from icharlotte_core.med_record_chronology import SelectableChronologyRow
        from icharlotte_core.med_record_extractor import MedRecordExtractorWorker

        row = SelectableChronologyRow(
            id="row-1",
            order=0,
            date="09/21/2020",
            page_no="source\n\nPg No: 2/3",
            provider="Kaiser Permanente",
            description="Emergency department note",
            flags="",
            record_filename="source",
            page_start=2,
            page_end=2,
        )
        worker = MedRecordExtractorWorker(
            case_path="C:/case",
            file_number="5800.013",
            chronology_path="C:/case/RECORDS/summary.docx",
            selected_rows=[row],
        )

        with patch.object(worker, "_parse_entries", side_effect=AssertionError("LLM path should not run")), \
             patch("icharlotte_core.med_record_extractor._build_file_index", return_value={"source": "C:/case/RECORDS/source.pdf"}), \
             patch("icharlotte_core.med_record_extractor._extract_pages") as extract_pages, \
             patch("icharlotte_core.med_record_extractor._ocr_pdf_if_needed", return_value=False), \
             patch("icharlotte_core.med_record_extractor._update_index_document"):
            results = worker._execute()

        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].output_path, "C:/case/NOTES/AI OUTPUT/Med Record Extracts/09-21-2020 - Kaiser Permanente - p2.pdf")
        extract_pages.assert_called_once()

    def test_update_index_document_validates_after_save(self):
        from unittest.mock import patch
        from icharlotte_core.med_record_chronology import SelectableChronologyRow
        from icharlotte_core.med_record_extractor import _update_index_document

        row = SelectableChronologyRow(
            id="row-1",
            order=0,
            date="09/21/2020",
            page_no="source\n\nPg No: 2/3",
            provider="Kaiser Permanente",
            description="Emergency department note",
            flags="",
            record_filename="source",
            page_start=2,
            page_end=2,
        )
        with tempfile.TemporaryDirectory() as td, \
             patch("icharlotte_core.med_record_extractor.validate_index_docx") as validate:
            validate.return_value.has_errors = False
            _update_index_document(td, [row])

        validate.assert_called_once()
```

- [ ] **Step 2: Run selected-row worker tests to verify they fail**

Run:

```powershell
python -m pytest tests\test_med_record_extractor.py -k "SelectedRowExtractionWorker" -q
```

Expected: FAIL because `MedRecordExtractorWorker.__init__` does not accept `chronology_path` or `selected_rows`, and `_update_index_document` does not call `validate_index_docx`.

- [ ] **Step 3: Update `ExtractionResult` and worker constructor**

In `icharlotte_core/med_record_extractor.py`, add this import near the other imports:

```python
from typing import Any

from icharlotte_core.word_validator import validate_index_docx
```

Do not import `SelectableChronologyRow` into `med_record_extractor.py`. `med_record_chronology.py` imports `_parse_page_no` and `_normalize_date` from this module, so importing the chronology module back here would create a runtime circular import. The worker can treat selected rows as objects with the expected row attributes.

Update `ExtractionResult`:

```python
@dataclass
class ExtractionResult:
    """Result of extracting pages for one entry or selected chronology row."""
    entry: Optional[ParsedEntry] = None
    matched_row: Optional[object] = None
    output_path: Optional[str] = None
    error: Optional[str] = None
```

Update `MedRecordExtractorWorker.__init__`:

```python
def __init__(
    self,
    case_path: str,
    file_number: str,
    user_text: str = "",
    parent=None,
    *,
    chronology_path: str = "",
    selected_rows: Optional[List[Any]] = None,
):
    super().__init__(parent)
    self.case_path = case_path
    self.file_number = file_number
    self.user_text = user_text
    self.chronology_path = chronology_path
    self.selected_rows = list(selected_rows or [])
```

- [ ] **Step 4: Add selected-row execution branch**

At the start of `_execute`, before parsing pasted text, add:

```python
if self.selected_rows:
    return self._execute_selected_rows(self.selected_rows)
```

Add this method to `MedRecordExtractorWorker`:

```python
def _execute_selected_rows(self, selected_rows: List[Any]) -> List[ExtractionResult]:
    self.progress.emit(f"Preparing to extract {len(selected_rows)} selected chronology row(s)...")
    file_index = _build_file_index(self.case_path)
    output_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT", "Med Record Extracts")
    os.makedirs(output_dir, exist_ok=True)

    results: List[ExtractionResult] = []
    for i, row in enumerate(selected_rows, 1):
        self.progress.emit(f"Processing row {i}/{len(selected_rows)}: {row.date} - {row.provider}")
        result = ExtractionResult(matched_row=row)
        if not row.record_filename or row.page_start <= 0:
            result.error = f"Could not parse record/pages from PAGE NO: {row.page_no[:60]}"
            results.append(result)
            continue

        pdf_path = _lookup_file(file_index, row.record_filename)
        if not pdf_path:
            result.error = f"PDF not found: {row.record_filename}"
            results.append(result)
            self.warning.emit(f"PDF not found: {row.record_filename}")
            continue

        provider_short = row.provider[:40].replace("/", "-").replace("\\", "-")
        date_safe = _normalize_date(row.date).replace("/", "-")
        page_label = f"p{row.page_start}" if row.page_start == row.page_end else f"pp{row.page_start}-{row.page_end}"
        out_name = f"{date_safe} - {provider_short} - {page_label}.pdf"
        out_path = os.path.join(output_dir, out_name)

        try:
            _extract_pages(pdf_path, row.page_start, row.page_end, out_path)
            result.output_path = out_path
            self.progress.emit(f"Extracted: {out_name}")
            try:
                if _ocr_pdf_if_needed(out_path):
                    self.progress.emit(f"OCR applied: {out_name}")
            except Exception as ocr_err:
                self.warning.emit(f"OCR failed for {out_name}: {ocr_err}")
        except Exception as exc:
            result.error = f"Page extraction failed: {exc}"
            self.warning.emit(f"Extraction error for {row.date}: {exc}")
        results.append(result)

    matched_rows = [r.matched_row for r in results if r.matched_row and not r.error]
    if matched_rows:
        self.progress.emit("Updating index document...")
        try:
            _update_index_document(output_dir, matched_rows)
        except Exception as exc:
            self.warning.emit(f"Failed to update index document: {exc}")
            logger.exception("Index document update failed")
    return results
```

- [ ] **Step 5: Make summary formatting tolerate row-only results**

In `run`, replace the failure-detail loop with:

```python
for f in failures:
    row = f.matched_row
    if row is not None:
        summary_parts.append(f"  - {row.date} {row.provider}: {f.error}")
    elif f.entry is not None:
        summary_parts.append(f"  - {f.entry.date} {f.entry.provider}: {f.error}")
    else:
        summary_parts.append(f"  - {f.error}")
```

- [ ] **Step 6: Validate index document after save**

At the end of `_update_index_document`, immediately after `doc.save(index_path)`, add:

```python
result = validate_index_docx(index_path)
if result.has_errors:
    result.print_summary()
    raise ValueError(f"Index validation failed: {result.error_count} error(s)")
```

- [ ] **Step 7: Run selected-row worker tests**

Run:

```powershell
python -m pytest tests\test_med_record_extractor.py -k "SelectedRowExtractionWorker or ParsePageNo" -q
```

Expected: PASS.

- [ ] **Step 8: Commit selected-row worker task**

Run:

```powershell
git add icharlotte_core/med_record_extractor.py tests/test_med_record_extractor.py
git diff --cached --check
git commit -m "feat(med-records): extract selected chronology rows"
```

---

## Task 4: Chronology Selection Page UI

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`
- Create: `tests/test_wizard/test_med_record_extractor_viewer.py`

- [ ] **Step 1: Write failing UI state tests**

Create `tests/test_wizard/test_med_record_extractor_viewer.py`:

```python
"""Tests for the Med Record Extractor chronology viewer."""
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QFileDialog

from tests.test_med_record_extractor import _build_chronology_docx


def test_selection_page_loads_chronology_document(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(case_path=td, file_number="5800.013", chronology_path=str(source))
        qtbot.addWidget(page)

        assert page.selected_count_label.text() == "0 rows selected"
        assert page.synopsis_panel.count() == 2
        assert page.table_panel.count() == 2


def test_direct_row_selection_emits_selected_rows(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(case_path=td, file_number="5800.013", chronology_path=str(source))
        qtbot.addWidget(page)
        page.table_panel.set_row_checked(page.document.rows[0].id, True)

        with qtbot.waitSignal(page.run_requested, timeout=500) as blocker:
            page.extract_btn.click()

        settings = blocker.args[0]
        assert settings["chronology_path"] == str(source)
        assert [row.id for row in settings["selected_rows"]] == [page.document.rows[0].id]


def test_synopsis_selection_auto_selects_confident_row(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(case_path=td, file_number="5800.013", chronology_path=str(source))
        qtbot.addWidget(page)
        page.synopsis_panel.set_paragraph_checked(page.document.synopsis_paragraphs[0].id, True)

        assert page.table_panel.is_row_checked(page.document.rows[0].id)
        assert page.selected_count_label.text() == "1 row selected"


def test_open_original_uses_os_startfile(qtbot):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    with tempfile.TemporaryDirectory() as td:
        source = _build_chronology_docx(Path(td) / "chronology.docx")
        page = MedChronologySelectionPage(case_path=td, file_number="5800.013", chronology_path=str(source))
        qtbot.addWidget(page)
        with patch("icharlotte_core.ui.wizard.pages.med_record_extractor_page.os.startfile") as startfile:
            page.open_original_btn.click()
        startfile.assert_called_once_with(str(source))
```

- [ ] **Step 2: Run UI tests to verify they fail**

Run:

```powershell
python -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -q
```

Expected: FAIL with `ModuleNotFoundError` for `med_record_extractor_page`.

- [ ] **Step 3: Create the viewer page module**

Create `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`:

```python
"""Structured chronology viewer for Med Record Extractor."""
from __future__ import annotations

import os

from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.med_record_chronology import (
    ChronologyDocument,
    MatchResult,
    SelectableChronologyRow,
    SelectionState,
    match_synopsis_to_rows,
    parse_chronology_document,
)
from icharlotte_core.ui.wizard import theme


class BriefSynopsisPanel(QListWidget):
    paragraph_toggled = Signal(str, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._items_by_id: dict[str, QListWidgetItem] = {}
        self.itemChanged.connect(self._on_item_changed)

    def load_paragraphs(self, paragraphs) -> None:
        self.clear()
        self._items_by_id.clear()
        for paragraph in paragraphs:
            item = QListWidgetItem(paragraph.text)
            item.setData(Qt.ItemDataRole.UserRole, paragraph.id)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.addItem(item)
            self._items_by_id[paragraph.id] = item

    def set_paragraph_checked(self, paragraph_id: str, checked: bool) -> None:
        item = self._items_by_id[paragraph_id]
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def mark_warning(self, paragraph_id: str, message: str) -> None:
        item = self._items_by_id.get(paragraph_id)
        if item:
            item.setToolTip(message)

    def _on_item_changed(self, item: QListWidgetItem) -> None:
        paragraph_id = item.data(Qt.ItemDataRole.UserRole)
        self.paragraph_toggled.emit(paragraph_id, item.checkState() == Qt.CheckState.Checked)


class ChronologyTablePanel(QTableWidget):
    row_toggled = Signal(str, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows_by_id: dict[str, int] = {}
        self.setColumnCount(5)
        self.setHorizontalHeaderLabels(["Select", "Date", "Pages", "Provider", "Description"])
        self.cellChanged.connect(self._on_cell_changed)

    def load_rows(self, rows: list[SelectableChronologyRow]) -> None:
        self.blockSignals(True)
        self.setRowCount(len(rows))
        self._rows_by_id.clear()
        for index, row in enumerate(rows):
            self._rows_by_id[row.id] = index
            checkbox = QTableWidgetItem("")
            checkbox.setFlags(checkbox.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            checkbox.setCheckState(Qt.CheckState.Unchecked)
            checkbox.setData(Qt.ItemDataRole.UserRole, row.id)
            self.setItem(index, 0, checkbox)
            self.setItem(index, 1, QTableWidgetItem(row.date))
            page_text = f"{row.page_start}-{row.page_end}" if row.page_start != row.page_end else str(row.page_start)
            self.setItem(index, 2, QTableWidgetItem(page_text if row.extractable else "unparsed"))
            self.setItem(index, 3, QTableWidgetItem(row.provider))
            self.setItem(index, 4, QTableWidgetItem(row.description[:220]))
        self.blockSignals(False)

    def set_row_checked(self, row_id: str, checked: bool, *, emit: bool = True) -> None:
        row_index = self._rows_by_id[row_id]
        item = self.item(row_index, 0)
        if not emit:
            self.blockSignals(True)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        if not emit:
            self.blockSignals(False)

    def is_row_checked(self, row_id: str) -> bool:
        row_index = self._rows_by_id[row_id]
        return self.item(row_index, 0).checkState() == Qt.CheckState.Checked

    def _on_cell_changed(self, row: int, column: int) -> None:
        if column != 0:
            return
        item = self.item(row, 0)
        row_id = item.data(Qt.ItemDataRole.UserRole)
        self.row_toggled.emit(row_id, item.checkState() == Qt.CheckState.Checked)


class AmbiguousMatchDialog(QDialog):
    def __init__(self, paragraph_text: str, candidates: list[SelectableChronologyRow], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Confirm matching chronology rows")
        self._selected: list[str] = []
        layout = QVBoxLayout(self)
        label = QLabel(paragraph_text)
        label.setWordWrap(True)
        layout.addWidget(label)
        self.table = ChronologyTablePanel()
        self.table.load_rows(candidates)
        layout.addWidget(self.table)
        buttons = QHBoxLayout()
        confirm = QPushButton("Use Selected Rows")
        cancel = QPushButton("Cancel")
        confirm.clicked.connect(self._accept_selected)
        cancel.clicked.connect(self.reject)
        buttons.addStretch()
        buttons.addWidget(cancel)
        buttons.addWidget(confirm)
        layout.addLayout(buttons)

    def selected_row_ids(self) -> list[str]:
        return list(self._selected)

    def _accept_selected(self) -> None:
        self._selected = [
            row_id for row_id in self.table._rows_by_id
            if self.table.is_row_checked(row_id)
        ]
        self.accept()


class MedChronologySelectionPage(QWidget):
    run_requested = Signal(dict)

    def __init__(self, case_path: str, file_number: str, chronology_path: str, parent=None):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.chronology_path = chronology_path
        self.document: ChronologyDocument = parse_chronology_document(chronology_path)
        self.selection = SelectionState()
        self._paragraphs = {p.id: p for p in self.document.synopsis_paragraphs}
        self._rows = {r.id: r for r in self.document.rows}
        self._setup_ui()
        self._refresh_extract_state()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)
        title = theme.page_title("Select medical chronology records")
        layout.addWidget(title)
        self.warning_label = QLabel("; ".join(self.document.warnings + self.document.blocking_errors))
        self.warning_label.setWordWrap(True)
        layout.addWidget(self.warning_label)
        self.synopsis_panel = BriefSynopsisPanel()
        self.synopsis_panel.load_paragraphs(self.document.synopsis_paragraphs)
        self.synopsis_panel.paragraph_toggled.connect(self._on_paragraph_toggled)
        layout.addWidget(QLabel("Brief Synopsis"))
        layout.addWidget(self.synopsis_panel, 1)
        self.table_panel = ChronologyTablePanel()
        self.table_panel.load_rows(self.document.rows)
        self.table_panel.row_toggled.connect(self._on_row_toggled)
        layout.addWidget(QLabel("Medical Chronology Table"))
        layout.addWidget(self.table_panel, 2)
        controls = QHBoxLayout()
        self.selected_count_label = QLabel("")
        controls.addWidget(self.selected_count_label)
        controls.addStretch()
        self.open_original_btn = theme.secondary_button("Open Original")
        self.open_original_btn.clicked.connect(self._open_original)
        controls.addWidget(self.open_original_btn)
        self.extract_btn = theme.primary_button("Extract Records")
        self.extract_btn.clicked.connect(self._emit_run)
        controls.addWidget(self.extract_btn)
        layout.addLayout(controls)

    def _on_paragraph_toggled(self, paragraph_id: str, checked: bool) -> None:
        if checked:
            self.selection.select_paragraph(paragraph_id)
            result = match_synopsis_to_rows(self._paragraphs[paragraph_id], self.document.rows)
            self._apply_match(paragraph_id, result)
        else:
            self.selection.deselect_paragraph(paragraph_id)
        self._sync_table_checks()
        self._refresh_extract_state()

    def _apply_match(self, paragraph_id: str, result: MatchResult) -> None:
        if result.status == "confident":
            for row_id in result.row_ids:
                self.selection.select_row(row_id, source=paragraph_id)
        elif result.status == "ambiguous":
            candidates = [self._rows[row_id] for row_id in result.candidate_row_ids]
            dialog = AmbiguousMatchDialog(self._paragraphs[paragraph_id].text, candidates, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                for row_id in dialog.selected_row_ids():
                    self.selection.select_row(row_id, source=paragraph_id)
            else:
                self.synopsis_panel.mark_warning(paragraph_id, result.reason)
        else:
            self.synopsis_panel.mark_warning(paragraph_id, result.reason)

    def _on_row_toggled(self, row_id: str, checked: bool) -> None:
        if checked:
            self.selection.select_row(row_id, source="manual")
        else:
            self.selection.deselect_row(row_id, source="manual")
        self._refresh_extract_state()

    def _sync_table_checks(self) -> None:
        for row_id in self._rows:
            self.table_panel.set_row_checked(row_id, self.selection.is_row_selected(row_id), emit=False)

    def _selected_rows(self) -> list[SelectableChronologyRow]:
        return [self._rows[row_id] for row_id in self.selection.selected_row_ids() if self._rows[row_id].extractable]

    def _refresh_extract_state(self) -> None:
        count = len(self._selected_rows())
        self.selected_count_label.setText(f"{count} row{'s' if count != 1 else ''} selected")
        self.extract_btn.setEnabled(count > 0 and not self.document.blocking_errors)

    def _emit_run(self) -> None:
        rows = self._selected_rows()
        if not rows:
            QMessageBox.information(self, "No extractable rows", "Select at least one extractable chronology row.")
            return
        self.run_requested.emit({
            "chronology_path": self.chronology_path,
            "selected_rows": rows,
        })

    def _open_original(self) -> None:
        os.startfile(self.chronology_path)
```

- [ ] **Step 4: Run UI tests**

Run:

```powershell
python -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit viewer page task**

Run:

```powershell
git add icharlotte_core/ui/wizard/pages/med_record_extractor_page.py tests/test_wizard/test_med_record_extractor_viewer.py
git diff --cached --check
git commit -m "feat(wizard): add med record chronology viewer"
```

---

## Task 5: Wizard Routing And File Picker

**Files:**
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py`
- Modify: `tests/test_wizard/test_in_process_task_tab.py`
- Modify: `tests/test_wizard/test_med_record_extractor_viewer.py`

- [ ] **Step 1: Write failing builder/file-picker tests**

Append these imports to `tests/test_wizard/test_med_record_extractor_viewer.py`:

```python
from icharlotte_core.ui.wizard.registry import get_task
```

Append these tests:

```python
def test_build_med_extractor_tab_uses_docx_picker_and_summary_folder(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.in_process_task_tab import build_med_extractor_tab
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    summary_dir = tmp_path / "RECORDS" / "Medical Summary - DO NOT PRODUCE"
    summary_dir.mkdir(parents=True)
    source = _build_chronology_docx(summary_dir / "chronology.docx")

    with patch(
        "icharlotte_core.ui.wizard.in_process_task_tab.QFileDialog.getOpenFileName",
        return_value=(str(source), "Word Documents (*.docx)"),
    ) as picker:
        tab = build_med_extractor_tab(
            get_task("med_record_extractor"),
            case_path=str(tmp_path),
            file_number="5800.013",
            parent=None,
        )

    qtbot.addWidget(tab)
    assert isinstance(tab.settings_page, MedChronologySelectionPage)
    assert picker.call_args.args[2] == str(summary_dir)


def test_build_med_extractor_tab_cancel_returns_none(tmp_path):
    from icharlotte_core.ui.wizard.in_process_task_tab import build_med_extractor_tab

    with patch(
        "icharlotte_core.ui.wizard.in_process_task_tab.QFileDialog.getOpenFileName",
        return_value=("", ""),
    ):
        tab = build_med_extractor_tab(
            get_task("med_record_extractor"),
            case_path=str(tmp_path),
            file_number="5800.013",
            parent=None,
        )

    assert tab is None
```

Remove or rewrite `test_med_extractor_settings_rejects_empty` and `test_med_extractor_settings_emits_text` in `tests/test_wizard/test_in_process_task_tab.py` because pasted text is no longer the task entry point. Replace them with:

```python
def test_med_extractor_output_show_result(qtbot, tmp_path):
    page = MedExtractorOutputPage(str(tmp_path))
    qtbot.addWidget(page)
    page.show_result("3 record(s) extracted")
    assert "3 record(s) extracted" in page.summary_view.toPlainText()
```

Keep only one copy of `test_med_extractor_output_show_result` if it already exists.

- [ ] **Step 2: Run builder tests to verify they fail**

Run:

```powershell
python -m pytest tests\test_wizard\test_med_record_extractor_viewer.py tests\test_wizard\test_in_process_task_tab.py -q
```

Expected: FAIL because `build_med_extractor_tab` still builds `MedExtractorSettingsPage` and does not open a `.docx` picker.

- [ ] **Step 3: Update imports in `in_process_task_tab.py`**

Add `find_medical_summary_folder` to the local imports near the existing builder imports by importing inside `build_med_extractor_tab`:

```python
from icharlotte_core.ui.wizard.file_picker import find_medical_summary_folder, resolve_default_folder
from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage
```

- [ ] **Step 4: Replace `build_med_extractor_tab`**

Replace the existing `build_med_extractor_tab` with:

```python
def build_med_extractor_tab(spec, case_path: str, file_number: str, parent: QWidget | None) -> InProcessTaskTab | None:
    from icharlotte_core.med_record_extractor import MedRecordExtractorWorker
    from icharlotte_core.ui.wizard.file_picker import find_medical_summary_folder, resolve_default_folder
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    start_dir = find_medical_summary_folder(case_path) or resolve_default_folder(case_path, ["RECORDS"])
    chronology_path, _ = QFileDialog.getOpenFileName(
        parent,
        "Select medical chronology summary",
        start_dir,
        "Word Documents (*.docx)",
    )
    if not chronology_path:
        return None

    def factory(cp, fn, settings, p):
        return MedRecordExtractorWorker(
            cp,
            fn,
            parent=p,
            chronology_path=settings.get("chronology_path", chronology_path),
            selected_rows=settings.get("selected_rows", []),
        )

    return InProcessTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        settings_widget=MedChronologySelectionPage(case_path, file_number, chronology_path),
        output_widget=MedExtractorOutputPage(case_path),
        worker_factory=factory,
        auto_run=False,
        parent=parent,
    )
```

- [ ] **Step 5: Include chronology file in task completion files**

In `InProcessTaskTab._on_worker_finished`, before building `entry`, compute:

```python
entry_files = []
settings_files = settings.get("files")
if isinstance(settings_files, list):
    entry_files.extend(str(path) for path in settings_files if path)
chronology_path = settings.get("chronology_path")
if chronology_path:
    entry_files.append(str(chronology_path))
```

Then set:

```python
"files": entry_files,
```

instead of `"files": []`.

- [ ] **Step 6: Run builder and in-process tests**

Run:

```powershell
python -m pytest tests\test_wizard\test_med_record_extractor_viewer.py tests\test_wizard\test_in_process_task_tab.py -q
```

Expected: PASS.

- [ ] **Step 7: Commit routing task**

Run:

```powershell
git add icharlotte_core/ui/wizard/in_process_task_tab.py tests/test_wizard/test_in_process_task_tab.py tests/test_wizard/test_med_record_extractor_viewer.py
git diff --cached --check
git commit -m "feat(wizard): route med extractor through chronology picker"
```

---

## Task 6: Restore, Error States, And Final Verification

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/med_record_extractor_page.py`
- Modify: `tests/test_wizard/test_med_record_extractor_viewer.py`
- Modify: `tests/test_wizard/test_task_routing.py` if route assertions need updated comments only

- [ ] **Step 1: Add restore/error-state tests**

Append these tests to `tests/test_wizard/test_med_record_extractor_viewer.py`:

```python
def test_selection_page_to_dict_and_from_dict_restore_selection(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    source = _build_chronology_docx(tmp_path / "chronology.docx")
    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)
    page.table_panel.set_row_checked(page.document.rows[0].id, True)
    saved = page.to_dict()

    restored = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(restored)
    restored.from_dict(saved)

    assert restored.table_panel.is_row_checked(restored.document.rows[0].id)


def test_selection_page_blocks_extract_without_table(qtbot, tmp_path):
    from docx import Document
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    source = tmp_path / "bad.docx"
    doc = Document()
    doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
    doc.add_paragraph("On 09/21/2020, she saw Kaiser Permanente.")
    doc.save(source)

    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)

    assert "No usable 5-column chronology table found." in page.warning_label.text()
    assert not page.extract_btn.isEnabled()


def test_selection_page_reports_malformed_page_no(qtbot, tmp_path):
    from docx import Document
    from icharlotte_core.ui.wizard.pages.med_record_extractor_page import MedChronologySelectionPage

    source = tmp_path / "bad_page.docx"
    doc = Document()
    doc.add_paragraph("BRIEF SYNOPSIS OF POST-INJURY MEDICAL RECORD:")
    doc.add_paragraph("On 09/21/2020, she saw Kaiser Permanente.")
    table = doc.add_table(rows=1, cols=5)
    for i, header in enumerate(["DATE", "PAGE NO", "PROVIDER", "DESCRIPTION", "Red Flags/Comments"]):
        table.rows[0].cells[i].text = header
    row = table.add_row().cells
    row[0].text = "09/21/2020"
    row[1].text = "source\n\nPg no: 17-"
    row[2].text = "Kaiser Permanente"
    row[3].text = "Bad page reference"
    row[4].text = ""
    doc.save(source)

    page = MedChronologySelectionPage(str(tmp_path), "5800.013", str(source))
    qtbot.addWidget(page)
    page.table_panel.set_row_checked(page.document.rows[0].id, True)

    assert page.selected_count_label.text() == "0 rows selected"
    assert not page.extract_btn.isEnabled()
    assert page.document.rows[0].warning.startswith("Could not parse record/pages")
```

- [ ] **Step 2: Run restore/error-state tests to verify they fail**

Run:

```powershell
python -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -k "restore or blocks or malformed" -q
```

Expected: FAIL because `to_dict`, `from_dict`, and malformed-row UI behavior are incomplete.

- [ ] **Step 3: Implement page persistence methods**

Add these methods to `MedChronologySelectionPage`:

```python
def to_dict(self) -> dict:
    return {
        "chronology_path": self.chronology_path,
        "selected_paragraph_ids": sorted(self.selection.selected_paragraph_ids),
        "selected_row_ids": self.selection.selected_row_ids(),
    }

def from_dict(self, data: dict) -> None:
    for paragraph_id in data.get("selected_paragraph_ids", []):
        if paragraph_id in self._paragraphs:
            self.synopsis_panel.set_paragraph_checked(paragraph_id, True)
    for row_id in data.get("selected_row_ids", []):
        if row_id in self._rows:
            self.table_panel.set_row_checked(row_id, True)
    self._refresh_extract_state()
```

- [ ] **Step 4: Mark unextractable rows in the table**

In `ChronologyTablePanel.load_rows`, after setting row cells, add:

```python
if row.warning:
    for column in range(self.columnCount()):
        cell = self.item(index, column)
        if cell is not None:
            cell.setToolTip(row.warning)
    checkbox.setFlags(checkbox.flags() & ~Qt.ItemFlag.ItemIsEnabled)
```

- [ ] **Step 5: Run focused tests**

Run:

```powershell
python -m pytest tests\test_med_record_extractor.py tests\test_wizard\test_med_record_extractor_viewer.py tests\test_wizard\test_in_process_task_tab.py -q
```

Expected: PASS.

- [ ] **Step 6: Run broader wizard tests**

Run:

```powershell
python -m pytest tests\test_wizard -q
```

Expected: PASS.

- [ ] **Step 7: Final diff audit**

Run:

```powershell
git diff --stat
git diff --check
git status --short
```

Expected: only files from this plan are modified, `git diff --check` passes, and unrelated pre-existing dirty files remain unstaged.

- [ ] **Step 8: Commit final cleanup task**

Run:

```powershell
git add icharlotte_core/ui/wizard/pages/med_record_extractor_page.py tests/test_wizard/test_med_record_extractor_viewer.py tests/test_wizard/test_task_routing.py
git diff --cached --check
git commit -m "test(wizard): cover med extractor viewer restore states"
```

---

## Completion Checklist

- [ ] `python -m pytest tests\test_med_record_extractor.py -q` passes.
- [ ] `python -m pytest tests\test_wizard\test_med_record_extractor_viewer.py -q` passes.
- [ ] `python -m pytest tests\test_wizard\test_in_process_task_tab.py -q` passes.
- [ ] `python -m pytest tests\test_wizard -q` passes, or any unrelated failures are documented with exact failure text.
- [ ] `git diff --check` passes.
- [ ] No untracked generated `.superpowers` brainstorming files are included.
- [ ] The final user report lists implemented behavior, test evidence, and any residual risks.
