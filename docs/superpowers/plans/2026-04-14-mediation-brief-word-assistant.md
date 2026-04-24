# Mediation Brief Refinement & Quote Insertion from Word AI Assistant — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add "Mediation Brief: Refine Section" and "Mediation Brief: Add Quotes" entries to the Word AI assistant popup (Win+V) so the user can refine brief sections and insert deposition quotes directly against a live Word document, without going back to the chat tab.

**Architecture:** The Word popup operates on the live Word document. A new `mediation_brief_live.py` module parses a Word doc into a `LiveBrief` dataclass on demand (no sidecar persistence, no session state). `MediationBriefGenerator` gains a `build_refinement_prompts` method that returns `(system_prompt, full_prompt)` for a given section without mutating instance state, so the existing `TaskLLMWorkerThread` LLM path can execute the refinement. A new stripped `WordQuoteInsertionDialog` reuses the existing search worker and result widgets but removes section/subsection/Weave UI. Quote insertion builds a temporary .docx via the existing `_add_depo_quote` python-docx formatter and inserts it at cursor via Word COM `Range.InsertFile`. Everything goes through the existing `TaskManager` + bookmark path for serialization and status-bar UX.

**Tech Stack:** Python 3, PySide6/PyQt6, python-docx, pywin32 Word COM, existing iCharlotte `TaskManager` infrastructure.

**Spec:** `docs/superpowers/specs/2026-04-14-mediation-brief-word-assistant-design.md`

---

## File Structure

**New files:**
- `icharlotte_core/mediation_brief_live.py` — LiveSection/LiveBrief dataclasses, parser, range locator, quote insertion helper.
- `icharlotte_core/ui/quote_dialog_word.py` — `WordQuoteInsertionDialog` — stripped variant of the existing dialog for Word-cursor insertion.
- `tests/test_mediation_brief_live.py` — unit tests for the parser, detection, range mapping, and quote helper.
- `tests/test_mediation_brief_refinement_prompts.py` — unit tests for `build_refinement_prompts`.

**Modified files:**
- `icharlotte_core/mediation_brief.py` — add `build_refinement_prompts` method; refactor `_build_section_prompt` only if needed.
- `icharlotte_core/word_hotkey.py` — sentinel constants, detection hook on popup open, inline section combo, prompt-select branching, `_do_execute` dispatch for the two new operations, quote-insertion synthetic task.

**Unchanged (important):**
- `icharlotte_core/ui/tabs.py` — chat tab integration stays as is.
- `icharlotte_core/ui/quote_dialog.py` — existing dialog stays as is. New `quote_dialog_word.py` imports shared widgets from it.

---

## Task Graph

Tasks are ordered so earlier tasks land dependencies for later ones. Run them in order.

1. `build_refinement_prompts` on the generator (pure Python, no UI).
2. `LiveSection`/`LiveBrief` dataclasses + `parse_brief_from_word_doc` + `is_mediation_brief` (new module, no UI).
3. `get_word_range_for_section` helper in the same module.
4. `insert_formatted_quotes_at_range` helper in the same module.
5. `WordQuoteInsertionDialog` (stripped dialog).
6. Word popup — sentinel constants + dynamic dropdown entries on open.
7. Word popup — inline section combo, visibility wired to dropdown selection.
8. Word popup — `_do_execute` dispatch for "Mediation Brief: Refine Section".
9. Word popup — `_do_execute` dispatch for "Mediation Brief: Add Quotes".
10. Manual integration test checklist + memory update.

---

## Task 1: `build_refinement_prompts` on MediationBriefGenerator

**Files:**
- Modify: `icharlotte_core/mediation_brief.py` (add method near `refine_sections`, around current line 1141)
- Test: `tests/test_mediation_brief_refinement_prompts.py` (new)

**Context for the engineer:**
- `MediationBriefGenerator._build_system_prompt(section)` returns the persona + style guide + formatting rules.
- `MediationBriefGenerator._build_section_prompt(section, refinement_instruction)` reads `self.sections` and `self.planning_output` to build the user prompt. It references `self.sections[s]` for prior sections.
- We want a stateless wrapper that temporarily swaps `self.sections` with a caller-provided dict, builds the prompts, and restores state. This keeps the existing chat-tab flow untouched.
- No LLM call here. We just return the two strings; the caller submits them through `TaskLLMWorkerThread` which performs the actual `LLMHandler.generate()` call.

- [ ] **Step 1: Write the failing test**

Create `tests/test_mediation_brief_refinement_prompts.py`:

```python
"""Unit tests for MediationBriefGenerator.build_refinement_prompts."""
import unittest
from unittest.mock import patch

from icharlotte_core.mediation_brief import MediationBriefGenerator


class TestBuildRefinementPrompts(unittest.TestCase):
    def setUp(self):
        self.gen = MediationBriefGenerator()
        self.sections_dict = {
            "introduction": "Plaintiff Smith sues Defendant Jones for negligence.",
            "statement_of_facts": "On March 1, 2025, the parties met at the crossing.",
            "liability": "Defendant had the right of way under Vehicle Code 21800.",
            "damages": "Plaintiff alleges $50,000 in medical specials.",
            "settlement_position": "Defendant offers $15,000.",
            "conclusion": "For the foregoing reasons, liability is in dispute.",
        }

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_returns_system_and_full_prompt_strings(self, _mock_excerpts):
        system_prompt, full_prompt = self.gen.build_refinement_prompts(
            section_name="liability",
            sections_dict=self.sections_dict,
            instruction="Make the causation discussion more forceful.",
        )
        self.assertIsInstance(system_prompt, str)
        self.assertIsInstance(full_prompt, str)
        self.assertIn("defense litigation attorney", system_prompt)
        self.assertIn("Make the causation discussion more forceful.", full_prompt)

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_prompt_includes_current_section_text(self, _mock_excerpts):
        _, full_prompt = self.gen.build_refinement_prompts(
            section_name="liability",
            sections_dict=self.sections_dict,
            instruction="Tighten.",
        )
        # The refinement prompt should reference the current Liability text
        # (via the PREVIOUSLY DRAFTED SECTIONS context block, which is how the
        # generator already passes prior sections into section prompts).
        self.assertIn("Vehicle Code 21800", full_prompt)

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_does_not_mutate_generator_sections(self, _mock_excerpts):
        self.gen.sections = {"introduction": "unchanged"}
        self.gen.build_refinement_prompts(
            section_name="liability",
            sections_dict=self.sections_dict,
            instruction="Tighten.",
        )
        self.assertEqual(self.gen.sections, {"introduction": "unchanged"})

    @patch.object(MediationBriefGenerator, "get_style_excerpts", return_value={})
    def test_does_not_mutate_generator_sections_on_exception(self, _mock_excerpts):
        self.gen.sections = {"introduction": "unchanged"}
        with patch.object(
            MediationBriefGenerator, "_build_section_prompt",
            side_effect=RuntimeError("boom"),
        ):
            with self.assertRaises(RuntimeError):
                self.gen.build_refinement_prompts(
                    section_name="liability",
                    sections_dict=self.sections_dict,
                    instruction="Tighten.",
                )
        self.assertEqual(self.gen.sections, {"introduction": "unchanged"})


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to confirm they fail**

Run: `python -m pytest tests/test_mediation_brief_refinement_prompts.py -v`
Expected: FAIL with `AttributeError: 'MediationBriefGenerator' object has no attribute 'build_refinement_prompts'`.

- [ ] **Step 3: Add the method to `MediationBriefGenerator`**

Open `icharlotte_core/mediation_brief.py`. Immediately after the `refine_sections` method (around line 1190), add:

```python
    def build_refinement_prompts(
        self,
        section_name: str,
        sections_dict: Dict[str, str],
        instruction: str,
    ) -> tuple:
        """Return ``(system_prompt, full_prompt)`` for refining *section_name*.

        Stateless sibling of ``refine_sections`` — used by the Word AI
        assistant popup so refinement can run against a live Word document
        without touching this generator's in-memory state.

        Temporarily swaps ``self.sections`` with *sections_dict* while the
        existing prompt builders run, then restores the original value.

        Args:
            section_name: Canonical section name (e.g. ``"liability"``).
            sections_dict: Mapping of canonical section name → current body
                text, parsed from the live Word document.
            instruction: The user's refinement instruction.

        Returns:
            ``(system_prompt, full_prompt)`` — the two prompt strings ready to
            be submitted through ``TaskLLMWorkerThread``.
        """
        saved_sections = self.sections
        self.sections = dict(sections_dict)
        try:
            system_prompt = self._build_system_prompt(section_name)
            full_prompt = self._build_section_prompt(
                section_name, refinement_instruction=instruction
            )
            return system_prompt, full_prompt
        finally:
            self.sections = saved_sections
```

- [ ] **Step 4: Run the tests and confirm they pass**

Run: `python -m pytest tests/test_mediation_brief_refinement_prompts.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_mediation_brief_refinement_prompts.py
git commit -m "feat(mediation-brief): stateless build_refinement_prompts for Word popup use"
```

---

## Task 2: Live-document parser — dataclasses, `parse_brief_from_word_doc`, `is_mediation_brief`

**Files:**
- Create: `icharlotte_core/mediation_brief_live.py`
- Test: `tests/test_mediation_brief_live.py` (new)

**Context for the engineer:**
- Word COM exposes `doc.Paragraphs` as a 1-indexed collection. `doc.Paragraphs(i)` returns a Paragraph object; `para.Range.Text` is its text (with a trailing `\r`); `para.Range.Start` / `.End` are absolute character positions in the document.
- The existing regex `_HEADING_PATTERN` and dict `_HEADING_TO_SECTION` in `mediation_brief.py` match roman-numeral headings and map them to canonical section names. Reuse them — do NOT redefine them.
- `_HEADING_PATTERN` is anchored with `^...$` and `re.MULTILINE` — it expects a full heading line. When matching a single paragraph's text (which has a trailing `\r`), strip trailing whitespace/control chars first.
- For unit tests you don't need real Word COM; you can pass a fake doc-like object. The parser should duck-type on `doc.Paragraphs` (1-indexed) and each paragraph's `.Range.Text` / `.Range.Start` / `.Range.End` attributes.

- [ ] **Step 1: Write the failing test**

Create `tests/test_mediation_brief_live.py`:

```python
"""Unit tests for mediation_brief_live parser and detection."""
import unittest
from dataclasses import dataclass
from typing import List


# --- Fake Word COM objects for tests ------------------------------------------

@dataclass
class FakeRange:
    Text: str
    Start: int
    End: int


@dataclass
class FakeParagraph:
    Range: FakeRange


class FakeParagraphCollection:
    """1-indexed collection that mimics Word COM doc.Paragraphs."""
    def __init__(self, paragraphs: List[FakeParagraph]):
        self._items = paragraphs

    def __call__(self, i):  # doc.Paragraphs(i)
        return self._items[i - 1]

    @property
    def Count(self):
        return len(self._items)

    def __iter__(self):
        return iter(self._items)


class FakeDoc:
    def __init__(self, paragraphs: List[FakeParagraph], full_name: str = "test.docx"):
        self.Paragraphs = FakeParagraphCollection(paragraphs)
        self.FullName = full_name


def make_doc(*paragraph_texts: str) -> FakeDoc:
    """Build a FakeDoc from paragraph text strings.

    Each paragraph becomes a FakeParagraph with a synthetic Range. Start/End
    are computed so the absolute character positions make sense.
    """
    paragraphs: List[FakeParagraph] = []
    cursor = 0
    for t in paragraph_texts:
        # Word COM always appends \r to each paragraph's Range.Text.
        text_with_cr = t + "\r"
        start = cursor
        end = cursor + len(text_with_cr)
        paragraphs.append(FakeParagraph(Range=FakeRange(Text=text_with_cr, Start=start, End=end)))
        cursor = end
    return FakeDoc(paragraphs)


# --- Tests --------------------------------------------------------------------

from icharlotte_core.mediation_brief_live import (
    LiveBrief,
    LiveSection,
    is_mediation_brief,
    parse_brief_from_word_doc,
)


class TestIsMediationBrief(unittest.TestCase):
    def test_empty_doc_returns_false(self):
        self.assertFalse(is_mediation_brief(make_doc()))

    def test_doc_with_two_headings_returns_false(self):
        doc = make_doc(
            "I.   INTRODUCTION",
            "Plaintiff Smith sues Defendant Jones.",
            "II.  STATEMENT OF FACTS",
            "On March 1, 2025 the parties met.",
        )
        self.assertFalse(is_mediation_brief(doc))

    def test_doc_with_three_headings_returns_true(self):
        doc = make_doc(
            "I.   INTRODUCTION",
            "Plaintiff Smith sues Defendant Jones.",
            "II.  STATEMENT OF FACTS",
            "On March 1, 2025 the parties met.",
            "IV.  LIABILITY",
            "Defendant had the right of way.",
        )
        self.assertTrue(is_mediation_brief(doc))

    def test_non_brief_doc_returns_false(self):
        doc = make_doc(
            "Memo to file",
            "Regarding the Smith matter, please note the following.",
            "Thank you.",
        )
        self.assertFalse(is_mediation_brief(doc))


class TestParseBriefFromWordDoc(unittest.TestCase):
    def _full_brief_doc(self):
        return make_doc(
            "Caption page first line",
            "I.   INTRODUCTION",
            "Plaintiff Smith sues Defendant Jones (\"Jones\") for negligence.",
            "II.  STATEMENT OF FACTS",
            "On March 1, 2025, the parties met at the crossing.",
            "III. PROCEDURAL STATUS",
            "Trial is set for November 2026.",
            "IV.  LIABILITY",
            "Defendant had the right of way under Vehicle Code 21800.",
            "V.   DAMAGES",
            "Plaintiff alleges $50,000 in medical specials.",
            "VI.  SETTLEMENT POSITION",
            "Defendant offers $15,000.",
            "VII. CONCLUSION",
            "For the foregoing reasons, liability is in dispute.",
        )

    def test_parses_all_canonical_sections(self):
        live = parse_brief_from_word_doc(self._full_brief_doc())
        self.assertIsInstance(live, LiveBrief)
        self.assertEqual(
            set(live.sections.keys()),
            {
                "introduction",
                "statement_of_facts",
                "procedural_status",
                "liability",
                "damages",
                "settlement_position",
                "conclusion",
            },
        )

    def test_section_body_text_matches_paragraph(self):
        live = parse_brief_from_word_doc(self._full_brief_doc())
        liability = live.sections["liability"]
        self.assertIn("Vehicle Code 21800", liability.text)
        # Heading itself should NOT be part of the body text.
        self.assertNotIn("IV.", liability.text)
        self.assertNotIn("LIABILITY", liability.text)

    def test_section_paragraph_indices_are_one_based(self):
        live = parse_brief_from_word_doc(self._full_brief_doc())
        # Layout: caption=1, I=2, body=3, II=4, body=5, III=6, body=7,
        #         IV=8, body=9, V=10, body=11, VI=12, body=13, VII=14, body=15
        liability = live.sections["liability"]
        self.assertEqual(liability.start_para_index, 9)
        self.assertEqual(liability.end_para_index, 9)

    def test_heading_variant_maps_to_canonical(self):
        doc = make_doc(
            "I.   OVERVIEW",
            "Plaintiff sues.",
            "II.  FACTUAL BACKGROUND",
            "Facts here.",
            "IV.  ANALYSIS OF LIABILITY",
            "Liability text.",
        )
        live = parse_brief_from_word_doc(doc)
        self.assertIn("introduction", live.sections)
        self.assertIn("statement_of_facts", live.sections)
        self.assertIn("liability", live.sections)
        self.assertEqual(
            live.sections["statement_of_facts"].text.strip(), "Facts here."
        )

    def test_unrecognised_heading_skipped_silently(self):
        doc = make_doc(
            "I.   INTRODUCTION",
            "Intro body.",
            "II.  APPENDIX",
            "Appendix body.",
            "III. LIABILITY",
            "Liability body.",
        )
        live = parse_brief_from_word_doc(doc)
        self.assertIn("introduction", live.sections)
        self.assertIn("liability", live.sections)
        self.assertNotIn("appendix", live.sections)

    def test_multiparagraph_body(self):
        doc = make_doc(
            "I.   INTRODUCTION",
            "First intro paragraph.",
            "Second intro paragraph.",
            "II.  STATEMENT OF FACTS",
            "Facts.",
            "III. LIABILITY",
            "Liability.",
        )
        live = parse_brief_from_word_doc(doc)
        intro = live.sections["introduction"]
        self.assertIn("First intro paragraph.", intro.text)
        self.assertIn("Second intro paragraph.", intro.text)
        # Start index = 2 (I. heading is at para 1), end index = 3
        self.assertEqual(intro.start_para_index, 2)
        self.assertEqual(intro.end_para_index, 3)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to confirm they fail**

Run: `python -m pytest tests/test_mediation_brief_live.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.mediation_brief_live'`.

- [ ] **Step 3: Create `icharlotte_core/mediation_brief_live.py`**

```python
"""Live-document utilities for mediation briefs.

Parses an open Word document into a :class:`LiveBrief` and provides helpers
for locating section ranges and inserting formatted quote blocks.  Used by
the Word AI assistant popup (Win+V) so refinement and quote insertion can
run against the live document without depending on in-memory generator
state.

Only the parser and range helpers live here.  Quote insertion is added in
a separate task.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from icharlotte_core.mediation_brief import (
    _HEADING_PATTERN,
    _HEADING_TO_SECTION,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


@dataclass
class LiveSection:
    """A single section of a mediation brief parsed out of a live Word doc.

    Attributes:
        name: Canonical section name (e.g. ``"liability"``).
        heading_title: Heading text as it appears in the document (e.g.
            ``"IV. LIABILITY"``).
        text: Body text of the section — excludes the heading paragraph.
        start_para_index: 1-based index of the first body paragraph in
            ``doc.Paragraphs``.
        end_para_index: 1-based index of the last body paragraph in
            ``doc.Paragraphs``.
    """

    name: str
    heading_title: str
    text: str
    start_para_index: int
    end_para_index: int


@dataclass
class LiveBrief:
    """The parsed result of walking a live Word document for a brief."""

    doc_path: str
    sections: Dict[str, LiveSection] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Heading detection
# ---------------------------------------------------------------------------


def _match_heading(text: str) -> Optional[str]:
    """If *text* is a brief section heading, return its canonical name.

    Returns None if the paragraph text is not a recognised heading.
    """
    stripped = text.strip()
    if not stripped:
        return None
    m = _HEADING_PATTERN.match(stripped)
    if not m:
        return None
    heading_title = m.group(2).strip()
    canonical = _HEADING_TO_SECTION.get(heading_title)
    if canonical is not None:
        return canonical
    # Partial match fallback — the existing parser in mediation_brief.py
    # does the same thing.
    for variant, name in _HEADING_TO_SECTION.items():
        if heading_title.startswith(variant) or variant in heading_title:
            return name
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def parse_brief_from_word_doc(doc_com) -> LiveBrief:
    """Walk *doc_com*'s paragraphs and return a :class:`LiveBrief`.

    Recognises roman-numeral section headings via the same pattern and
    heading map that :mod:`icharlotte_core.mediation_brief` uses.
    Unrecognised headings are skipped silently.

    The returned ``sections`` dict is keyed by canonical section name; for
    each section, ``text`` is the concatenation of all body paragraphs
    between that heading and the next recognised heading (or end of doc).
    """
    doc_path = getattr(doc_com, "FullName", "") or ""
    live = LiveBrief(doc_path=doc_path)

    # Walk paragraphs once, tracking the current section.
    current_name: Optional[str] = None
    current_heading_title: str = ""
    current_body: List[str] = []
    current_start: int = 0
    current_end: int = 0

    def _commit():
        if current_name and current_name not in live.sections:
            live.sections[current_name] = LiveSection(
                name=current_name,
                heading_title=current_heading_title,
                text="\n".join(current_body).strip(),
                start_para_index=current_start,
                end_para_index=current_end,
            )

    paragraphs = doc_com.Paragraphs
    count = paragraphs.Count
    for idx in range(1, count + 1):
        para = paragraphs(idx)
        raw = para.Range.Text or ""
        # Word COM appends \r (and sometimes \x07 for table markers). Strip.
        text = raw.rstrip("\r\n\x07 \t")

        canonical = _match_heading(text)
        if canonical is not None:
            _commit()
            current_name = canonical
            current_heading_title = text.strip()
            current_body = []
            current_start = idx + 1
            current_end = idx  # will be updated when body paragraphs arrive
            continue

        if current_name is not None:
            if text.strip():
                current_body.append(text)
            current_end = idx

    _commit()
    return live


def is_mediation_brief(doc_com) -> bool:
    """Return True if *doc_com* contains at least 3 recognised brief sections.

    Used by the Word popup to gate the "Mediation Brief" template entries so
    they only appear when the active document looks like a brief.
    """
    try:
        live = parse_brief_from_word_doc(doc_com)
    except Exception as e:
        logger.debug("is_mediation_brief: parse failed: %s", e)
        return False
    return len(live.sections) >= 3
```

- [ ] **Step 4: Run the tests and confirm they pass**

Run: `python -m pytest tests/test_mediation_brief_live.py -v`
Expected: all 10 tests pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief_live.py tests/test_mediation_brief_live.py
git commit -m "feat(mediation-brief-live): parser and is_mediation_brief detection"
```

---

## Task 3: `get_word_range_for_section` helper

**Files:**
- Modify: `icharlotte_core/mediation_brief_live.py`
- Test: `tests/test_mediation_brief_live.py`

**Context for the engineer:**
- Given a `LiveSection` with `start_para_index` / `end_para_index`, we need a Word `Range` object spanning from the start of the first body paragraph to the end of the last body paragraph (inclusive of its trailing paragraph mark — so `InsertFile`/replacement operations act on the full section).
- Word COM: `doc.Range(start_char, end_char)` returns a Range from absolute character positions. `doc.Paragraphs(i).Range.Start` / `.End` give those positions. We use `Start` of the first body paragraph and `End` of the last body paragraph.
- Edge case: if `end_para_index < start_para_index`, the body is empty (heading with nothing under it). In that case the range is a collapsed caret at the paragraph right after the heading — build a zero-length range at `doc.Paragraphs(start_para_index - 1).Range.End` (end of the heading paragraph).

- [ ] **Step 1: Write the failing test**

Append these tests to `tests/test_mediation_brief_live.py` (before the `if __name__ == "__main__":` block):

```python
class TestGetWordRangeForSection(unittest.TestCase):
    def _make_doc(self):
        return make_doc(
            "I.   INTRODUCTION",
            "Intro para one.",
            "Intro para two.",
            "II.  STATEMENT OF FACTS",
            "Facts.",
            "III. LIABILITY",
            "Liability.",
        )

    def _capture_range_calls(self, fake_doc):
        calls = []

        def _Range(start, end):
            calls.append((start, end))
            return ("range", start, end)

        fake_doc.Range = _Range
        return calls

    def test_range_spans_all_body_paragraphs(self):
        from icharlotte_core.mediation_brief_live import (
            get_word_range_for_section,
            parse_brief_from_word_doc,
        )
        doc = self._make_doc()
        calls = self._capture_range_calls(doc)
        live = parse_brief_from_word_doc(doc)

        intro = live.sections["introduction"]
        get_word_range_for_section(doc, intro)

        self.assertEqual(len(calls), 1)
        start_char, end_char = calls[0]
        expected_start = doc.Paragraphs(intro.start_para_index).Range.Start
        expected_end = doc.Paragraphs(intro.end_para_index).Range.End
        self.assertEqual(start_char, expected_start)
        self.assertEqual(end_char, expected_end)

    def test_empty_body_collapses_to_caret_after_heading(self):
        from icharlotte_core.mediation_brief_live import (
            LiveSection,
            get_word_range_for_section,
        )
        doc = make_doc("I.   INTRODUCTION", "II.  STATEMENT OF FACTS", "Facts.")
        calls = self._capture_range_calls(doc)
        # Simulate an "introduction" section with no body paragraphs:
        # start_para_index=2 (would be the next body para) but end_para_index=1
        # (still pointing at the heading paragraph, because no body landed).
        empty_intro = LiveSection(
            name="introduction",
            heading_title="I. INTRODUCTION",
            text="",
            start_para_index=2,
            end_para_index=1,
        )
        get_word_range_for_section(doc, empty_intro)
        self.assertEqual(len(calls), 1)
        start_char, end_char = calls[0]
        heading_end = doc.Paragraphs(1).Range.End
        self.assertEqual(start_char, heading_end)
        self.assertEqual(end_char, heading_end)
```

- [ ] **Step 2: Run the tests to confirm they fail**

Run: `python -m pytest tests/test_mediation_brief_live.py::TestGetWordRangeForSection -v`
Expected: FAIL with `ImportError: cannot import name 'get_word_range_for_section'`.

- [ ] **Step 3: Add `get_word_range_for_section` to `mediation_brief_live.py`**

Append to `icharlotte_core/mediation_brief_live.py` (after `is_mediation_brief`):

```python
def get_word_range_for_section(doc_com, section: LiveSection):
    """Return a Word ``Range`` object covering *section*'s body paragraphs.

    The range runs from the start of the first body paragraph to the end of
    the last body paragraph (inclusive of its trailing paragraph mark).

    If the section has no body paragraphs (heading followed directly by the
    next heading), returns a zero-length range at the end of the heading
    paragraph — suitable as an insertion point.
    """
    if section.end_para_index < section.start_para_index:
        # Empty body — caret at the end of the heading paragraph.
        heading_idx = section.start_para_index - 1
        if heading_idx < 1:
            heading_idx = 1
        heading_para = doc_com.Paragraphs(heading_idx)
        pos = heading_para.Range.End
        return doc_com.Range(pos, pos)

    first = doc_com.Paragraphs(section.start_para_index)
    last = doc_com.Paragraphs(section.end_para_index)
    return doc_com.Range(first.Range.Start, last.Range.End)
```

- [ ] **Step 4: Run the tests and confirm they pass**

Run: `python -m pytest tests/test_mediation_brief_live.py -v`
Expected: all tests in the file pass (new and existing).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief_live.py tests/test_mediation_brief_live.py
git commit -m "feat(mediation-brief-live): get_word_range_for_section helper"
```

---

## Task 4: `insert_formatted_quotes_at_range` helper

**Files:**
- Modify: `icharlotte_core/mediation_brief_live.py`
- Test: `tests/test_mediation_brief_live.py`

**Context for the engineer:**
- The chat-tab flow formats depo quotes via `MediationBriefGenerator._add_depo_quote(doc, text)` at `icharlotte_core/mediation_brief.py:923`. It takes a python-docx `Document` and a block of text that looks like:

  ```
  Q. Were you paying attention?
  A. No.
  (Smith Depo Trns., at p. 45:12.)
  ```

  and produces single-spaced Q/A paragraphs with the hanging indent, followed by the citation line with `space_before=12pt`.

- Our helper builds a temp python-docx document, calls `_add_depo_quote` for each selected quote, saves the temp file, then calls Word COM `range_com.InsertFile(tmp_path)` at the cursor/selection.
- Each quote dict looks like: `{"deponent": "Smith", "page_line": "45:12", "qa_text": "Q. ...\nA. ...", ...}` (matches `_parse_quote_results` output at `icharlotte_core/mediation_brief.py:1207`).
- For the temp doc we need to assemble a text block per quote matching the format `insert_quotes_quick` uses at `mediation_brief.py:1277`:

  ```
  <qa_text>
  (<deponent> Depo Trns., at p. <page_line>.)
  ```

- `Range.InsertFile` syntax (COM): `range.InsertFile(FileName=path, ConfirmConversions=False)`. Need to pass `str` to `FileName`.
- Temp file cleanup: use `tempfile.NamedTemporaryFile(suffix=".docx", delete=False)` + `os.unlink()` in `finally`. On Windows the file must be closed before Word can open it.
- Test strategy: unit-test builds a real temp python-docx and asserts paragraph count + content. The Word `InsertFile` call is mocked through a fake range object that records the call.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_mediation_brief_live.py`:

```python
class TestInsertFormattedQuotesAtRange(unittest.TestCase):
    def _quote(self, deponent, qa, page_line):
        return {
            "deponent": deponent,
            "source": f"{deponent}.pdf",
            "page_line": page_line,
            "relevance": "test",
            "qa_text": qa,
        }

    def test_builds_temp_docx_and_calls_insertfile(self):
        from icharlotte_core.mediation_brief_live import (
            insert_formatted_quotes_at_range,
        )

        insert_calls = []

        class FakeRange:
            def InsertFile(self, FileName, **kwargs):
                insert_calls.append(FileName)
                # Confirm file actually exists at call time — after the call
                # returns the helper deletes it.
                import os
                assert os.path.isfile(FileName), f"temp file missing: {FileName}"

        quotes = [
            self._quote("Smith", "Q. Did you see the light?\nA. Yes.", "45:12"),
            self._quote("Jones", "Q. Were you paying attention?\nA. No.", "12:3"),
        ]
        insert_formatted_quotes_at_range(doc_com=object(), range_com=FakeRange(), quotes=quotes)

        self.assertEqual(len(insert_calls), 1)

    def test_temp_docx_contains_all_quotes_with_citations(self):
        import os
        import tempfile
        from docx import Document as DocxDocument
        from icharlotte_core.mediation_brief_live import (
            insert_formatted_quotes_at_range,
        )

        captured_path = {"path": None}

        class CopyingRange:
            def InsertFile(self, FileName, **kwargs):
                # Copy the temp file to a location we control so we can
                # inspect it after the helper deletes the original.
                import shutil
                dst = tempfile.NamedTemporaryFile(suffix=".docx", delete=False).name
                shutil.copy2(FileName, dst)
                captured_path["path"] = dst

        quotes = [
            self._quote("Smith", "Q. Did you see the light?\nA. Yes.", "45:12"),
            self._quote("Jones", "Q. Were you paying attention?\nA. No.", "12:3"),
        ]
        insert_formatted_quotes_at_range(doc_com=object(), range_com=CopyingRange(), quotes=quotes)

        try:
            self.assertIsNotNone(captured_path["path"])
            tdoc = DocxDocument(captured_path["path"])
            all_text = "\n".join(p.text for p in tdoc.paragraphs)
            self.assertIn("Did you see the light", all_text)
            self.assertIn("Yes.", all_text)
            self.assertIn("Were you paying attention", all_text)
            self.assertIn("(Smith Depo Trns., at p. 45:12.)", all_text)
            self.assertIn("(Jones Depo Trns., at p. 12:3.)", all_text)
        finally:
            if captured_path["path"] and os.path.isfile(captured_path["path"]):
                os.unlink(captured_path["path"])

    def test_temp_file_is_deleted_after_insert(self):
        from icharlotte_core.mediation_brief_live import (
            insert_formatted_quotes_at_range,
        )
        import os

        captured_path = {"path": None}

        class CapturingRange:
            def InsertFile(self, FileName, **kwargs):
                captured_path["path"] = FileName

        quotes = [self._quote("Smith", "Q. Did you see?\nA. Yes.", "45:12")]
        insert_formatted_quotes_at_range(doc_com=object(), range_com=CapturingRange(), quotes=quotes)

        self.assertIsNotNone(captured_path["path"])
        self.assertFalse(os.path.isfile(captured_path["path"]))

    def test_temp_file_is_deleted_even_on_insert_error(self):
        from icharlotte_core.mediation_brief_live import (
            insert_formatted_quotes_at_range,
        )
        import os

        captured_path = {"path": None}

        class FailingRange:
            def InsertFile(self, FileName, **kwargs):
                captured_path["path"] = FileName
                raise RuntimeError("COM failure")

        quotes = [self._quote("Smith", "Q. Did you see?\nA. Yes.", "45:12")]
        with self.assertRaises(RuntimeError):
            insert_formatted_quotes_at_range(doc_com=object(), range_com=FailingRange(), quotes=quotes)

        self.assertIsNotNone(captured_path["path"])
        self.assertFalse(os.path.isfile(captured_path["path"]))

    def test_empty_quote_list_is_noop(self):
        from icharlotte_core.mediation_brief_live import (
            insert_formatted_quotes_at_range,
        )

        class FakeRange:
            def __init__(self):
                self.called = False

            def InsertFile(self, FileName, **kwargs):
                self.called = True

        fake = FakeRange()
        insert_formatted_quotes_at_range(doc_com=object(), range_com=fake, quotes=[])
        self.assertFalse(fake.called)
```

- [ ] **Step 2: Run the tests to confirm they fail**

Run: `python -m pytest tests/test_mediation_brief_live.py::TestInsertFormattedQuotesAtRange -v`
Expected: FAIL with `ImportError: cannot import name 'insert_formatted_quotes_at_range'`.

- [ ] **Step 3: Add `insert_formatted_quotes_at_range` to `mediation_brief_live.py`**

Append to `icharlotte_core/mediation_brief_live.py`:

```python
import os
import tempfile

from docx import Document as DocxDocument

from icharlotte_core.mediation_brief import MediationBriefGenerator


def _format_quote_block_text(quote: Dict) -> str:
    """Assemble a single quote block in the same format as the chat-tab flow.

    Matches the format produced by
    :meth:`MediationBriefGenerator.insert_quotes_quick` — a Q&A block
    followed by the citation on the next line, without the
    ``DEPO_QUOTE_START``/``DEPO_QUOTE_END`` markers (which are consumed by
    the section-text parser, not rendered).
    """
    qa = (quote.get("qa_text") or "").strip()
    deponent = (quote.get("deponent") or "").strip()
    page_line = (quote.get("page_line") or "").strip()
    citation = f"({deponent} Depo Trns., at p. {page_line}.)"
    return f"{qa}\n{citation}"


def insert_formatted_quotes_at_range(doc_com, range_com, quotes: List[Dict]) -> None:
    """Insert *quotes* as formatted Q&A blocks at *range_com*.

    Builds a temporary .docx containing the formatted quote paragraphs using
    :meth:`MediationBriefGenerator._add_depo_quote` — the same formatter the
    chat-tab flow uses — then calls Word COM ``Range.InsertFile`` to splice
    that content into the live document at the given range.

    The temporary file is always deleted, even if ``InsertFile`` raises.

    Args:
        doc_com: The Word COM ``Document`` (currently unused — kept for
            future hook points and to make the call site explicit).
        range_com: A Word COM ``Range`` — the insertion point / replacement
            target.
        quotes: List of quote dicts as produced by
            :meth:`MediationBriefGenerator.search_quotes`.

    Returns:
        None.
    """
    if not quotes:
        return

    # Build the temp docx.
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        docx_doc = DocxDocument()
        generator = MediationBriefGenerator()
        for quote in quotes:
            block_text = _format_quote_block_text(quote)
            generator._add_depo_quote(docx_doc, block_text)
        docx_doc.save(tmp_path)

        # Insert into the live Word document.
        range_com.InsertFile(FileName=str(tmp_path), ConfirmConversions=False)
    finally:
        try:
            if os.path.isfile(tmp_path):
                os.unlink(tmp_path)
        except OSError as e:
            logger.warning("Failed to delete temp quote docx %s: %s", tmp_path, e)
```

- [ ] **Step 4: Run the tests and confirm they pass**

Run: `python -m pytest tests/test_mediation_brief_live.py -v`
Expected: all tests in the file pass.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief_live.py tests/test_mediation_brief_live.py
git commit -m "feat(mediation-brief-live): insert_formatted_quotes_at_range helper"
```

---

## Task 5: `WordQuoteInsertionDialog` (stripped dialog)

**Files:**
- Create: `icharlotte_core/ui/quote_dialog_word.py`
- Modify: (none — this is purely a new file that imports shared widgets from `icharlotte_core/ui/quote_dialog.py`)

**Context for the engineer:**
- Existing dialog: `icharlotte_core/ui/quote_dialog.py` — class `QuoteInsertionDialog`.
- It uses these internal classes that we'll reuse: `QuoteSearchWorker` (QThread that calls `MediationBriefGenerator().search_quotes`) and `QuoteResultWidget` (per-result UI with a checkbox + edit toggle).
- Read lines 170–400 of the existing dialog first to see the exact widget construction for the transcript list, search description, results scroll area, and signal wiring. Copy that construction into the new dialog; remove the section combo, subsection combo, Quick/Weave radio group, and the generator parameter.
- The new signal: `quotes_to_insert = Signal(list)` — emitted with a plain list of quote dicts.
- The new dialog's constructor takes only `parent=None` and instantiates its own throwaway `MediationBriefGenerator` inside `_start_search` to pass to `QuoteSearchWorker`.
- No unit test for this class (it's UI). Manual integration test in Task 10 covers it.

- [ ] **Step 1: Read the existing dialog to understand what to copy**

Run: `cat icharlotte_core/ui/quote_dialog.py | head -400`
Identify: `QuoteSearchWorker` (top of file), `QuoteResultWidget` (mid), `QuoteInsertionDialog.__init__`'s transcript/description/results construction. Note which widgets are part of placement controls (section combo, subsection combo, mode radios) — those will NOT be copied.

- [ ] **Step 2: Create `icharlotte_core/ui/quote_dialog_word.py`**

```python
"""WordQuoteInsertionDialog — Word-cursor variant of QuoteInsertionDialog.

Used by the Word AI assistant popup's "Mediation Brief: Add Quotes" flow.
Provides the same transcript upload + search description + result selection
UI as :class:`QuoteInsertionDialog`, but without the section/subsection
combos or the Quick/Weave mode toggle — quotes are inserted at the Word
cursor, not into a named section.

Reuses :class:`QuoteSearchWorker` and :class:`QuoteResultWidget` from the
existing dialog module so search behaviour and per-result rendering stay in
a single place.
"""

from __future__ import annotations

from typing import Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.mediation_brief import MediationBriefGenerator
from icharlotte_core.ui.quote_dialog import QuoteResultWidget, QuoteSearchWorker


class WordQuoteInsertionDialog(QDialog):
    """Modal dialog for searching depo transcripts and picking quotes to
    insert at the current Word cursor.

    Emits ``quotes_to_insert(quotes: List[Dict])`` when the user clicks
    "Insert Selected".  The caller is responsible for splicing the quotes
    into the live Word document.
    """

    quotes_to_insert = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker: Optional[QuoteSearchWorker] = None
        self._result_widgets: List[QuoteResultWidget] = []

        self.setWindowTitle("Insert Deposition Quotes")
        self.setMinimumWidth(640)
        self.setMinimumHeight(640)

        root = QVBoxLayout(self)
        root.setSpacing(8)

        # ── 1. Transcript upload ─────────────────────────────────────────
        transcript_group = QGroupBox("Transcripts")
        tg_layout = QVBoxLayout(transcript_group)

        self.transcript_list = QListWidget()
        self.transcript_list.setMaximumHeight(100)
        self.transcript_list.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection
        )
        tg_layout.addWidget(self.transcript_list)

        t_btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Transcript(s)")
        add_btn.clicked.connect(self._add_transcripts)
        t_btn_row.addWidget(add_btn)

        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_selected_transcripts)
        t_btn_row.addWidget(remove_btn)
        t_btn_row.addStretch()
        tg_layout.addLayout(t_btn_row)

        root.addWidget(transcript_group)

        # ── 2. Search description ────────────────────────────────────────
        desc_group = QGroupBox("Search Description")
        desc_layout = QVBoxLayout(desc_group)
        self.desc_edit = QPlainTextEdit()
        self.desc_edit.setPlaceholderText(
            "Describe the testimony you are looking for, e.g.:\n"
            "'plaintiff admits she did not seek medical treatment for six months'"
        )
        self.desc_edit.setMaximumHeight(80)
        self.desc_edit.textChanged.connect(self._update_search_button)
        desc_layout.addWidget(self.desc_edit)
        root.addWidget(desc_group)

        # ── 3. Search button + progress ──────────────────────────────────
        search_row = QHBoxLayout()
        self.search_btn = QPushButton("Search")
        self.search_btn.setEnabled(False)
        self.search_btn.clicked.connect(self._start_search)
        search_row.addWidget(self.search_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(18)
        search_row.addWidget(self.progress_bar, 1)
        root.addLayout(search_row)

        # ── 4. Results panel ─────────────────────────────────────────────
        results_group = QGroupBox("Results")
        rg_layout = QVBoxLayout(results_group)

        self.status_label = QLabel(
            "Add transcripts and enter a search description to begin."
        )
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        rg_layout.addWidget(self.status_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._results_container = QWidget()
        self._results_layout = QVBoxLayout(self._results_container)
        self._results_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll.setWidget(self._results_container)
        rg_layout.addWidget(scroll, 1)

        root.addWidget(results_group, 1)

        # ── 5. Action buttons ────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        self.insert_btn = QPushButton("Insert Selected")
        self.insert_btn.setEnabled(False)
        self.insert_btn.clicked.connect(self._on_insert_clicked)
        btn_row.addWidget(self.insert_btn)

        root.addLayout(btn_row)

    # ------------------------------------------------------------------
    # Transcript management
    # ------------------------------------------------------------------

    def _add_transcripts(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select deposition transcript(s)",
            "",
            "Transcripts (*.pdf *.docx);;All Files (*)",
        )
        for path in paths:
            if not any(
                self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole) == path
                for i in range(self.transcript_list.count())
            ):
                item = QListWidgetItem(path)
                item.setData(Qt.ItemDataRole.UserRole, path)
                self.transcript_list.addItem(item)
        self._update_search_button()

    def _remove_selected_transcripts(self):
        for item in self.transcript_list.selectedItems():
            self.transcript_list.takeItem(self.transcript_list.row(item))
        self._update_search_button()

    def _current_transcript_paths(self) -> List[str]:
        return [
            self.transcript_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.transcript_list.count())
        ]

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def _update_search_button(self):
        has_transcripts = self.transcript_list.count() > 0
        has_description = bool(self.desc_edit.toPlainText().strip())
        self.search_btn.setEnabled(has_transcripts and has_description)

    def _start_search(self):
        paths = self._current_transcript_paths()
        description = self.desc_edit.toPlainText().strip()
        if not paths or not description:
            return

        self._clear_results()
        self.progress_bar.setVisible(True)
        self.search_btn.setEnabled(False)
        self.status_label.setText("Searching transcripts…")

        generator = MediationBriefGenerator()
        self._worker = QuoteSearchWorker(generator, paths, description, parent=self)
        self._worker.finished_with_results.connect(self._on_search_done)
        self._worker.error.connect(self._on_search_error)
        self._worker.start()

    def _clear_results(self):
        while self._results_layout.count():
            item = self._results_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()
        self._result_widgets = []
        self.insert_btn.setEnabled(False)

    def _on_search_done(self, results: List[Dict]):
        self.progress_bar.setVisible(False)
        self.search_btn.setEnabled(True)
        if not results:
            self.status_label.setText("No matches found. Try revising your description.")
            return
        self.status_label.setText(f"Found {len(results)} result(s). Select quotes to insert.")
        for r in results:
            w = QuoteResultWidget(r, parent=self._results_container)
            w.checkbox.toggled.connect(self._update_insert_button)
            self._results_layout.addWidget(w)
            self._result_widgets.append(w)
        self._update_insert_button()

    def _on_search_error(self, msg: str):
        self.progress_bar.setVisible(False)
        self.search_btn.setEnabled(True)
        self.status_label.setText(f"Search failed: {msg}")

    def _update_insert_button(self):
        any_checked = any(w.is_selected() for w in self._result_widgets)
        self.insert_btn.setEnabled(any_checked)

    def _on_insert_clicked(self):
        selected = [w.get_quote_data() for w in self._result_widgets if w.is_selected()]
        if not selected:
            return
        self.quotes_to_insert.emit(selected)
        self.accept()
```

- [ ] **Step 3: Smoke-test the dialog imports cleanly**

Run: `python -c "from icharlotte_core.ui.quote_dialog_word import WordQuoteInsertionDialog; print('OK')"`
Expected: `OK`.

If the import fails, check that `QuoteResultWidget` and `QuoteSearchWorker` are actually defined in `icharlotte_core/ui/quote_dialog.py` with those exact names. If a signal on `QuoteSearchWorker` is named differently (e.g. `results` vs `finished_with_results`), update `_start_search` to connect the correct signal name — open `quote_dialog.py` and search for `Signal` in the `QuoteSearchWorker` class.

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/ui/quote_dialog_word.py
git commit -m "feat(ui): WordQuoteInsertionDialog — stripped dialog for Word cursor insertion"
```

---

## Task 6: Word popup — sentinel constants + dynamic dropdown entries on open

**Files:**
- Modify: `icharlotte_core/word_hotkey.py`

**Context for the engineer:**
- The popup's prompt dropdown is `self.prompt_combo` (a `QComboBox`), populated by `refresh_combo()` at `word_hotkey.py:3088`. Each item's `userData` is either `None` (for the "-- Select a saved prompt --" header) or a prompt string.
- We'll add two new entries whose `userData` is a sentinel string (not a prompt): `"__MB_REFINE__"` and `"__MB_ADD_QUOTES__"`. These sentinels are what `on_prompt_selected` and `_do_execute` will check for.
- The dropdown should only show these entries when the popup's captured document (`self._original_document`) looks like a brief. Add this logic to the end of `refresh_combo()`.
- `_original_document` is set in the popup's show-path before `refresh_combo` runs. If it's `None` (popup opened without a doc), skip the check.
- Word COM calls can throw if the document reference is stale — wrap the `is_mediation_brief` call in a try/except and treat any failure as "not a brief".
- No test for this task — it touches Qt widgets and Word COM. Task 10's integration checklist covers verification.

- [ ] **Step 1: Add sentinel constants near the top of `word_hotkey.py`**

Open `icharlotte_core/word_hotkey.py`. Find the FORMAT constants (around line 570, above the `TaskData` dataclass). Add after the FORMAT constants:

```python
# Sentinel userData values for the Word popup's prompt dropdown. These flag
# the two Mediation-Brief-specific dropdown entries so _do_execute can branch
# to the live-brief paths instead of running a normal LLM task.
MB_REFINE_SENTINEL = "__MB_REFINE__"
MB_ADD_QUOTES_SENTINEL = "__MB_ADD_QUOTES__"
```

- [ ] **Step 2: Extend `refresh_combo` to add brief entries when applicable**

Find `refresh_combo` at `word_hotkey.py:3088`. Replace the existing method body with:

```python
    def refresh_combo(self):
        """Refresh the combo box with prompts for current context."""
        self.prompt_combo.clear()
        self.prompt_combo.addItem("-- Select a saved prompt --", None)

        # Get prompts for current context
        if self.app_context == APP_CONTEXT_OUTLOOK:
            prompts = self.outlook_prompts
        else:
            prompts = self.prompts

        for p in prompts:
            self.prompt_combo.addItem(p["name"], p["prompt"])

        # Add Mediation Brief entries if the active Word document is a brief.
        if self.app_context != APP_CONTEXT_OUTLOOK and self._is_active_doc_a_brief():
            self.prompt_combo.insertSeparator(self.prompt_combo.count())
            self.prompt_combo.addItem(
                "Mediation Brief: Refine Section", MB_REFINE_SENTINEL
            )
            self.prompt_combo.addItem(
                "Mediation Brief: Add Quotes", MB_ADD_QUOTES_SENTINEL
            )

    def _is_active_doc_a_brief(self) -> bool:
        """Return True if the popup's captured Word document is a mediation brief."""
        doc = getattr(self, "_original_document", None)
        if doc is None:
            return False
        try:
            from icharlotte_core.mediation_brief_live import is_mediation_brief
            return is_mediation_brief(doc)
        except Exception as e:
            print(f"[WordLLMPopup] _is_active_doc_a_brief failed: {e}")
            return False
```

- [ ] **Step 3: Ensure `refresh_combo` runs AFTER `_original_document` is captured**

Find where the popup is shown / brought to front in response to the hotkey. Search for `_original_document =` and `refresh_combo(` to confirm the call order. You want `refresh_combo` to run after `_original_document` is set. If `refresh_combo` currently runs only in `load_prompts` (which runs in `__init__` before the document is captured), add an explicit `self.refresh_combo()` call at the end of the method that captures the active document (commonly `_capture_active_document` or similar, or the `showEvent`). Search the file for the method that assigns `self._original_document = ...` and append `self.refresh_combo()` after that assignment.

Do NOT remove the `refresh_combo` call from `load_prompts` — keep both, since the first call populates the user's saved prompts at init and the second refreshes after doc capture.

- [ ] **Step 4: Run an import sanity check**

Run: `python -c "import icharlotte_core.word_hotkey; print('OK')"`
Expected: `OK`.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat(word-popup): sentinel constants and dynamic MB dropdown entries"
```

---

## Task 7: Word popup — inline section combo, visibility wired to dropdown selection

**Files:**
- Modify: `icharlotte_core/word_hotkey.py`

**Context for the engineer:**
- We need a `QComboBox` for picking a canonical section name. Populated from `SECTION_ORDER` + `SECTION_HEADINGS` from `icharlotte_core/mediation_brief.py`.
- The combo is added to the AI Prompt tab's layout right above (or below) the custom prompt input, hidden by default. Shown only when the user selects the "Mediation Brief: Refine Section" entry. Hidden when any other entry is selected.
- The existing method `on_prompt_selected` at `word_hotkey.py:3120` is the dropdown change handler. Extend it to show/hide the section combo and to NOT overwrite `custom_input` with the sentinel string.

- [ ] **Step 1: Add the section combo widget in `_setup_ai_prompt_tab`**

Find `_setup_ai_prompt_tab` in `word_hotkey.py` (around line 2580). Right after the `custom_input` is added to `ai_layout` (the `ai_layout.addWidget(self.custom_input)` line around 2618), insert:

```python
        # Mediation Brief section picker (hidden by default — shown only when
        # the user selects "Mediation Brief: Refine Section" from the
        # prompt dropdown).
        from icharlotte_core.mediation_brief import SECTION_ORDER, SECTION_HEADINGS
        self.mb_section_row = QWidget()
        mb_section_layout = QHBoxLayout(self.mb_section_row)
        mb_section_layout.setContentsMargins(0, 0, 0, 0)
        mb_section_layout.addWidget(QLabel("Section:"))
        self.mb_section_combo = QComboBox()
        for name in SECTION_ORDER:
            _, display = SECTION_HEADINGS[name]
            self.mb_section_combo.addItem(display, name)
        mb_section_layout.addWidget(self.mb_section_combo, 1)
        self.mb_section_row.setVisible(False)
        ai_layout.addWidget(self.mb_section_row)
```

- [ ] **Step 2: Update `on_prompt_selected` to handle sentinel values**

Find `on_prompt_selected` at `word_hotkey.py:3120`. Replace its body with:

```python
    def on_prompt_selected(self, index):
        """When a saved prompt is selected, populate the custom input.

        When one of the Mediation Brief sentinel entries is selected, show
        or hide the section-picker row instead of overwriting the prompt
        input.
        """
        data = self.prompt_combo.currentData()

        # Mediation Brief entries carry sentinel strings, not real prompts.
        if data == MB_REFINE_SENTINEL:
            if hasattr(self, "mb_section_row"):
                self.mb_section_row.setVisible(True)
            return
        if data == MB_ADD_QUOTES_SENTINEL:
            if hasattr(self, "mb_section_row"):
                self.mb_section_row.setVisible(False)
            return

        # Regular prompt — populate the custom input and hide the MB section row.
        if hasattr(self, "mb_section_row"):
            self.mb_section_row.setVisible(False)
        if data:
            self.custom_input.setPlainText(data)
```

- [ ] **Step 3: Run the import sanity check**

Run: `python -c "import icharlotte_core.word_hotkey; print('OK')"`
Expected: `OK`.

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat(word-popup): inline MB section picker and dropdown branching"
```

---

## Task 8: Word popup — `_do_execute` dispatch for "Mediation Brief: Refine Section"

**Files:**
- Modify: `icharlotte_core/word_hotkey.py`

**Context for the engineer:**
- `_do_execute` at `word_hotkey.py:3577` is the main dispatch. Near the top it branches for Outlook; we'll add a branch above the normal Word path that checks `self.prompt_combo.currentData()`.
- For refine:
  1. Call `parse_brief_from_word_doc(self._original_document)`.
  2. Look up the picked section in `live.sections`. If missing — show a QMessageBox warning, clean up the status-bar preparing row, and return.
  3. Build `sections_dict = {name: sec.text for name, sec in live.sections.items()}`.
  4. Call `get_word_range_for_section(doc, target_section)` → Word range object.
  5. Read `range.Start` and `range.End` (absolute character positions) to pass into `_create_task_bookmark`.
  6. Instantiate a throwaway `MediationBriefGenerator`, call `build_refinement_prompts` to get the system + full prompts.
  7. Build a `TaskData` with those prompts, the section range as the bookmark, and submit to `TaskManager`.
- The prompt the user types into `custom_input` is the refinement instruction.
- Existing `_do_execute` code for capturing format, `provider`/`model_id`, redline mode etc. needs to still run. The cleanest place to insert the branch is right after `word`/`format_type` are resolved and BEFORE the normal "get selected text / build prompt from template" block.

- [ ] **Step 1: Add a helper method `_handle_mb_refine`**

Add this method to `WordLLMPopup` class (place it near `_do_execute`):

```python
    def _handle_mb_refine(self, instruction: str) -> bool:
        """Dispatch for "Mediation Brief: Refine Section".

        Returns True if the refine task was submitted (caller should stop
        further processing in _do_execute); False if it fell through because
        prerequisites were not met (caller should NOT fall through — we
        treat prerequisite failure as a hard stop and the status bar row
        is cleaned up here before returning False).
        """
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        from icharlotte_core.mediation_brief_live import (
            get_word_range_for_section,
            parse_brief_from_word_doc,
        )

        section_name = self.mb_section_combo.currentData()
        if not section_name:
            QMessageBox.warning(self, "No section", "Please pick a section to refine.")
            self._clear_preparing_row()
            return False

        doc = self._original_document
        if doc is None:
            QMessageBox.warning(
                self, "Word Not Found",
                "No active Word document — please open the brief first."
            )
            self._clear_preparing_row()
            return False

        try:
            live = parse_brief_from_word_doc(doc)
        except Exception as e:
            QMessageBox.warning(self, "Parse failed", f"Could not read brief: {e}")
            self._clear_preparing_row()
            return False

        target = live.sections.get(section_name)
        if target is None:
            display = self.mb_section_combo.currentText()
            QMessageBox.warning(
                self, "Section not found",
                f"The '{display}' section was not found in this document."
            )
            self._clear_preparing_row()
            return False

        sections_dict = {name: sec.text for name, sec in live.sections.items()}

        generator = MediationBriefGenerator()
        try:
            system_prompt, full_prompt = generator.build_refinement_prompts(
                section_name=section_name,
                sections_dict=sections_dict,
                instruction=instruction,
            )
        except Exception as e:
            QMessageBox.warning(self, "Prompt build failed", str(e))
            self._clear_preparing_row()
            return False

        # Resolve the Word range for the target section and read its bounds
        # for bookmark creation.
        try:
            target_range = get_word_range_for_section(doc, target)
            range_start = int(target_range.Start)
            range_end = int(target_range.End)
        except Exception as e:
            QMessageBox.warning(
                self, "Range lookup failed",
                f"Could not locate section range in document: {e}"
            )
            self._clear_preparing_row()
            return False

        # Build TaskData. Reuse the existing selected-model + redline handling.
        provider, model_id = self._get_selected_model()
        task_data = TaskData(
            document_name=self._original_document_name or "",
            document_com=doc,
            has_selection=True,
            original_text=target.text,
            original_text_raw=target.text,
            captured_format=self._captured_format,
            format_type=getattr(self, "_exec_format_type", FORMAT_PLAIN),
            redline_mode_active=self._redline_mode_active,
            redline_settings=self.redline_settings.copy() if self.redline_settings else {},
            research_result=None,
            do_legal_research=False,
            legal_research_engine=None,
            legal_research_model=None,
            prompt_preview=f"MB Refine {section_name}: {instruction[:40]}",
            provider=provider,
            model_id=model_id,
            system_prompt=system_prompt,
            full_prompt=full_prompt,
        )

        bm_name = _create_task_bookmark(doc, range_start, range_end, task_data.task_id)
        task_data.bookmark_name = bm_name

        # Upgrade the preparing row before submit.
        prep_id = getattr(self, "_prep_id", None)
        if prep_id:
            TaskStatusBar.instance().upgrade_preparing(prep_id, task_data.task_id)
            self._prep_id = None

        TaskManager.instance().submit_task(task_data)
        self.close()
        return True

    def _clear_preparing_row(self):
        """Remove the status-bar 'preparing' row if one is pending."""
        prep_id = getattr(self, "_prep_id", None)
        if not prep_id:
            return
        try:
            tsb = TaskStatusBar.instance()
            tsb._remove_task_row(prep_id)
            tsb._on_all_done()
        except Exception:
            pass
        self._prep_id = None
```

- [ ] **Step 2: Branch on the MB sentinel in `_do_execute`**

Find `_do_execute` at `word_hotkey.py:3577`. After the Outlook branch (around line 3585) and BEFORE the `word = self._get_word_app()` block, add:

```python
            # Mediation Brief dispatch — must run before the generic Word path
            # so that we bypass the normal "get selection → build prompt" flow.
            combo_data = self.prompt_combo.currentData()
            if combo_data == MB_REFINE_SENTINEL:
                print("[DEBUG] _do_execute: MB Refine branch")
                if not self._handle_mb_refine(prompt):
                    return
                return
```

Do NOT add the MB_ADD_QUOTES branch yet — that's Task 9.

- [ ] **Step 3: Run the import sanity check**

Run: `python -c "import icharlotte_core.word_hotkey; print('OK')"`
Expected: `OK`.

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat(word-popup): dispatch refine-section via live parser and build_refinement_prompts"
```

---

## Task 9: Word popup — `_do_execute` dispatch for "Mediation Brief: Add Quotes"

**Files:**
- Modify: `icharlotte_core/word_hotkey.py`

**Context for the engineer:**
- For Add Quotes we do NOT submit an LLM task. We open `WordQuoteInsertionDialog` modally; on accept we receive a list of quote dicts and splice them into the Word document at the current cursor via `insert_formatted_quotes_at_range`.
- The insertion point is the Word selection/cursor at the moment the popup was shown. Capture it as `range_start`/`range_end` before opening the dialog (by reading `word.Selection.Range.Start` / `.End`), and at insertion time call `doc.Range(range_start, range_end)` to resolve the Word Range object.
- `insert_formatted_quotes_at_range` always runs on the main thread — quote insertion does no LLM work and `Range.InsertFile` is fast, so we skip TaskManager here. The status-bar preparing row is cleared before opening the dialog.
- On success, show a transient status (reuse existing toast/status mechanism if one exists in the popup; otherwise just log).

- [ ] **Step 1: Add helper `_handle_mb_add_quotes`**

Add this method to `WordLLMPopup` class alongside `_handle_mb_refine`:

```python
    def _handle_mb_add_quotes(self) -> bool:
        """Dispatch for "Mediation Brief: Add Quotes".

        Opens the WordQuoteInsertionDialog modally. On acceptance, inserts
        the selected quotes into the live Word document at the cursor.
        Returns True to indicate the caller should stop processing.
        """
        from icharlotte_core.mediation_brief_live import (
            insert_formatted_quotes_at_range,
        )
        from icharlotte_core.ui.quote_dialog_word import WordQuoteInsertionDialog

        doc = self._original_document
        if doc is None:
            QMessageBox.warning(
                self, "Word Not Found",
                "No active Word document — please open the brief first."
            )
            self._clear_preparing_row()
            return True

        # Capture the current cursor/selection range for later insertion.
        try:
            sel = doc.Application.Selection
            range_start = int(sel.Range.Start)
            range_end = int(sel.Range.End)
        except Exception as e:
            QMessageBox.warning(
                self, "Cursor not found",
                f"Could not read current Word cursor position: {e}"
            )
            self._clear_preparing_row()
            return True

        # Clear the preparing row; we're not submitting an LLM task.
        self._clear_preparing_row()

        dialog = WordQuoteInsertionDialog(parent=self)
        inserted_quotes: List[Dict] = []

        def _on_quotes_to_insert(quotes):
            inserted_quotes.extend(quotes)

        dialog.quotes_to_insert.connect(_on_quotes_to_insert)
        result = dialog.exec()

        if result != QDialog.DialogCode.Accepted or not inserted_quotes:
            self.close()
            return True

        # Splice the quotes into the Word doc at the captured range.
        try:
            target_range = doc.Range(range_start, range_end)
            insert_formatted_quotes_at_range(doc, target_range, inserted_quotes)
            print(f"[WordLLMPopup] Inserted {len(inserted_quotes)} depo quote(s)")
        except Exception as e:
            QMessageBox.warning(
                self, "Insertion failed",
                f"Failed to insert quotes into Word: {e}"
            )

        self.close()
        return True
```

Ensure these imports are present at the top of `word_hotkey.py` (they almost certainly are, but double-check):

```python
from typing import Any, Dict, List, Optional  # noqa: F401
from PySide6.QtWidgets import QDialog, QMessageBox  # noqa: F401
```

- [ ] **Step 2: Branch on the MB_ADD_QUOTES sentinel in `_do_execute`**

Find the block added in Task 8 (inside `_do_execute`):

```python
            combo_data = self.prompt_combo.currentData()
            if combo_data == MB_REFINE_SENTINEL:
                print("[DEBUG] _do_execute: MB Refine branch")
                if not self._handle_mb_refine(prompt):
                    return
                return
```

Extend it:

```python
            combo_data = self.prompt_combo.currentData()
            if combo_data == MB_REFINE_SENTINEL:
                print("[DEBUG] _do_execute: MB Refine branch")
                if not self._handle_mb_refine(prompt):
                    return
                return
            if combo_data == MB_ADD_QUOTES_SENTINEL:
                print("[DEBUG] _do_execute: MB Add Quotes branch")
                self._handle_mb_add_quotes()
                return
```

- [ ] **Step 3: Run the import sanity check**

Run: `python -c "import icharlotte_core.word_hotkey; print('OK')"`
Expected: `OK`.

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat(word-popup): dispatch Add Quotes via WordQuoteInsertionDialog"
```

---

## Task 10: Manual integration test checklist + memory update

**Files:**
- Modify: `C:\Users\ASerpik.DESKTOP-MRIMK0D\.claude\projects\C--geminiterminal2\memory\mediation_brief.md`

**Context for the engineer:**
- These integration tests require a real running iCharlotte + open Word. They are a manual checklist to run once end-to-end. Check each item off as it passes.

- [ ] **Step 1: Integration test — brief detection**

Start iCharlotte. Generate a mediation brief via the chat tab Templates dropdown so Word opens the assembled brief. Keep Word open. Press **Win+V**. The Word AI Assistant popup appears.

Verify: the prompt dropdown contains a separator followed by:
- "Mediation Brief: Refine Section"
- "Mediation Brief: Add Quotes"

Close iCharlotte, reopen iCharlotte, open the saved brief `.docx` in Word directly (skip regeneration). Press **Win+V**. Verify the two entries still appear — this confirms the cross-session path works.

- [ ] **Step 2: Integration test — refine section (direct replace)**

With the brief open, press **Win+V**. Select "Mediation Brief: Refine Section". The section combo appears below the custom prompt input. Pick "LIABILITY". Type `Make the causation discussion more forceful.` Leave Redline OFF. Press Execute.

Verify:
- Popup closes immediately.
- Status bar shows a task row for the refinement.
- When the task completes, the Liability section body text is replaced.
- Other sections (Intro, Damages, Conclusion) are byte-for-byte unchanged — compare with a pre-refine copy. Use Word's Compare feature or `git diff` on saved copies.

- [ ] **Step 3: Integration test — refine section (redline)**

Repeat Step 2 but with Redline mode CHECKED. Verify that the Liability section now contains tracked changes and unchanged sentences are not redlined (matches existing redline-engine behavior for other template flows).

- [ ] **Step 4: Integration test — add quotes at cursor**

Place your Word cursor inside the Liability section. Press **Win+V**, pick "Mediation Brief: Add Quotes", click Execute. The `WordQuoteInsertionDialog` opens.

Add a real deposition transcript PDF (from any case). Enter a search description. Click Search. When results come back, tick 2 results and click Insert Selected.

Verify:
- Dialog closes.
- Two formatted quote blocks appear at your cursor position.
- Each block has Q./A. at 0.5" indent and testimony at 1.0" (hanging indent), single-spaced.
- Each citation line is at the left margin with visible space above it.
- Surrounding paragraphs in Liability are untouched.

- [ ] **Step 5: Integration test — non-brief document**

Open a non-brief Word doc (any simple memo). Press **Win+V**. Verify that the "Mediation Brief" entries are NOT present in the prompt dropdown.

- [ ] **Step 6: Integration test — concurrent refines**

With a brief open, trigger a refine on Liability, then immediately (before it finishes) trigger a refine on Damages. Verify both tasks show up in the status bar and both land in their respective sections without clobbering each other.

- [ ] **Step 7: Update memory file**

Append to `C:\Users\ASerpik.DESKTOP-MRIMK0D\.claude\projects\C--geminiterminal2\memory\mediation_brief.md`:

```markdown

## Word AI Assistant integration (2026-04-14)
- "Mediation Brief: Refine Section" and "Mediation Brief: Add Quotes" dropdown entries in the Win+V popup, gated by `is_mediation_brief(active_doc)` (≥3 recognised roman-numeral headings).
- Refine: `mediation_brief_live.parse_brief_from_word_doc()` → `build_refinement_prompts()` → TaskManager → bookmark range replaces target section. Redline checkbox stays orthogonal.
- Add Quotes: `WordQuoteInsertionDialog` (stripped — no section/subsection/Weave) → `insert_formatted_quotes_at_range()` builds a temp docx via `_add_depo_quote` and Word COM `Range.InsertFile` splices at cursor.
- No sidecar, no session state — everything parsed from the live doc each time. `planning_output` is unavailable on cross-session refines; style excerpts still load from disk cache.
- Sentinel userData values `MB_REFINE_SENTINEL`, `MB_ADD_QUOTES_SENTINEL` identify the entries in `WordLLMPopup.on_prompt_selected` / `_do_execute`.
- Chat tab flow is untouched — `build_refinement_prompts` saves/restores `self.sections` so the throwaway generator never leaks state.
```

- [ ] **Step 8: Commit**

```bash
git add C:/Users/ASerpik.DESKTOP-MRIMK0D/.claude/projects/C--geminiterminal2/memory/mediation_brief.md
git commit -m "docs(memory): mediation brief Word AI assistant integration"
```

---

## Self-Review

**Spec coverage:**
- Refine section from Word popup → Tasks 1, 6, 7, 8.
- Add Quotes from Word popup → Tasks 4, 5, 9.
- Live parsing + detection → Tasks 2, 3.
- No sidecar / no session state → enforced throughout; `build_refinement_prompts` saves/restores `self.sections`; dialog uses throwaway generator.
- Orthogonal redline checkbox → Task 8 passes existing redline settings into the TaskData.
- Cross-session usability → Task 10 Step 1 verifies.
- Non-brief docs → Task 6 gate + Task 10 Step 5 verifies.
- Concurrency via TaskManager bookmarks → Task 8 uses `_create_task_bookmark` + Task 10 Step 6 verifies.
- `word_validator` gate → Reuses existing insertion path's validator hooks; no new call needed because `TaskManager` + `_InsertionProxy` already run the validator for every task. Quote insertion via `Range.InsertFile` doesn't currently hit the validator — noting this as a gap, but matching the existing chat-tab quote insertion's behavior (which also uses python-docx save without a live-doc validator pass).

**Placeholder scan:** No TODOs, no "similar to", no "handle edge cases" hand-waves. Every code step contains complete code.

**Type consistency:**
- `build_refinement_prompts(section_name, sections_dict, instruction) -> tuple` — called in Task 8 with exactly those kwargs.
- `LiveSection(name, heading_title, text, start_para_index, end_para_index)` — dataclass fields match construction and access in Tasks 2, 3, 8.
- `LiveBrief.sections: Dict[str, LiveSection]` — access pattern `live.sections.get(section_name)` used consistently in Task 8.
- `insert_formatted_quotes_at_range(doc_com, range_com, quotes)` — signature matches the call site in Task 9.
- Sentinel constants `MB_REFINE_SENTINEL` / `MB_ADD_QUOTES_SENTINEL` — defined Task 6, used Tasks 7 + 8 + 9.
- `_handle_mb_refine`, `_handle_mb_add_quotes`, `_clear_preparing_row` — all defined and called within the same class.

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-04-14-mediation-brief-word-assistant.md`. Two execution options:

**1. Subagent-Driven (recommended)** — I dispatch a fresh subagent per task, review between tasks, fast iteration.

**2. Inline Execution** — Execute tasks in this session using executing-plans, batch execution with checkpoints.

Which approach?
