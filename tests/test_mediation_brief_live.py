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

    def test_empty_section_body_yields_end_less_than_start(self):
        """A heading immediately followed by the next heading is a valid
        empty-body section. The contract is: end_para_index < start_para_index.
        Task 3's get_word_range_for_section uses this as the empty-body signal.
        """
        doc = make_doc(
            "I.   INTRODUCTION",
            "II.  STATEMENT OF FACTS",
            "Facts body.",
            "IV.  LIABILITY",
            "Liability body.",
        )
        live = parse_brief_from_word_doc(doc)
        intro = live.sections["introduction"]
        self.assertEqual(intro.text, "")
        # Introduction heading is at para 1; no body follows.
        # start_para_index should be 2 (the would-be first body paragraph)
        # and end_para_index should stay at 1 (the heading itself).
        self.assertEqual(intro.start_para_index, 2)
        self.assertEqual(intro.end_para_index, 1)
        # And the empty-body signal:
        self.assertLess(intro.end_para_index, intro.start_para_index)


if __name__ == "__main__":
    unittest.main()
