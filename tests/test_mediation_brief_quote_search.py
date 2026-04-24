"""Tests for MediationBriefGenerator.search_quotes and its helpers.

Covers the ID-based grounded quote search pipeline:
  - Keyword pre-filter ranking
  - Candidate formatting for the LLM
  - Match parsing rejects unknown/duplicate/invalid IDs
  - End-to-end search_quotes with mocked transcript parser and LLM
"""

import os
import sys
import unittest
from unittest.mock import patch, MagicMock

# Disable WMI platform query which hangs on this machine.
import platform
platform._wmi_query = lambda *a, **k: {}

from icharlotte_core.mediation_brief import MediationBriefGenerator
from icharlotte_core.deposition.models import (
    DeponentInfo, QAExchange, TranscriptIndex,
)


def _make_exchange(id_, q, a, page=10, line_start=1, line_end=5):
    return QAExchange(
        id=id_,
        question=q,
        answer=a,
        page_start=page,
        line_start=line_start,
        page_end=page,
        line_end=line_end,
    )


def _new_generator():
    # Bypass __init__ (reads the sample briefs directory) for unit tests
    return MediationBriefGenerator.__new__(MediationBriefGenerator)


class TestKeywordRanking(unittest.TestCase):
    def setUp(self):
        self.gen = _new_generator()

    def test_matching_exchange_ranks_above_nonmatching(self):
        matching = _make_exchange(
            1, "Did you inspect the warning signs?",
            "No, I never looked at any warning signs on the property."
        )
        unrelated = _make_exchange(
            2, "What time did you arrive?",
            "About three in the afternoon."
        )
        candidates = [
            (unrelated, "Smith", "smith.pdf"),
            (matching, "Smith", "smith.pdf"),
        ]
        ranked = self.gen._rank_exchanges_by_keywords(
            candidates, "testimony about warning signs on the premises", top_n=5
        )
        self.assertEqual(len(ranked), 1)
        self.assertIs(ranked[0][0], matching)

    def test_top_n_is_respected(self):
        candidates = []
        for i in range(20):
            ex = _make_exchange(
                i, f"Question number {i}",
                f"Answer number {i} about liability and warning signs."
            )
            candidates.append((ex, "Smith", "smith.pdf"))
        ranked = self.gen._rank_exchanges_by_keywords(
            candidates, "liability warning signs", top_n=5
        )
        self.assertEqual(len(ranked), 5)

    def test_empty_query_returns_full_list(self):
        # No content tokens → pass everything through; the LLM is the only
        # filter in the abstract-query degenerate case.
        candidates = [
            (_make_exchange(i, f"q{i}", f"a{i}"), "X", "x.pdf")
            for i in range(10)
        ]
        ranked = self.gen._rank_exchanges_by_keywords(candidates, "  ")
        self.assertEqual(len(ranked), 10)

    def test_no_overlap_returns_full_list(self):
        # Query has content tokens but none match any candidate — fall
        # through to the LLM rather than preserve an arbitrary prefix.
        candidates = [
            (_make_exchange(i, f"question {i}", f"answer {i}"), "X", "x.pdf")
            for i in range(5)
        ]
        ranked = self.gen._rank_exchanges_by_keywords(
            candidates, "supercalifragilistic expialidocious"
        )
        self.assertEqual(len(ranked), 5)

    def test_stopwords_alone_returns_full_list(self):
        # Query made entirely of stopwords has no real tokens.
        ex = _make_exchange(1, "Was the box on the table?", "Yes it was.")
        candidates = [(ex, "X", "x.pdf")]
        ranked = self.gen._rank_exchanges_by_keywords(candidates, "the of a")
        self.assertEqual(len(ranked), 1)

    def test_legal_meta_stopwords_are_ignored(self):
        # "testimony" and "liability" are in the meta-stopword list, so
        # a purely meta query should degrade gracefully to full-list.
        ex_match = _make_exchange(
            1, "Did you see the warning signs?", "No I did not."
        )
        ex_nomatch = _make_exchange(
            2, "What did you eat?", "A sandwich."
        )
        candidates = [
            (ex_match, "X", "x.pdf"),
            (ex_nomatch, "X", "x.pdf"),
        ]
        ranked = self.gen._rank_exchanges_by_keywords(
            candidates, "testimony supporting the liability argument"
        )
        # No content tokens → full list returned, order preserved
        self.assertEqual(len(ranked), 2)


class TestCandidateFormatting(unittest.TestCase):
    def test_format_includes_id_deponent_citation_and_qa(self):
        gen = _new_generator()
        ex = _make_exchange(
            42, "Did you trip on the step?", "Yes I did.",
            page=47, line_start=12, line_end=18,
        )
        text = gen._format_candidates_for_llm([(ex, "Smith", "smith.pdf")])
        self.assertIn("[EX_0000]", text)
        self.assertIn("Smith", text)
        self.assertIn("47:12-18", text)
        self.assertIn("Q. Did you trip on the step?", text)
        self.assertIn("A. Yes I did.", text)


class TestMatchParsing(unittest.TestCase):
    def setUp(self):
        self.gen = _new_generator()
        self.candidates = [
            (_make_exchange(1, "Q one", "A one"), "Smith", "smith.pdf"),
            (_make_exchange(2, "Q two", "A two"), "Smith", "smith.pdf"),
            (_make_exchange(3, "Q three", "A three"), "Smith", "smith.pdf"),
        ]

    def test_parses_valid_match(self):
        output = (
            "MATCH_START\n"
            "ID: EX_0001\n"
            "RELEVANCE: Supports the liability argument.\n"
            "MATCH_END\n"
        )
        results = self.gen._parse_quote_match_results(output, self.candidates)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["relevance"], "Supports the liability argument.")
        self.assertEqual(results[0]["page_line"], "10:1-5")
        self.assertEqual(results[0]["deponent"], "Smith")
        self.assertEqual(results[0]["source"], "smith.pdf")
        self.assertIn("Q. Q two", results[0]["qa_text"])
        self.assertIn("A. A two", results[0]["qa_text"])

    def test_unknown_id_is_dropped(self):
        output = (
            "MATCH_START\nID: EX_0099\nRELEVANCE: fake\nMATCH_END\n"
            "MATCH_START\nID: EX_0000\nRELEVANCE: real\nMATCH_END\n"
        )
        results = self.gen._parse_quote_match_results(output, self.candidates)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["relevance"], "real")
        self.assertIn("Q. Q one", results[0]["qa_text"])

    def test_duplicate_ids_are_collapsed(self):
        output = (
            "MATCH_START\nID: EX_0001\nRELEVANCE: first\nMATCH_END\n"
            "MATCH_START\nID: EX_0001\nRELEVANCE: second\nMATCH_END\n"
        )
        results = self.gen._parse_quote_match_results(output, self.candidates)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["relevance"], "first")

    def test_no_matches_sentinel(self):
        results = self.gen._parse_quote_match_results(
            "NO_MATCHES_FOUND", self.candidates
        )
        self.assertEqual(results, [])

    def test_empty_output(self):
        results = self.gen._parse_quote_match_results("", self.candidates)
        self.assertEqual(results, [])

    def test_malformed_id_is_dropped(self):
        output = (
            "MATCH_START\nID: EX_abc\nRELEVANCE: bogus\nMATCH_END\n"
        )
        # Regex requires \d+ so this block simply doesn't match at all
        results = self.gen._parse_quote_match_results(output, self.candidates)
        self.assertEqual(results, [])

    def test_grounding_uses_indexed_text_not_llm_text(self):
        # Even if the LLM output contained its own Q/A text, the parser
        # ignores it and pulls Q/A from the candidate list. This is the
        # core anti-fabrication guarantee.
        output = (
            "MATCH_START\n"
            "ID: EX_0002\n"
            "RELEVANCE: supports claim\n"
            "Q. TOTALLY MADE UP QUESTION\n"
            "A. TOTALLY MADE UP ANSWER\n"
            "MATCH_END\n"
        )
        results = self.gen._parse_quote_match_results(output, self.candidates)
        self.assertEqual(len(results), 1)
        self.assertNotIn("MADE UP", results[0]["qa_text"])
        self.assertIn("Q. Q three", results[0]["qa_text"])


class TestMatchParsingContext(unittest.TestCase):
    """Preceding-exchange context should be prepended when the full
    per-transcript index is supplied."""

    def setUp(self):
        self.gen = _new_generator()
        self.ex1 = _make_exchange(1, "Q one", "A one")
        self.ex2 = _make_exchange(2, "Q two", "A two")
        self.ex3 = _make_exchange(3, "Q three", "A three")
        # Candidate list only exposes ex2 to the LLM.
        self.candidates = [(self.ex2, "Smith", "smith.pdf")]
        # But the full id map contains every usable exchange, so we can
        # look up the preceding one (ex1).
        self.full_by_source = {
            ("Smith", "smith.pdf"): {
                self.ex1.id: self.ex1,
                self.ex2.id: self.ex2,
                self.ex3.id: self.ex3,
            }
        }
        self.output = (
            "MATCH_START\nID: EX_0000\nRELEVANCE: key admission\nMATCH_END\n"
        )

    def test_preceding_exchange_is_prepended_as_context(self):
        results = self.gen._parse_quote_match_results(
            self.output, self.candidates, self.full_by_source
        )
        self.assertEqual(len(results), 1)
        qa = results[0]["qa_text"]
        # Context Q/A comes first, then a blank line, then the matched Q/A.
        self.assertTrue(
            qa.startswith("Q. Q one\nA. A one\n\nQ. Q two\nA. A two"),
            f"unexpected qa_text: {qa!r}",
        )

    def test_no_context_when_preceding_exchange_missing(self):
        # First exchange in a transcript (id=1) has no preceding usable
        # exchange (id=0) — qa_text should contain only the match.
        candidates = [(self.ex1, "Smith", "smith.pdf")]
        results = self.gen._parse_quote_match_results(
            self.output, candidates, self.full_by_source
        )
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["qa_text"], "Q. Q one\nA. A one")

    def test_no_context_when_full_map_is_none(self):
        # Default call without the map (legacy callers) keeps behaviour
        # stable: no context is added.
        results = self.gen._parse_quote_match_results(
            self.output, self.candidates, None
        )
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["qa_text"], "Q. Q two\nA. A two")

    def test_preceding_skipped_if_not_in_usable_map(self):
        # If the preceding id is missing from the map (e.g. filtered out
        # by the quality filter), don't fabricate — just skip context.
        partial_map = {
            ("Smith", "smith.pdf"): {
                self.ex2.id: self.ex2,  # ex1 missing
            }
        }
        results = self.gen._parse_quote_match_results(
            self.output, self.candidates, partial_map
        )
        self.assertEqual(results[0]["qa_text"], "Q. Q two\nA. A two")


class TestSearchQuotesIntegration(unittest.TestCase):
    def _make_index(self):
        return TranscriptIndex(
            source_pdf="fake.pdf",
            deponent=DeponentInfo(full_name="Valerie Smith", last_name="Smith"),
            exchanges=[
                _make_exchange(
                    1, "Did you see the warning signs before entering?",
                    "No, I did not see any signs that day.",
                    page=47, line_start=12, line_end=15,
                ),
                _make_exchange(
                    2, "What did you have for lunch?",
                    "A turkey sandwich.",
                    page=10, line_start=3, line_end=4,
                ),
                _make_exchange(
                    3, "Were the warning signs visible from the entrance?",
                    "No, they were blocked by the display case.",
                    page=48, line_start=1, line_end=5,
                ),
            ],
            total_transcript_pages=100,
            is_condensed=True,
        )

    @patch("icharlotte_core.mediation_brief.LLMCaller")
    @patch("icharlotte_core.deposition.transcript_parser.TranscriptParser")
    def test_end_to_end_returns_grounded_results(
        self, MockParser, MockLLMCaller,
    ):
        # Parser returns our synthetic index for any path
        parser_instance = MagicMock()
        parser_instance.parse.return_value = self._make_index()
        MockParser.return_value = parser_instance

        # LLM picks the two relevant exchanges by ID
        llm_instance = MagicMock()
        llm_instance.call.return_value = (
            "MATCH_START\n"
            "ID: EX_0000\n"
            "RELEVANCE: Witness admits not seeing warning signs.\n"
            "MATCH_END\n"
            "MATCH_START\n"
            "ID: EX_0001\n"
            "RELEVANCE: Warning signs were blocked from view.\n"
            "MATCH_END\n"
        )
        MockLLMCaller.return_value = llm_instance

        gen = _new_generator()
        # Stub out the prompt loader so we don't depend on the file system
        with patch.object(gen, "_get_section_prompt", return_value="PROMPT"):
            results = gen.search_quotes(
                ["fake.pdf"], "testimony about warning signs not being visible"
            )

        self.assertEqual(len(results), 2)
        for r in results:
            self.assertEqual(r["deponent"], "Smith")
            self.assertEqual(r["source"], "fake.pdf")
            self.assertIn("warning signs", r["qa_text"].lower())
            self.assertTrue(r["page_line"])
            self.assertTrue(r["relevance"])

        # Pre-filter must have dropped the turkey-sandwich exchange before
        # ever reaching the LLM — confirm by checking the text that was
        # actually sent to the LLM.
        sent_text = llm_instance.call.call_args.kwargs["text"]
        self.assertIn("warning signs", sent_text.lower())
        self.assertNotIn("turkey", sent_text.lower())

    @patch("icharlotte_core.mediation_brief.LLMCaller")
    @patch("icharlotte_core.deposition.transcript_parser.TranscriptParser")
    def test_no_matches_returns_empty_list(
        self, MockParser, MockLLMCaller,
    ):
        parser_instance = MagicMock()
        parser_instance.parse.return_value = self._make_index()
        MockParser.return_value = parser_instance

        llm_instance = MagicMock()
        llm_instance.call.return_value = "NO_MATCHES_FOUND"
        MockLLMCaller.return_value = llm_instance

        gen = _new_generator()
        with patch.object(gen, "_get_section_prompt", return_value="PROMPT"):
            results = gen.search_quotes(
                ["fake.pdf"], "something completely unrelated xyz123"
            )
        self.assertEqual(results, [])

    @patch("icharlotte_core.mediation_brief.LLMCaller")
    @patch("icharlotte_core.deposition.transcript_parser.TranscriptParser")
    def test_parser_failure_skips_transcript(
        self, MockParser, MockLLMCaller,
    ):
        parser_instance = MagicMock()
        parser_instance.parse.side_effect = RuntimeError("parse boom")
        MockParser.return_value = parser_instance

        gen = _new_generator()
        with patch.object(gen, "_get_section_prompt", return_value="PROMPT"):
            results = gen.search_quotes(["bad.pdf"], "anything")

        self.assertEqual(results, [])
        MockLLMCaller.assert_not_called()

    @patch("icharlotte_core.mediation_brief.LLMCaller")
    @patch("icharlotte_core.deposition.transcript_parser.TranscriptParser")
    def test_non_pdf_transcript_is_skipped(
        self, MockParser, MockLLMCaller,
    ):
        gen = _new_generator()
        with patch.object(gen, "_get_section_prompt", return_value="PROMPT"):
            results = gen.search_quotes(["foo.docx"], "anything")
        self.assertEqual(results, [])
        MockParser.return_value.parse.assert_not_called()
        MockLLMCaller.assert_not_called()


class TestBriefContext(unittest.TestCase):
    """Brief context passthrough — verifies the LLM gets oriented on the
    actual defense arguments in the brief.
    """

    def setUp(self):
        self.gen = _new_generator()

    def test_format_brief_context_none_returns_empty(self):
        self.assertEqual(self.gen._format_brief_context(None), "")

    def test_format_brief_context_empty_dict_returns_empty(self):
        self.assertEqual(self.gen._format_brief_context({}), "")

    def test_format_brief_context_skips_empty_section_bodies(self):
        ctx = self.gen._format_brief_context({
            "liability": "The hazard was open and obvious.",
            "damages":   "   ",   # whitespace-only — skip
            "conclusion": "",     # empty — skip
        })
        self.assertIn("BRIEF CONTEXT", ctx)
        self.assertIn("IV. LIABILITY", ctx)
        self.assertIn("open and obvious", ctx)
        self.assertNotIn("V. DAMAGES", ctx)
        self.assertNotIn("VII. CONCLUSION", ctx)

    def test_format_brief_context_uses_roman_numerals(self):
        ctx = self.gen._format_brief_context({
            "damages": "Plaintiff's medical bills are inflated.",
        })
        self.assertIn("V. DAMAGES", ctx)
        self.assertIn("inflated", ctx)

    def test_format_brief_context_emits_sections_in_canonical_order(self):
        ctx = self.gen._format_brief_context({
            # deliberately reversed input order
            "damages": "DAMAGES BODY",
            "liability": "LIABILITY BODY",
            "statement_of_facts": "FACTS BODY",
        })
        facts_idx = ctx.index("FACTS BODY")
        liab_idx = ctx.index("LIABILITY BODY")
        dam_idx = ctx.index("DAMAGES BODY")
        self.assertLess(facts_idx, liab_idx)
        self.assertLess(liab_idx, dam_idx)

    @patch("icharlotte_core.mediation_brief.LLMCaller")
    @patch("icharlotte_core.deposition.transcript_parser.TranscriptParser")
    def test_search_quotes_passes_brief_context_to_llm(
        self, MockParser, MockLLMCaller,
    ):
        # Parser returns a minimal index with one usable exchange
        ex = _make_exchange(
            1, "Did the damages include lost wages?",
            "No, I never took any time off work.",
            page=50, line_start=1, line_end=5,
        )
        idx = TranscriptIndex(
            source_pdf="fake.pdf",
            deponent=DeponentInfo(full_name="Jane Doe", last_name="Doe"),
            exchanges=[ex],
            total_transcript_pages=100,
        )
        MockParser.return_value.parse.return_value = idx

        llm_instance = MagicMock()
        llm_instance.call.return_value = "NO_MATCHES_FOUND"
        MockLLMCaller.return_value = llm_instance

        gen = _new_generator()
        brief_sections = {
            "damages": "Plaintiff's damages claim is grossly inflated. Medical records show no lost wages.",
            "liability": "Plaintiff assumed the risk.",
        }
        with patch.object(gen, "_get_section_prompt", return_value="PROMPT"):
            gen.search_quotes(
                ["fake.pdf"],
                "testimony supporting our damages arguments",
                brief_sections=brief_sections,
            )

        sent_text = llm_instance.call.call_args.kwargs["text"]
        self.assertIn("BRIEF CONTEXT", sent_text)
        self.assertIn("IV. LIABILITY", sent_text)
        self.assertIn("V. DAMAGES", sent_text)
        self.assertIn("grossly inflated", sent_text)
        self.assertIn("assumed the risk", sent_text)
        self.assertIn("CANDIDATE EXCHANGES", sent_text)
        self.assertIn("[EX_0000]", sent_text)

    @patch("icharlotte_core.mediation_brief.LLMCaller")
    @patch("icharlotte_core.deposition.transcript_parser.TranscriptParser")
    def test_search_quotes_without_brief_context_omits_block(
        self, MockParser, MockLLMCaller,
    ):
        ex = _make_exchange(
            1, "Did you see the hazard?", "Yes I saw it clearly."
        )
        idx = TranscriptIndex(
            source_pdf="fake.pdf",
            deponent=DeponentInfo(last_name="X"),
            exchanges=[ex],
        )
        MockParser.return_value.parse.return_value = idx

        llm_instance = MagicMock()
        llm_instance.call.return_value = "NO_MATCHES_FOUND"
        MockLLMCaller.return_value = llm_instance

        gen = _new_generator()
        with patch.object(gen, "_get_section_prompt", return_value="PROMPT"):
            gen.search_quotes(["fake.pdf"], "anything")  # no brief_sections

        sent_text = llm_instance.call.call_args.kwargs["text"]
        self.assertNotIn("BRIEF CONTEXT", sent_text)
        self.assertNotIn("CANDIDATE EXCHANGES", sent_text)  # header only added when brief present
        self.assertIn("[EX_0000]", sent_text)


class TestQuoteSearchWorkerForwarding(unittest.TestCase):
    """Verify the Qt worker threads forward brief_sections through to the
    generator. Uses a MagicMock generator to avoid Qt thread startup.
    """

    def test_worker_forwards_brief_sections(self):
        from icharlotte_core.ui.quote_dialog import QuoteSearchWorker

        generator = MagicMock()
        generator.search_quotes.return_value = []

        worker = QuoteSearchWorker(
            generator, ["x.pdf"], "find damages",
            brief_sections={"damages": "body"},
        )
        worker.run()

        generator.search_quotes.assert_called_once_with(
            ["x.pdf"], "find damages",
            brief_sections={"damages": "body"},
        )

    def test_worker_defaults_brief_sections_to_none(self):
        from icharlotte_core.ui.quote_dialog import QuoteSearchWorker

        generator = MagicMock()
        generator.search_quotes.return_value = []

        worker = QuoteSearchWorker(generator, ["x.pdf"], "anything")
        worker.run()

        generator.search_quotes.assert_called_once_with(
            ["x.pdf"], "anything", brief_sections=None,
        )


if __name__ == "__main__":
    unittest.main()
