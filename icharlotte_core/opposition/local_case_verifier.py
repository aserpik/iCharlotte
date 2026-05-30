"""Verify case citations against the LocalCaseCorpus (no network).

Mirrors CaseVerifier.verify: corpus lookup by reporter citation -> NOT_FOUND if
absent, else run the shared verify_citation prompt against local full text.
Carries the corpus's good-law soft signal onto the verdict.
"""
from __future__ import annotations

import logging
from typing import Callable

from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.opposition.statute_verifier import _parse_verdict_response
from icharlotte_core.prompt_manager import get_prompt

logger = logging.getLogger(__name__)

LLMCallback = Callable[[str, str], str]
_VALID = {"SUPPORTED", "PARTIAL", "NOT_SUPPORTED"}


class LocalCaseVerifier:
    def __init__(self, *, corpus, llm_callback: LLMCallback) -> None:
        self.corpus = corpus
        self.llm = llm_callback

    def verify(self, citation: Citation) -> CitationVerification:
        cv = CitationVerification(
            citation_text=citation.raw_text, normalized_citation=citation.normalized,
            kind="case", case_name=citation.case_name, date=citation.year,
            proposition=citation.proposition, body_offset=citation.body_offset,
        )
        cite = citation.reporter_citation or citation.normalized
        rec = None
        try:
            rec = self.corpus.lookup_by_citation(cite)
        except Exception:
            logger.warning("local corpus lookup failed for %s", cite, exc_info=True)
        if not rec:
            cv.verdict = "NOT_FOUND"
            cv.note = ("This citation is not in the local CA corpus (CAP through "
                       "~2020 + CourtListener recent); it may be invented, mis-cited, "
                       "out-of-state, or newer than the latest corpus build.")
            return cv

        cv.cluster_id = rec.get("case_uid", "")
        cv.opinion_url = rec.get("url", "")
        cv.court = rec.get("court", "")
        cv.date = rec.get("decision_date", "") or cv.date
        if rec.get("name") and not cv.case_name:
            cv.case_name = rec["name"]
        cv.citation_count = rec.get("citation_count")
        cv.latest_citing_year = rec.get("latest_citing_year", "") or ""

        text = rec.get("full_text") or ""
        template = get_prompt("oppose_motion", "verify_citation") or ""
        if not text or not template:
            cv.verdict = "UNVERIFIED"
            cv.note = "Corpus hit but no text/prompt available; verify manually."
            return cv

        user_prompt = template.format(
            proposition=citation.proposition or "(no proposition extracted)",
            citation_text=citation.raw_text, authority_text=text,
        )
        try:
            response = self.llm("", user_prompt) or ""
        except Exception:
            logger.warning("local verifier LLM call failed", exc_info=True)
            cv.verdict = "UNVERIFIED"; cv.note = "Verifier LLM call failed; verify manually."
            return cv

        verdict, evidence, note = _parse_verdict_response(response)
        if verdict not in _VALID:
            cv.verdict = "UNVERIFIED"; cv.note = "Could not parse verifier response; verify manually."
            return cv
        cv.verdict = verdict; cv.evidence = evidence; cv.note = note
        return cv
