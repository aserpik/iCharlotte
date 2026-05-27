"""Verify statute citations against California Legislative Information.

Wraps the existing CALegInfoClient with on-disk JSON caching and an LLM
comparison step. Returns a verdict-bearing CitationVerification.
"""

from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import asdict
from typing import Callable

from icharlotte_core.legal_research.models import StatuteResult
from icharlotte_core.legal_research.sources.ca_leginfo import CALegInfoClient
from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.prompt_manager import get_prompt

logger = logging.getLogger(__name__)

LLMCallback = Callable[[str, str], str]

_VALID_VERDICTS = {"SUPPORTED", "PARTIAL", "NOT_SUPPORTED"}


class StatuteVerifier:
    def __init__(
        self,
        *,
        leginfo_client: CALegInfoClient,
        llm_callback: LLMCallback,
        cache_dir: str,
    ) -> None:
        self.leginfo = leginfo_client
        self.llm = llm_callback
        self.cache_dir = cache_dir
        os.makedirs(self.cache_dir, exist_ok=True)

    def _cache_path(self, law_code: str, section_num: str) -> str:
        safe = re.sub(r"[^A-Za-z0-9._-]", "_", f"{law_code}_{section_num}")
        return os.path.join(self.cache_dir, f"{safe}.json")

    def _load_cached(self, law_code: str, section_num: str) -> StatuteResult | None:
        path = self._cache_path(law_code, section_num)
        if not os.path.exists(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return StatuteResult(
                code=data.get("code", law_code),
                section=data.get("section", section_num),
                title=data.get("title", ""),
                text=data.get("text", ""),
                url=data.get("url", ""),
            )
        except (OSError, ValueError, json.JSONDecodeError):
            logger.warning("Could not read statute cache: %s", path, exc_info=True)
            return None

    def _save_cached(self, statute: StatuteResult) -> None:
        path = self._cache_path(statute.code, statute.section)
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(asdict(statute), f, indent=2)
        except OSError:
            logger.warning("Could not write statute cache: %s", path, exc_info=True)

    def verify(self, citation: Citation) -> CitationVerification:
        cv = CitationVerification(
            citation_text=citation.raw_text,
            normalized_citation=citation.normalized,
            kind="statute",
            law_code=citation.law_code,
            section_num=citation.section_num,
            proposition=citation.proposition,
            body_offset=citation.body_offset,
        )

        # 1. Cache check
        statute = self._load_cached(citation.law_code, citation.section_num)

        # 2. Fetch from leginfo if not cached
        if statute is None:
            statute = self.leginfo.get_section(citation.law_code, citation.section_num)
            if statute is not None:
                self._save_cached(statute)

        # 3. NOT_FOUND short-circuit
        if statute is None or not statute.text.strip():
            cv.verdict = "NOT_FOUND"
            cv.note = (
                "This statute section was not found at leginfo; it may be "
                "invented, repealed, or mis-cited."
            )
            return cv

        cv.opinion_url = statute.url
        cv.case_name = statute.title  # repurposed for header display

        # 4. LLM comparison
        prompt_template = get_prompt("oppose_motion", "verify_citation") or ""
        if not prompt_template:
            cv.verdict = "UNVERIFIED"
            cv.note = "Verifier prompt not configured."
            return cv

        user_prompt = prompt_template.format(
            proposition=citation.proposition or "(no proposition extracted)",
            citation_text=citation.raw_text,
            authority_text=statute.text,
        )
        try:
            response = self.llm("", user_prompt) or ""
        except Exception:
            logger.warning("LLM verifier call failed", exc_info=True)
            cv.verdict = "UNVERIFIED"
            cv.note = "Verifier LLM call failed; verify manually."
            return cv

        verdict, evidence, note = _parse_verdict_response(response)
        if verdict not in _VALID_VERDICTS:
            cv.verdict = "UNVERIFIED"
            cv.note = "Could not parse verifier response; verify manually."
            return cv

        cv.verdict = verdict
        cv.evidence = evidence
        cv.note = note
        return cv


def _parse_verdict_response(text: str) -> tuple[str, str, str]:
    if not isinstance(text, str):
        return "", "", ""
    cleaned = text.strip()
    fenced = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", cleaned, re.DOTALL | re.I)
    if fenced:
        cleaned = fenced.group(1).strip()
    try:
        data = json.loads(cleaned)
    except (TypeError, ValueError):
        return "", "", ""
    if not isinstance(data, dict):
        return "", "", ""
    verdict = (data.get("verdict") or "").strip().upper()
    evidence = (data.get("evidence") or "").strip()
    note = (data.get("note") or "").strip()
    return verdict, evidence, note
