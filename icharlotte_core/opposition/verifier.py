"""Orchestrates per-citation verification across case + statute paths.

Routes each Citation to its appropriate verifier, deduplicates work by
``normalized`` form (re-using the verdict for repeated cites), runs in a
bounded thread pool, and emits per-citation progress messages.
"""

from __future__ import annotations

import concurrent.futures
import logging
import re as _re
from typing import Callable

from icharlotte_core.opposition.case_verifier import CaseVerifier
from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.models import CitationVerification, RetrievedAuthority
from icharlotte_core.opposition.statute_verifier import StatuteVerifier

logger = logging.getLogger(__name__)

ProgressCallback = Callable[[str], None]


class OppositionVerifier:
    def __init__(
        self,
        *,
        case_verifier: CaseVerifier,
        statute_verifier: StatuteVerifier,
        max_workers: int = 4,
    ) -> None:
        self.case = case_verifier
        self.statute = statute_verifier
        self.max_workers = max(1, int(max_workers))

    def verify_all(
        self,
        citations: list[Citation],
        *,
        on_progress: ProgressCallback | None = None,
    ) -> list[CitationVerification]:
        if not citations:
            return []

        # Dedup by normalized form. Keep first-occurrence Citation as representative.
        unique: dict[str, Citation] = {}
        for c in citations:
            key = c.normalized or c.raw_text
            if key not in unique:
                unique[key] = c

        # Verify uniques in a bounded thread pool.
        verdicts: dict[str, CitationVerification] = {}

        def _do_verify(c: Citation) -> tuple[str, CitationVerification]:
            key = c.normalized or c.raw_text
            try:
                if c.kind == "case":
                    cv = self.case.verify(c)
                elif c.kind == "statute":
                    cv = self.statute.verify(c)
                else:
                    cv = _unverified_for(c)
            except Exception:
                logger.warning("Verifier raised for %s", c.raw_text, exc_info=True)
                cv = _unverified_for(c, note="Verifier raised an exception; verify manually.")
            return key, cv

        if self.max_workers == 1:
            for c in unique.values():
                key, cv = _do_verify(c)
                verdicts[key] = cv
                if on_progress:
                    on_progress(_progress_line(cv))
        else:
            with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as pool:
                futures = {pool.submit(_do_verify, c): c for c in unique.values()}
                for fut in concurrent.futures.as_completed(futures):
                    key, cv = fut.result()
                    verdicts[key] = cv
                    if on_progress:
                        on_progress(_progress_line(cv))

        # Project unique verdicts back across all input citations (preserving order).
        results: list[CitationVerification] = []
        for c in citations:
            key = c.normalized or c.raw_text
            cv = verdicts.get(key)
            if cv is None:
                cv = _unverified_for(c)
            else:
                # Clone so per-cite body_offset is preserved (uniques used first occurrence).
                cv = _clone_with_offset(cv, c.body_offset)
            results.append(cv)
        return results


def _unverified_for(c: Citation, *, note: str = "") -> CitationVerification:
    if not note:
        note = (
            "Verifier does not cover this source (federal, treatise, local rule, "
            "or California Rule of Court in v1); verify manually."
        )
    return CitationVerification(
        citation_text=c.raw_text,
        normalized_citation=c.normalized,
        kind=c.kind or "unknown",
        proposition=c.proposition,
        body_offset=c.body_offset,
        case_name=c.case_name,
        law_code=c.law_code,
        section_num=c.section_num,
        verdict="UNVERIFIED",
        note=note,
    )


def _clone_with_offset(cv: CitationVerification, body_offset: int | None) -> CitationVerification:
    return CitationVerification.from_dict({**cv.to_dict(), "body_offset": body_offset})


def _progress_line(cv: CitationVerification) -> str:
    verdict_glyph = {
        "SUPPORTED": "OK",
        "PARTIAL": "PARTIAL",
        "NOT_SUPPORTED": "FAILED",
        "NOT_FOUND": "NOT FOUND",
        "UNVERIFIED": "skipped",
    }.get(cv.verdict, cv.verdict or "?")
    label = cv.citation_text or cv.normalized_citation or "(citation)"
    return f"  {verdict_glyph}: {label}"


def _norm_reporter(s: str) -> str:
    """Normalize a reporter citation for loose comparison: drop spaces, lowercase."""
    return _re.sub(r"\s+", "", (s or "")).lower()


def _reporter_in(needle: str, haystack: str) -> bool:
    """True if ``needle`` appears in ``haystack`` bounded by non-digits, so a
    volume/page number can't partially overlap a longer number (e.g.
    "0cal.app.4th100" must NOT match inside "100cal.app.4th1000")."""
    if not needle or not haystack:
        return False
    start = 0
    while True:
        idx = haystack.find(needle, start)
        if idx == -1:
            return False
        before_ok = idx == 0 or not haystack[idx - 1].isdigit()
        after = idx + len(needle)
        after_ok = after >= len(haystack) or not haystack[after].isdigit()
        if before_ok and after_ok:
            return True
        start = idx + 1


def _reporter_matches(cite: str, pool_citation: str) -> bool:
    """Whether a parsed cite and a pool citation refer to the same reporter cite.

    Exact after whitespace/case normalization, or one embedded in the other at a
    digit boundary — which tolerates a case-name prefix (``normalized_citation``)
    or a trailing pincite/parallel on the pool side, without the digit-overlap
    false positives a bare substring test produced.
    """
    c, p = _norm_reporter(cite), _norm_reporter(pool_citation)
    if not c or not p:
        return False
    if c == p:
        return True
    return _reporter_in(p, c) or _reporter_in(c, p)


def pool_membership_check(
    citations: list[Citation],
    retrieved: list[RetrievedAuthority],
) -> tuple[list[Citation], list[CitationVerification]]:
    """Split citations into (to_verify, off_pool_results).

    Case cites whose reporter citation is not present in the retrieved pool get
    a deterministic NOT_FOUND verdict (likely model-introduced). Statutes and
    rules always pass through. If the pool is empty (grounding produced
    nothing), everything passes through so the network verifier still runs.
    """
    if not retrieved:
        return list(citations), []

    pool_citations = [a.citation for a in retrieved if a.citation]
    to_verify: list[Citation] = []
    off_pool: list[CitationVerification] = []
    for c in citations:
        if c.kind != "case":
            to_verify.append(c)
            continue
        cite_str = c.reporter_citation or c.normalized
        in_pool = any(_reporter_matches(cite_str, p) for p in pool_citations)
        if in_pool:
            to_verify.append(c)
        else:
            off_pool.append(
                CitationVerification(
                    citation_text=c.raw_text,
                    normalized_citation=c.normalized,
                    kind="case",
                    case_name=c.case_name,
                    proposition=c.proposition,
                    body_offset=c.body_offset,
                    verdict="NOT_FOUND",
                    note=(
                        "Cited a case that was not in the researched authority "
                        "pool — likely model-introduced; verify or replace."
                    ),
                )
            )
    return to_verify, off_pool


def enrich_with_pool_signals(
    verifications: list[CitationVerification],
    retrieved: list[RetrievedAuthority],
) -> None:
    """Copy citation_count / latest_citing_year from the pool onto matching
    case verifications (matched by normalized reporter citation). Mutates in place."""
    if not retrieved:
        return
    pool = [a for a in retrieved if a.citation]
    for cv in verifications:
        if cv.kind != "case":
            continue
        for a in pool:
            if _reporter_matches(cv.normalized_citation, a.citation):
                cv.citation_count = a.citation_count
                cv.latest_citing_year = a.latest_citing_year
                break


import os as _os
from icharlotte_core.legal_research.sources.ca_leginfo import CALegInfoClient as _CALeg
from icharlotte_core.legal_research.sources.courtlistener import (
    CourtListenerClient as _CL,
)


def build_opposition_verifier(
    *,
    courtlistener_token: str,
    llm_callback: Callable[[str, str], str],
    max_workers: int = 4,
    cache_root: str | None = None,
) -> "OppositionVerifier":
    """Construct an OppositionVerifier wired to project cache dirs."""
    if cache_root is None:
        # Cache colocated with prompts.
        repo_root = _os.path.dirname(_os.path.dirname(_os.path.dirname(__file__)))
        cache_root = _os.path.join(
            repo_root, "Scripts", "prompts", "oppose_motion", ".cache"
        )
    opinion_cache = _os.path.join(cache_root, "opinions")
    statute_cache = _os.path.join(cache_root, "statutes")
    case_v = CaseVerifier(
        courtlistener_client=_CL(courtlistener_token),
        llm_callback=llm_callback,
        cache_dir=opinion_cache,
    )
    statute_v = StatuteVerifier(
        leginfo_client=_CALeg(),
        llm_callback=llm_callback,
        cache_dir=statute_cache,
    )
    return OppositionVerifier(
        case_verifier=case_v,
        statute_verifier=statute_v,
        max_workers=max_workers,
    )


def build_local_opposition_verifier(
    *,
    corpus,
    llm_callback: Callable[[str, str], str],
    max_workers: int = 4,
    cache_root: str | None = None,
) -> "OppositionVerifier":
    """OppositionVerifier whose case path is the local corpus (no network).

    Statute path keeps the existing leginfo verifier (not rate-limited).
    """
    from icharlotte_core.opposition.local_case_verifier import LocalCaseVerifier
    if cache_root is None:
        repo_root = _os.path.dirname(_os.path.dirname(_os.path.dirname(__file__)))
        cache_root = _os.path.join(repo_root, "Scripts", "prompts", "oppose_motion", ".cache")
    statute_v = StatuteVerifier(
        leginfo_client=_CALeg(), llm_callback=llm_callback,
        cache_dir=_os.path.join(cache_root, "statutes"),
    )
    return OppositionVerifier(
        case_verifier=LocalCaseVerifier(corpus=corpus, llm_callback=llm_callback),
        statute_verifier=statute_v,
        max_workers=max_workers,
    )


def find_replacement_candidates(
    *,
    failed_citation: CitationVerification,
    verifier: "OppositionVerifier",
    llm_callback: Callable[[str, str], str],
) -> list[CitationVerification]:
    """Propose and verify replacement candidates for a failed citation."""
    from icharlotte_core.opposition.citation_parser import extract_citations
    from icharlotte_core.opposition import prompts as default_prompts
    from icharlotte_core.prompt_manager import get_prompt
    import json as _json
    import re as _re

    template = get_prompt("oppose_motion", "find_replacement") or default_prompts.FIND_REPLACEMENT_PROMPT
    user_prompt = template.format(
        proposition=failed_citation.proposition or "",
        failed_citation=failed_citation.citation_text or "",
        verifier_note=failed_citation.note or "",
    )
    try:
        response = llm_callback("", user_prompt) or ""
    except Exception:
        logger.warning("find_replacement LLM call failed", exc_info=True)
        return []

    cleaned = response.strip()
    fenced = _re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", cleaned, _re.DOTALL | _re.I)
    if fenced:
        cleaned = fenced.group(1).strip()
    try:
        data = _json.loads(cleaned)
    except (TypeError, ValueError):
        return []
    raw_candidates = (data.get("candidates") if isinstance(data, dict) else []) or []

    # Parse each candidate's citation_text into a Citation and verify.
    citations = []
    for c in raw_candidates:
        if not isinstance(c, dict):
            continue
        text = c.get("citation_text", "") or ""
        parsed = extract_citations(text)
        if parsed:
            citations.append(parsed[0])

    if not citations:
        return []
    return verifier.verify_all(citations)
