"""Legal Research Engine -- core pipeline for searching and verifying legal authority."""
import hashlib
import json
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Dict, List, Optional, Tuple

from .models import CaseResult, ResearchResult, StatuteResult, VerificationStatus
from .prompts import (
    QUERY_PLANNING_PROMPT,
    RELEVANCE_RANKING_PROMPT,
    SYNTHESIS_PROMPT,
    VERIFICATION_PROMPT,
)
from .sources.courtlistener import CourtListenerClient
from .sources.ca_leginfo import CALegInfoClient
from .sources.ca_courts import CACourtsClient

logger = logging.getLogger(__name__)

# Type alias: llm_callback(system_prompt, user_prompt) -> str
LLMCallback = Callable[[str, str], str]

# Max cases to keep after relevance filtering
_MAX_RELEVANT_CASES = 15
# Max chars of opinion text to fetch per case
_OPINION_SNIPPET_LENGTH = 1500


class LegalResearchEngine:
    """Stateless research engine -- no Qt dependency.

    Pipeline:
    1. Query planning (LLM extracts search terms from natural language)
    2. Parallel source search (CourtListener, CA Leginfo, CA Courts)
    3. Relevance filtering (LLM selects top cases)
    4. Case enrichment (fetch opinion text for top cases)
    5. Synthesis (LLM produces memo with citations)
    6. Verification (LLM cross-checks every citation against source data)
    """

    def __init__(self, courtlistener_token: str):
        self.cl_client = CourtListenerClient(token=courtlistener_token)
        self.leginfo_client = CALegInfoClient()
        self.courts_client = CACourtsClient()

    def research(
        self,
        query: str,
        llm_callback: LLMCallback,
        status_callback: Optional[Callable[[str], None]] = None,
        file_number: Optional[str] = None,
        data_manager=None,
    ) -> ResearchResult:
        """Run the full research pipeline.

        Args:
            query: Natural language legal question or document context.
            llm_callback: Function(system_prompt, user_prompt) -> str for LLM calls.
            status_callback: Optional function(status_text) for progress updates.
            file_number: Optional case file number for caching.
            data_manager: Optional CaseDataManager instance for caching.

        Returns:
            A ResearchResult with cases, statutes, memo, and verification.
        """
        def status(msg: str):
            if status_callback:
                status_callback(msg)

        # Check cache
        if file_number and data_manager:
            cache_key = self._cache_key(query)
            try:
                cached = data_manager.get_value(file_number, cache_key)
                if cached and isinstance(cached, dict):
                    status("Using cached research results")
                    return self._from_cache(cached)
            except Exception:
                logger.debug("Cache miss for %s", cache_key)

        # Phase 1: Query Planning
        status("Analyzing legal question...")
        search_plan = self._plan_queries(query, llm_callback)

        # Phase 2: Parallel Source Search
        status("Searching legal databases...")
        cases, statutes = self._search_sources(search_plan)

        # Phase 3: Relevance Filtering
        if len(cases) > _MAX_RELEVANT_CASES:
            status("Filtering most relevant cases...")
            cases = self._rank_and_filter(query, cases, llm_callback)

        # Phase 4: Case Enrichment (fetch opinion text for top cases)
        if cases:
            status("Fetching case holdings...")
            cases = self._enrich_top_cases(cases)

        # Phase 5: Synthesis
        status("Synthesizing legal analysis...")
        result = ResearchResult(query=query, cases=cases, statutes=statutes)
        authority_block = result.format_authority_block()
        memo = llm_callback(
            SYNTHESIS_PROMPT,
            f"User's legal question: {query}\n\n{authority_block}",
        )
        result.memo = memo

        # Phase 6: Verification
        if cases or statutes:
            status("Verifying citations...")
            result.verification = self._verify_citations(
                memo, authority_block, llm_callback
            )
            # Apply corrections from verification
            corrected = self._apply_corrections(result.verification, memo)
            if corrected:
                result.memo = corrected

        # Save to cache
        if file_number and data_manager:
            cache_key = self._cache_key(query)
            try:
                data_manager.save_variable(
                    file_number, cache_key, result.to_dict(), tag="legal_research"
                )
            except Exception:
                logger.debug("Failed to save research to cache")

        status("Research complete")
        return result

    def _plan_queries(self, query: str, llm_callback: LLMCallback) -> Dict:
        """Use LLM to extract structured search terms from natural language."""
        raw = llm_callback(QUERY_PLANNING_PROMPT, query)
        try:
            # Strip markdown code fences if present
            cleaned = raw.strip()
            if cleaned.startswith("```"):
                cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
                if cleaned.endswith("```"):
                    cleaned = cleaned[:-3]
                cleaned = cleaned.strip()
            return json.loads(cleaned)
        except (json.JSONDecodeError, ValueError):
            # Fallback: use the raw query as a single search term
            return {
                "case_queries": [query[:200]],
                "statute_queries": [],
                "legal_doctrines": [],
            }

    def _search_sources(
        self, plan: Dict
    ) -> Tuple[List[CaseResult], List[StatuteResult]]:
        """Search all sources in parallel based on the query plan."""
        cases: List[CaseResult] = []
        statutes: List[StatuteResult] = []

        with ThreadPoolExecutor(max_workers=5) as pool:
            futures = {}

            # CourtListener searches
            for q in plan.get("case_queries", []):
                f = pool.submit(self.cl_client.search_opinions, q)
                futures[f] = ("case", q)

            # CA Courts recent opinions (limit to first 2 queries)
            for q in plan.get("case_queries", [])[:2]:
                f = pool.submit(self.courts_client.search_recent, q)
                futures[f] = ("recent_case", q)

            # Statute lookups
            for ref in plan.get("statute_queries", []):
                code, section = self.leginfo_client.parse_code_reference(ref)
                if code and section:
                    f = pool.submit(self.leginfo_client.get_section, code, section)
                    futures[f] = ("statute", ref)

            # Collect results
            for future in as_completed(futures):
                source_type, _ = futures[future]
                try:
                    result = future.result()
                    if source_type in ("case", "recent_case"):
                        if isinstance(result, list):
                            cases.extend(result)
                    elif source_type == "statute":
                        if result is not None:
                            statutes.append(result)
                except Exception as e:
                    logger.warning("Source search error: %s", e)

        # Deduplicate cases by (name, citation)
        seen = set()
        unique_cases = []
        for c in cases:
            key = (c.name.lower(), c.citation.lower())
            if key not in seen:
                seen.add(key)
                unique_cases.append(c)

        return unique_cases, statutes

    def _rank_and_filter(
        self,
        query: str,
        cases: List[CaseResult],
        llm_callback: LLMCallback,
    ) -> List[CaseResult]:
        """Use LLM to select the most relevant cases."""
        # Build a compact list for the LLM
        case_list = []
        for i, c in enumerate(cases):
            entry = f"[{i}] {c.formatted_citation}"
            if c.snippet:
                entry += f" -- {c.snippet[:150]}"
            case_list.append(entry)

        user_prompt = (
            f"LEGAL QUESTION:\n{query[:3000]}\n\n"
            f"CASES ({len(cases)} total):\n" + "\n".join(case_list)
        )

        raw = llm_callback(RELEVANCE_RANKING_PROMPT, user_prompt)
        try:
            cleaned = raw.strip()
            if cleaned.startswith("```"):
                cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
                if cleaned.endswith("```"):
                    cleaned = cleaned[:-3]
                cleaned = cleaned.strip()
            data = json.loads(cleaned)
            selected = data.get("selected_cases", [])

            filtered = []
            for entry in selected:
                idx = entry.get("index")
                if isinstance(idx, int) and 0 <= idx < len(cases):
                    case = cases[idx]
                    # Store relevance note in snippet if snippet is empty
                    relevance = entry.get("relevance", "")
                    if relevance and not case.snippet:
                        case = CaseResult(
                            name=case.name,
                            citation=case.citation,
                            date=case.date,
                            court=case.court,
                            snippet=relevance,
                            url=case.url,
                            cluster_id=case.cluster_id,
                            negative_treatment=case.negative_treatment,
                            relevance_score=case.relevance_score,
                        )
                    filtered.append(case)

            return filtered if filtered else cases[:_MAX_RELEVANT_CASES]

        except (json.JSONDecodeError, ValueError):
            # Fallback: just take the first N
            return cases[:_MAX_RELEVANT_CASES]

    def _enrich_top_cases(self, cases: List[CaseResult]) -> List[CaseResult]:
        """Fetch opinion text for top cases that have cluster IDs."""
        enrichable = [c for c in cases if c.cluster_id]
        if not enrichable:
            return cases

        # Fetch opinion text in parallel (limit to top 10)
        enrichable = enrichable[:10]
        cluster_to_text: Dict[int, str] = {}

        with ThreadPoolExecutor(max_workers=5) as pool:
            futures = {}
            for c in enrichable:
                f = pool.submit(self.cl_client.get_opinion_text, c.cluster_id)
                futures[f] = c.cluster_id

            for future in as_completed(futures):
                cid = futures[future]
                try:
                    text = future.result()
                    if text:
                        cluster_to_text[cid] = text[:_OPINION_SNIPPET_LENGTH]
                except Exception:
                    logger.debug("Failed to fetch opinion for cluster %s", cid)

        # Replace snippets with actual opinion text
        enriched = []
        for c in cases:
            if c.cluster_id and c.cluster_id in cluster_to_text:
                enriched.append(CaseResult(
                    name=c.name,
                    citation=c.citation,
                    date=c.date,
                    court=c.court,
                    snippet=cluster_to_text[c.cluster_id],
                    url=c.url,
                    cluster_id=c.cluster_id,
                    negative_treatment=c.negative_treatment,
                    relevance_score=c.relevance_score,
                ))
            else:
                enriched.append(c)

        return enriched

    def _verify_citations(
        self,
        draft: str,
        authority_block: str,
        llm_callback: LLMCallback,
    ) -> List[VerificationStatus]:
        """Verify every citation in the draft against source data."""
        prompt = (
            f"DRAFT TEXT:\n{draft}\n\n"
            f"SOURCE DATA:\n{authority_block}\n\n"
            f"Verify every citation in the draft."
        )
        raw = llm_callback(VERIFICATION_PROMPT, prompt)
        try:
            cleaned = raw.strip()
            if cleaned.startswith("```"):
                cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
                if cleaned.endswith("```"):
                    cleaned = cleaned[:-3]
                cleaned = cleaned.strip()
            data = json.loads(cleaned)
            verifications = []
            for v in data.get("verifications", []):
                verifications.append(VerificationStatus(
                    citation=v.get("citation", ""),
                    status=v.get("status", "FLAGGED"),
                    detail=v.get("detail", ""),
                    original=v.get("original"),
                    corrected=v.get("corrected"),
                ))
            return verifications
        except (json.JSONDecodeError, ValueError):
            return []

    def _apply_corrections(
        self, verifications: List[VerificationStatus], draft: str
    ) -> Optional[str]:
        """Apply FIXED corrections and mark FLAGGED citations in the draft."""
        corrected = draft
        had_changes = False
        for v in verifications:
            if v.status == "FIXED" and v.original and v.corrected:
                if v.original in corrected:
                    corrected = corrected.replace(v.original, v.corrected, 1)
                    had_changes = True
            elif v.status == "FLAGGED" and v.citation:
                # Mark flagged citations
                if v.citation in corrected:
                    corrected = corrected.replace(
                        v.citation, f"{v.citation} [UNVERIFIED]", 1
                    )
                    had_changes = True
        return corrected if had_changes else None

    @staticmethod
    def _cache_key(query: str) -> str:
        """Generate a cache key for a query."""
        normalized = query.strip().lower()
        return f"legal_research_{hashlib.md5(normalized.encode()).hexdigest()[:12]}"

    @staticmethod
    def _from_cache(data: Dict) -> ResearchResult:
        """Reconstruct a ResearchResult from a cached dictionary."""
        cases = [
            CaseResult(
                name=c.get("name", ""),
                citation=c.get("citation", ""),
                date=c.get("date", ""),
                court=c.get("court", ""),
                snippet=c.get("snippet", ""),
                url=c.get("url", ""),
                cluster_id=c.get("cluster_id"),
                negative_treatment=c.get("negative_treatment"),
                relevance_score=c.get("relevance_score", 0.0),
            )
            for c in data.get("cases", [])
        ]
        statutes = [
            StatuteResult(
                code=s.get("code", ""),
                section=s.get("section", ""),
                title=s.get("title", ""),
                text=s.get("text", ""),
                url=s.get("url", ""),
            )
            for s in data.get("statutes", [])
        ]
        verifications = [
            VerificationStatus(
                citation=v.get("citation", ""),
                status=v.get("status", "FLAGGED"),
                detail=v.get("detail", ""),
                original=v.get("original"),
                corrected=v.get("corrected"),
            )
            for v in data.get("verification", [])
        ]
        return ResearchResult(
            query=data.get("query", ""),
            cases=cases,
            statutes=statutes,
            memo=data.get("memo", ""),
            verification=verifications,
        )
