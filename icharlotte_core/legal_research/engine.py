"""Legal Research Engine -- core pipeline for searching and verifying legal authority."""
import hashlib
import json
import logging
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Dict, List, Optional, Tuple

from .models import CaseResult, ResearchResult, StatuteResult, VerificationStatus
from .prompts import (
    get_query_planning_prompt,
    get_relevance_ranking_prompt,
    get_synthesis_prompt,
    get_verification_prompt,
)
from .sources.courtlistener import CourtListenerClient
from .sources.ca_leginfo import CALegInfoClient
from .sources.ca_courts import CACourtsClient

logger = logging.getLogger(__name__)

# Type alias: llm_callback(system_prompt, user_prompt) -> str
LLMCallback = Callable[[str, str], str]

# Max cases to keep after relevance filtering
_MAX_RELEVANT_CASES = 15
# Max chars of opinion text to fetch per case (no limit — full text)
_OPINION_SNIPPET_LENGTH = 0  # 0 = no truncation


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
        original_prompt: Optional[str] = None,
    ) -> ResearchResult:
        """Run the full research pipeline.

        Args:
            query: Natural language legal question or document context.
            llm_callback: Function(system_prompt, user_prompt) -> str for LLM calls.
            status_callback: Optional function(status_text) for progress updates.
            file_number: Optional case file number for caching.
            data_manager: Optional CaseDataManager instance for caching.
            original_prompt: Optional original user prompt text. Used to
                extract specific case names the user mentioned, even if they
                were stripped out during query extraction.

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

        # Also extract case names from the original prompt (if provided),
        # since query extraction may have stripped them out.
        if original_prompt:
            orig_cases = self._extract_case_names_regex(original_prompt)
            plan_cases = search_plan.get("requested_cases", [])
            seen = {c.lower() for c in plan_cases}
            for oc in orig_cases:
                if oc.lower() not in seen:
                    plan_cases.append(oc)
                    seen.add(oc.lower())
            search_plan["requested_cases"] = plan_cases

        # Phase 2: Parallel Source Search
        status("Searching legal databases...")
        cases, statutes, requested_cases, unfound_cases = self._search_sources(
            search_plan
        )

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
        result = ResearchResult(
            query=query,
            cases=cases,
            statutes=statutes,
            requested_cases=requested_cases,
            unfound_cases=unfound_cases,
        )
        authority_block = result.format_authority_block()
        memo = llm_callback(
            get_synthesis_prompt(),
            f"User's legal question: {query}\n\n{authority_block}",
        )
        result.memo = memo

        # Phase 6: LLM-based Verification
        if cases or statutes:
            status("Verifying citations...")
            known_names = result.get_known_case_names()
            result.verification = self._verify_citations(
                memo, authority_block, llm_callback, known_names
            )
            # Apply corrections from verification
            corrected = self._apply_corrections(result.verification, memo)
            if corrected:
                result.memo = corrected

        # Phase 7: Deterministic Citation Cross-Check (non-LLM safety net)
        status("Running deterministic citation check...")
        result.memo = self._deterministic_citation_check(
            result.memo, result.get_known_case_names()
        )

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
        raw = llm_callback(get_query_planning_prompt(), query)
        try:
            # Strip markdown code fences if present
            cleaned = raw.strip()
            if cleaned.startswith("```"):
                cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
                if cleaned.endswith("```"):
                    cleaned = cleaned[:-3]
                cleaned = cleaned.strip()
            plan = json.loads(cleaned)
        except (json.JSONDecodeError, ValueError):
            # Fallback: use the raw query as a single search term
            plan = {
                "case_queries": [query[:200]],
                "statute_queries": [],
                "legal_doctrines": [],
                "requested_cases": [],
            }

        # Also extract case names directly via regex as a fallback
        # (in case LLM query planning misses them)
        regex_cases = self._extract_case_names_regex(query)
        llm_cases = plan.get("requested_cases", [])
        # Merge, dedup by lowercase
        seen = {c.lower() for c in llm_cases}
        for rc in regex_cases:
            if rc.lower() not in seen:
                llm_cases.append(rc)
                seen.add(rc.lower())
        plan["requested_cases"] = llm_cases

        return plan

    @staticmethod
    def _extract_case_names_regex(text: str) -> List[str]:
        """Extract case names from text using regex (e.g., 'Smith v. Jones').

        This is a fallback for when LLM query planning misses case names.
        Handles patterns like:
          - Smith v. Jones
          - Polibrid Coatings, Inc. v. Superior Court
          - People v. Smith
        """
        # Match "Word(s) v. Word(s)" — stop at sentence boundaries
        # Allow commas and periods within names (Inc., Co., etc.)
        pattern = re.compile(
            r'([A-Z][A-Za-z\'\-]+(?:[\s,\.]+(?:Inc|Co|Corp|Ltd|LLC|L\.P|et al)[\.]?)*'
            r'(?:\s+[A-Z][A-Za-z\'\-]+(?:[\s,\.]+(?:Inc|Co|Corp|Ltd|LLC|L\.P|et al)[\.]?)*)*'
            r'\s+v\.?\s+'
            r'[A-Z][A-Za-z\'\-]+(?:[\s,\.]+(?:Inc|Co|Corp|Ltd|LLC|L\.P|et al)[\.]?)*'
            r'(?:\s+[A-Z][A-Za-z\'\-]+(?:[\s,\.]+(?:Inc|Co|Corp|Ltd|LLC|L\.P|et al)[\.]?)*)*)'
        )
        matches = pattern.findall(text)
        cleaned = []
        for m in matches:
            m = m.strip().rstrip(",. ")
            if 5 < len(m) < 120:
                cleaned.append(m)
        return cleaned

    def _search_sources(
        self, plan: Dict
    ) -> Tuple[List[CaseResult], List[StatuteResult], List[str], List[str]]:
        """Search all sources in parallel based on the query plan.

        Returns:
            Tuple of (cases, statutes, requested_cases, unfound_cases).
        """
        cases: List[CaseResult] = []
        statutes: List[StatuteResult] = []
        requested_cases: List[str] = plan.get("requested_cases", [])
        # Track which requested cases were found
        found_requested: Dict[str, bool] = {rc: False for rc in requested_cases}

        with ThreadPoolExecutor(max_workers=5) as pool:
            futures = {}

            # CourtListener searches (doctrine-based)
            for q in plan.get("case_queries", []):
                f = pool.submit(self.cl_client.search_opinions, q)
                futures[f] = ("case", q)

            # Direct case name searches for user-requested cases
            for case_name in requested_cases:
                # Strip punctuation that causes CourtListener 502 errors
                # (commas, periods in "Inc.", "Co.", etc.)
                clean_name = re.sub(r'[,\.]', '', case_name)
                # Strategy 1: caseName field search (most precise)
                f1 = pool.submit(
                    self.cl_client.search_opinions,
                    f'caseName:"{clean_name}"',
                    max_results=5,
                )
                futures[f1] = ("requested_case", case_name)
                # Strategy 2: unquoted fallback (broader)
                f2 = pool.submit(
                    self.cl_client.search_opinions,
                    clean_name,
                    max_results=5,
                )
                futures[f2] = ("requested_case", case_name)

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
                source_type, source_key = futures[future]
                try:
                    result = future.result()
                    if source_type in ("case", "recent_case"):
                        if isinstance(result, list):
                            cases.extend(result)
                    elif source_type == "requested_case":
                        if isinstance(result, list) and result:
                            cases.extend(result)
                            found_requested[source_key] = True
                        else:
                            logger.info(
                                "Requested case '%s' NOT FOUND on CourtListener",
                                source_key,
                            )
                    elif source_type == "statute":
                        if result is not None:
                            statutes.append(result)
                except Exception as e:
                    logger.warning("Source search error: %s", e)

        # Check if any requested cases were found via doctrine searches
        # (the user's case might appear in general results too)
        for rc in requested_cases:
            if not found_requested[rc]:
                rc_lower = rc.lower()
                for c in cases:
                    if self._case_name_matches(rc_lower, c.name.lower()):
                        found_requested[rc] = True
                        break

        unfound_cases = [rc for rc, found in found_requested.items() if not found]

        # Deduplicate cases by (name, citation)
        seen = set()
        unique_cases = []
        for c in cases:
            key = (c.name.lower(), c.citation.lower())
            if key not in seen:
                seen.add(key)
                unique_cases.append(c)

        return unique_cases, statutes, requested_cases, unfound_cases

    @staticmethod
    def _case_name_matches(requested: str, found: str) -> bool:
        """Check if a found case name matches a user-requested case name.

        Uses fuzzy matching: checks if key words from the requested name
        appear in the found name (handles "Inc." vs "Inc", extra parties, etc.).
        """
        # Extract the core party names (before and after "v." or "v")
        def core_words(name: str) -> List[str]:
            # Remove common suffixes and punctuation
            cleaned = re.sub(r'[,\.]', '', name.lower())
            # Split on "v." or "v "
            parts = re.split(r'\s+v\.?\s+', cleaned)
            words = []
            for p in parts:
                for w in p.split():
                    if len(w) > 2 and w not in ('the', 'inc', 'llc', 'ltd', 'corp', 'co'):
                        words.append(w)
            return words

        req_words = core_words(requested)
        if not req_words:
            return False
        # Require at least 60% of requested words to appear in found
        matches = sum(1 for w in req_words if w in found)
        return matches / len(req_words) >= 0.6

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

        raw = llm_callback(get_relevance_ranking_prompt(), user_prompt)
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
                        cluster_to_text[cid] = text
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
        known_case_names: Optional[List[str]] = None,
    ) -> List[VerificationStatus]:
        """Verify every citation in the draft against source data."""
        names_list = ""
        if known_case_names:
            names_list = (
                "\n\nKNOWN CASE NAMES (only these cases may be cited):\n"
                + "\n".join(f"  - {n}" for n in known_case_names)
            )
        prompt = (
            f"DRAFT TEXT:\n{draft}\n\n"
            f"SOURCE DATA:\n{authority_block}"
            f"{names_list}\n\n"
            f"Verify every citation in the draft. A citation is ONLY valid if "
            f"the case name appears in KNOWN CASE NAMES above."
        )
        raw = llm_callback(get_verification_prompt(), prompt)
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
    def _deterministic_citation_check(
        text: str, known_case_names: List[str]
    ) -> str:
        """Non-LLM safety net: tag any citation whose case name isn't in known sources.

        Extracts all case-law citations from text using regex, checks each
        against the known case names, and tags unverified ones inline plus
        appends a summary warning.
        """
        if not text or not known_case_names:
            return text

        # Common legal abbreviations for normalization
        _LEGAL_ABBREVS = {
            'ct': 'court', 'cts': 'courts',
            'corp': 'corporation', 'co': 'company',
            'dept': 'department', 'dist': 'district',
            'assn': 'association', 'bd': 'board',
            'comm': 'commission', 'commn': 'commission',
            'hosp': 'hospital', 'ins': 'insurance',
            'mfg': 'manufacturing', 'natl': 'national',
            'ry': 'railway', 'rr': 'railroad',
            'transp': 'transportation', 'univ': 'university',
            'govt': 'government', 'cty': 'county',
        }

        def normalize_name(name: str) -> str:
            """Normalize a case name for comparison."""
            cleaned = re.sub(r'[,\.\'\"\(\)]', '', name.lower()).strip()
            words = cleaned.split()
            normalized = []
            for w in words:
                w = _LEGAL_ABBREVS.get(w, w)
                normalized.append(w)
            return ' '.join(normalized)

        # Build a set of normalized known names for matching
        known_normalized = {normalize_name(n) for n in known_case_names}

        # Also build word sets for fuzzy matching
        def name_words(name: str) -> set:
            normed = normalize_name(name)
            return {w for w in normed.split() if len(w) > 2 and w not in (
                'the', 'inc', 'llc', 'ltd', 'corporation', 'company',
                'of', 'and', 'for',
            )}

        known_word_sets = [(n, name_words(n)) for n in known_case_names]

        def is_known(case_name: str) -> bool:
            """Check if case_name matches any known case."""
            cn_norm = normalize_name(case_name)
            # Exact normalized match
            if cn_norm in known_normalized:
                return True
            # Check if known name is contained in cited name or vice versa
            for kn in known_normalized:
                if kn in cn_norm or cn_norm in kn:
                    return True
            # Fuzzy: check word overlap
            cn_words = name_words(case_name)
            if not cn_words:
                return True  # Can't verify, don't flag
            for _, kw_set in known_word_sets:
                if not kw_set:
                    continue
                overlap = cn_words & kw_set
                # Require majority overlap in both directions
                if (len(overlap) / len(cn_words) >= 0.5 and
                        len(overlap) / len(kw_set) >= 0.3):
                    return True
            return False

        # Regex to find case citations: Name v. Name (Year) Vol Reporter Page
        # Name part allows commas (for "Inc.,"), dots. After the initial
        # uppercase word, continuation words must start with uppercase OR be
        # known entity suffixes (Inc., Co., LLC, etc.) to avoid capturing
        # preceding prose like "supported by".
        _entity_suffix = r'(?:Inc|Co|Corp|Ltd|LLC|L\.P|et\s+al|of|the|de|la|del|County|City|State|Dept|Dist|Assn|Bd|Hosp)'
        _name_part = (
            r'('
            r'[A-Z][A-Za-z\'\-\.]+'                             # First word (uppercase start)
            r'(?:[,]?\s+(?:[A-Z][A-Za-z\'\-\.]*|' + _entity_suffix + r'\.?))*'  # More words
            r'\s+v\.?\s+'                                        # " v. " or " v "
            r'[A-Z][A-Za-z\'\-\.]+'                              # First word of second party
            r'(?:[,]?\s+(?:[A-Z][A-Za-z\'\-\.]*|' + _entity_suffix + r'\.?))*'  # More words
            r')'
        )
        citation_pattern = re.compile(
            _name_part
            + r'\s*\(\d{4}\)\s*\d+\s+[A-Za-z][A-Za-z\.\d\s]*?(?:2d|3d|4th|5th|6th)\s+\d+'
            + r'|'
            + _name_part
            + r'\s*\(\d{4}\)\s*\d+\s+[A-Za-z][A-Za-z\.\s]+\d+'
        )

        unverified = []
        tagged_text = text

        for match in citation_pattern.finditer(text):
            # group(1) from first alternative, group(2) from second
            case_name = (match.group(1) or match.group(2) or "").strip().rstrip(",. ")
            full_citation = match.group(0).strip()

            if not is_known(case_name):
                # Check if already tagged
                tag = " [UNVERIFIED - NOT IN PROVIDED SOURCES]"
                if tag not in tagged_text[match.start():match.end() + len(tag)]:
                    tagged_text = tagged_text.replace(
                        full_citation, f"{full_citation}{tag}", 1
                    )
                    unverified.append(full_citation)

        # Append summary warning if any unverified citations found
        if unverified:
            warning = (
                "\n\n---\n"
                "CITATION VERIFICATION WARNING: The following citation(s) "
                "could NOT be verified against the provided legal sources and "
                "may be fabricated. DO NOT include in any court filing without "
                "manual verification by the attorney:\n"
            )
            for i, cit in enumerate(unverified, 1):
                warning += f"  {i}. {cit}\n"
            tagged_text += warning

        return tagged_text

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
                snippet_source=c.get("snippet_source", ""),
                snippet_parenthetical_id=c.get("snippet_parenthetical_id", ""),
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
            requested_cases=data.get("requested_cases", []),
            unfound_cases=data.get("unfound_cases", []),
        )
