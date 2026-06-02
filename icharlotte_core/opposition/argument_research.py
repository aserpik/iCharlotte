"""Retrieval-first grounding for opposition drafting.

Per argument: generate CourtListener search queries, hybrid-search CA case
law, fetch real opinion text for the top candidates, then have an LLM
re-rank/select the best 3-5 cases with a VERBATIM supporting passage.
Returns RetrievedAuthority records the drafter cites from.
"""

from __future__ import annotations

import concurrent.futures
import json
import logging
import os
import re
from typing import Any, Callable

from icharlotte_core.opposition.models import RetrievedAuthority

logger = logging.getLogger(__name__)

LLMCallback = Callable[[str, str], str]


def _loads_json(text: str) -> dict[str, Any]:
    if not isinstance(text, str):
        return {}
    cleaned = text.strip()
    fenced = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", cleaned, re.DOTALL | re.I)
    if fenced:
        cleaned = fenced.group(1).strip()
    try:
        data = json.loads(cleaned)
    except (TypeError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def generate_search_queries(argument: str, *, llm_callback: LLMCallback) -> list[str]:
    """Turn one argument into 1-2 CourtListener search queries."""
    from icharlotte_core.prompt_manager import get_prompt
    from icharlotte_core.opposition import prompts as default_prompts

    template = get_prompt("oppose_motion", "research_queries") or default_prompts.RESEARCH_QUERIES_PROMPT
    user_prompt = template.format(argument=argument or "")
    try:
        response = llm_callback("", user_prompt) or ""
    except Exception:
        logger.warning("research query generation failed", exc_info=True)
        return []
    data = _loads_json(response)
    raw = data.get("queries") if isinstance(data, dict) else None
    if not isinstance(raw, list):
        return []
    queries = [str(q).strip() for q in raw if str(q).strip()]
    return queries[:2]


def _normalize_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


# Generic, high-frequency words that carry no topical signal for locating the
# on-point part of an opinion (legal boilerplate + stopwords).
_EXCERPT_STOP = {
    "the", "and", "that", "this", "with", "from", "have", "has", "was", "were",
    "for", "are", "not", "but", "its", "his", "her", "their", "which", "such",
    "court", "courts", "case", "cases", "opinion", "plaintiff", "plaintiffs",
    "defendant", "defendants", "appellant", "respondent", "petitioner", "motion",
    "appeal", "trial", "order", "judgment", "party", "parties", "argues", "here",
    "under", "would", "should", "because", "exists", "proper", "subject",
}


def _relevant_excerpt(text: str, proposition: str, *, max_chars: int = 6000) -> str:
    """Return the ~max_chars window of ``text`` most relevant to ``proposition``.

    Opinion text starts with caption/procedural/factual material; the on-point
    holding is usually thousands of chars deep. Showing a reranker only the
    first N chars hides the very content that made the case match in retrieval,
    so it selects nothing. This slides a window over the opinion and picks the
    densest cluster of proposition key-terms.
    """
    text = text or ""
    if len(text) <= max_chars:
        return text
    terms = {t for t in re.findall(r"[a-z]{4,}", (proposition or "").lower())
             if t not in _EXCERPT_STOP}
    if not terms:
        return text[:max_chars]
    low = text.lower()
    hits: list[int] = []
    for t in terms:
        start = 0
        while True:
            i = low.find(t, start)
            if i < 0:
                break
            hits.append(i)
            start = i + len(t)
    if not hits:
        return text[:max_chars]
    hits.sort()
    import bisect
    best_start, best_score = 0, -1
    for h in hits:
        ws = max(0, h - 250)            # small lead-in for context
        we = ws + max_chars
        score = bisect.bisect_right(hits, we) - bisect.bisect_left(hits, ws)
        if score > best_score:
            best_score, best_start = score, ws
    end = min(len(text), best_start + max_chars)
    prefix = "…" if best_start > 0 else ""
    suffix = "…" if end < len(text) else ""
    return f"{prefix}{text[best_start:end]}{suffix}"


def _format_candidates(candidates: list[dict], *, proposition: str = "",
                       excerpt_chars: int = 6000) -> str:
    blocks: list[str] = []
    for c in candidates:
        excerpt = _relevant_excerpt(c.get("text") or "", proposition, max_chars=excerpt_chars)
        blocks.append(
            f"[{c.get('cluster_id')}] {c.get('case_name', '')}, {c.get('citation', '')}\n{excerpt}"
        )
    return "\n\n".join(blocks)


def select_authorities(
    proposition: str,
    candidates: list[dict],
    *,
    argument_text: str,
    argument_id: str = "",
    llm_callback: LLMCallback,
) -> list[RetrievedAuthority]:
    """LLM picks the best candidates; citation comes from metadata, and the
    quoted passage must appear verbatim in the candidate's opinion text."""
    from icharlotte_core.prompt_manager import get_prompt
    from icharlotte_core.opposition import prompts as default_prompts

    if not candidates:
        return []

    by_id = {str(c.get("cluster_id")): c for c in candidates}
    template = get_prompt("oppose_motion", "rerank_select") or default_prompts.RERANK_SELECT_PROMPT
    user_prompt = template.format(
        proposition=proposition or "",
        candidates=_format_candidates(candidates, proposition=proposition or ""),
    )
    try:
        response = llm_callback("", user_prompt) or ""
    except Exception:
        logger.warning("rerank/select failed", exc_info=True)
        return []

    data = _loads_json(response)
    selections = data.get("selections") if isinstance(data, dict) else None
    if not isinstance(selections, list):
        return []

    out: list[RetrievedAuthority] = []
    for sel in selections:
        if not isinstance(sel, dict):
            continue
        cand = by_id.get(str(sel.get("id")))
        if not cand:
            continue
        passage = str(sel.get("passage", "")).strip()
        if not passage or _normalize_ws(passage) not in _normalize_ws(cand.get("text", "")):
            continue  # drop unverifiable / fabricated passages
        out.append(
            RetrievedAuthority(
                argument_id=argument_id,
                argument_text=argument_text,
                cluster_id=str(cand.get("cluster_id") or ""),
                case_name=cand.get("case_name", ""),
                citation=cand.get("citation", ""),
                supports=str(sel.get("supports", "")).strip(),
                passage=passage,
                opinion_url=cand.get("opinion_url", ""),
            )
        )
    return out


def _load_cached_opinion(cache_dir: str | None, cluster_id: str) -> str | None:
    if not cache_dir or not cluster_id:
        return None
    path = os.path.join(cache_dir, f"{cluster_id}.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return (json.load(f) or {}).get("text") or None
    except (OSError, ValueError):
        return None


def _save_cached_opinion(cache_dir: str | None, cluster_id: str, text: str) -> None:
    if not cache_dir or not cluster_id:
        return
    try:
        os.makedirs(cache_dir, exist_ok=True)
        with open(os.path.join(cache_dir, f"{cluster_id}.json"), "w", encoding="utf-8") as f:
            json.dump({"cluster_id": cluster_id, "text": text}, f)
    except OSError:
        logger.warning("could not cache opinion %s", cluster_id, exc_info=True)


def _opinion_text(cl_client, cache_dir: str | None, cluster_id: str) -> str:
    cached = _load_cached_opinion(cache_dir, cluster_id)
    if cached is not None:
        return cached
    # Pass the id through as-is. CourtListener cluster_ids are numeric strings
    # that drop straight into the request URL; LocalCaseCorpus expects a string
    # case_uid (e.g. "cap:269732"). Forcing int() here broke the local path.
    try:
        text = cl_client.get_opinion_text(cluster_id) or ""
    except Exception:
        logger.warning("opinion fetch failed for %s", cluster_id, exc_info=True)
        text = ""
    if text:
        _save_cached_opinion(cache_dir, cluster_id, text)
    return text


def _hybrid_search(cl_client, query: str, max_results: int) -> list:
    """Union semantic + keyword results by cluster_id, semantic first."""
    found: dict[str, Any] = {}
    for semantic in (True, False):
        try:
            results = cl_client.search_opinions(
                query, semantic=semantic, max_results=max_results, published_only=True
            ) or []
        except Exception:
            logger.warning("search failed (semantic=%s)", semantic, exc_info=True)
            results = []
        for r in results:
            key = str(getattr(r, "cluster_id", "") or "")
            if key and key not in found:
                found[key] = r
    return list(found.values())


def _broaden(query: str) -> str:
    parts = query.split()
    return " ".join(parts[:-1]) if len(parts) > 1 else query


def research_argument(
    argument: str,
    *,
    cl_client,
    query_llm: LLMCallback,
    rerank_llm: LLMCallback,
    argument_id: str = "",
    max_candidates: int = 20,
    fetch_top: int = 8,
    cache_dir: str | None = None,
) -> list[RetrievedAuthority]:
    """Research one argument end-to-end; returns selected RetrievedAuthority."""
    queries = generate_search_queries(argument, llm_callback=query_llm)
    if not queries:
        queries = [argument]

    def _run(query_list: list[str]) -> list[RetrievedAuthority]:
        candidates: dict[str, Any] = {}
        for q in query_list:
            for r in _hybrid_search(cl_client, q, max_candidates):
                key = str(getattr(r, "cluster_id", "") or "")
                if key and key not in candidates:
                    candidates[key] = r
        ordered = list(candidates.values())[:fetch_top]
        cand_dicts: list[dict] = []
        for r in ordered:
            cid = str(getattr(r, "cluster_id", "") or "")
            text = _opinion_text(cl_client, cache_dir, cid)
            if not text:
                continue
            cand_dicts.append({
                "cluster_id": cid,
                "case_name": getattr(r, "name", ""),
                "citation": getattr(r, "citation", ""),
                "text": text,
                "opinion_url": getattr(r, "url", ""),
            })
        return select_authorities(
            argument, cand_dicts, argument_text=argument,
            argument_id=argument_id, llm_callback=rerank_llm,
        )

    selected = _run(queries)
    if not selected:
        broadened = _broaden(queries[0])
        if broadened and broadened != queries[0]:
            selected = _run([broadened])

    # Stamp the soft good-law hint (citation count + latest citing year) on the
    # final selected authorities only — a bounded number of extra calls. Guarded
    # so a non-dict return (e.g. a test MagicMock) is a no-op.
    for ra in selected:
        try:
            signals = cl_client.get_authority_signals(ra.cluster_id)
        except Exception:
            signals = None
        if isinstance(signals, dict):
            ra.citation_count = signals.get("citation_count")
            ra.latest_citing_year = signals.get("latest_citing_year", "")
    return selected


ProgressCallback = Callable[[str], None]


def research_arguments(
    arguments: list[str],
    *,
    cl_client,
    query_llm: LLMCallback,
    rerank_llm: LLMCallback,
    max_workers: int = 4,
    on_progress: ProgressCallback | None = None,
    cache_dir: str | None = None,
) -> list[RetrievedAuthority]:
    """Research every argument in parallel; flatten the RetrievedAuthority list."""
    args = [a.strip() for a in (arguments or []) if a and a.strip()]
    if not args:
        return []

    def _one(idx_arg: tuple[int, str]) -> tuple[int, list[RetrievedAuthority]]:
        idx, arg = idx_arg
        result = research_argument(
            arg, cl_client=cl_client, query_llm=query_llm, rerank_llm=rerank_llm,
            argument_id=f"arg-{idx}", cache_dir=cache_dir,
        )
        if on_progress:
            if result:
                on_progress(f"  {arg[:60]} — {len(result)} case(s) found")
            else:
                on_progress(f"  {arg[:60]} — no on-point authority retrieved")
        return idx, result

    indexed = list(enumerate(args))
    by_index: dict[int, list[RetrievedAuthority]] = {}
    workers = max(1, int(max_workers))
    if workers == 1:
        for pair in indexed:
            i, res = _one(pair)
            by_index[i] = res
    else:
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
            for i, res in pool.map(_one, indexed):
                by_index[i] = res

    flat: list[RetrievedAuthority] = []
    for i, _arg in indexed:
        flat.extend(by_index.get(i, []))
    return flat
