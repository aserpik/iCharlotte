# Oppose-a-Motion: Citation Grounding (Retrieval-First) — Design

**Date:** 2026-05-27
**Status:** Approved in brainstorming; pending written-spec review
**Builds on:** `2026-05-26-oppose-motion-redesign-design.md` (back-loaded verification pipeline, now shipped)

---

## Goal

Make the caselaw cited in the *Oppose a Motion* wizard task correct by construction, rather than caught after the fact. Today most citations are flagged by the verifier because the drafter writes them from the LLM's own memory of California case law — and LLMs cannot reliably recall exact reporter citations *and* holdings, so they confabulate. The verifier is working; the failure is upstream, in **generation**.

This redesign adds **retrieval-first grounding**: before drafting, the pipeline researches real California authority for each argument against CourtListener, then hands the drafter actual opinion text and citations pulled from CourtListener metadata. The drafter may cite **only** from that retrieved pool. The existing verifier is retained as a safety net and should now run almost entirely green.

This is the Harvey / CoCounsel / Lexis+ AI pattern (retrieve → re-rank → grounded generation), implemented on the corpus and search we already have access to: CourtListener.

---

## What changed since the last design

The 2026-05-26 redesign deliberately chose **back-loaded** verification (draft from memory → verify after) over front-loaded grounding, citing speed (~5–7 min estimate) and a preference for attorney control. Two facts change that calculus:

1. **CourtListener shipped a hosted semantic-search API (Nov 2025).** A GET request with `semantic=true` embeds the query server-side (768-dim fine-tuned ModernBERT, the "Citegeist" engine) and returns nearest-neighbor opinions with the same `court` / date / precedential-status filters as keyword search, plus per-hit snippets. POST accepts a pre-computed 768-dim embedding for local/private embedding. There is no local vector index to build or maintain.
   - Announcement: https://free.law/2025/11/05/semantic-search-api/
   - API wiki: https://wiki.free.law/c/courtlistener/help/api/rest/v4/search
2. **Our client already does most of the work.** `icharlotte_core/legal_research/sources/courtlistener.py` already has `search_opinions()` (keyword, CA-court-filtered), `get_opinion_text()` (robust multi-field full-text fetch with caching), `lookup_citations()`, `get_cluster()`, `get_citing_cases()`, and `_send_with_retry()` (429/5xx exponential backoff). Adding semantic search is a one-parameter extension.

With the semantic-search blocker gone, retrieval becomes one fast API call per argument, and the speed objection to front-loading no longer holds.

---

## Pipeline Architecture

```
1. analyze_motion          (existing)
        ↓
2. generate_outline        (existing)
        ↓
3. research_arguments      (NEW — per-argument retrieval + re-rank, parallel)
        ↓
4. draft_memorandum        (REWRITTEN — single call, labeled global authority pool, cite-only-from-pool)
        ↓
   citation_parser         (existing — now also runs the pool-membership check)
        ↓
5. verify_citation         (existing — safety net; should be ~all green)
```

Stages 1, 2, and 5 are unchanged from the shipped pipeline. Stage 3 is new. Stage 4 is rewritten. The citation parser gains a pool-membership check.

### Drafting structure decision

The drafter is called **once**, with a single **labeled global authority pool** (all per-argument retrieval results assembled into one block, grouped by which argument each case supports), alongside the existing style exemplars. This was chosen over section-by-section drafting to preserve prose flow, cross-section transitions, and the style-exemplar voice investment, while still grounding per-argument because the pool is labeled.

---

## Stage 3 — `research_arguments` (new)

New module: `icharlotte_core/opposition/argument_research.py`.

For each principal argument / selected outline section, run a three-step retrieve → fetch → re-rank cycle. Arguments are processed in **parallel** via a bounded `ThreadPoolExecutor` (default 4 workers, mirroring `OppositionVerifier`).

### 3a. Query generation (`research_queries` prompt — new)

A workbench-editable prompt turns one argument (its heading + the proposition the opposition will assert) into **1–2 CourtListener search queries**, mixing terms of art with natural-language phrasing. Output is strict JSON: `{"queries": ["...", "..."]}`.

### 3b. Hybrid search (CourtListener)

Extend `CourtListenerClient` with semantic support. Proposed signature:

```python
def search_opinions(
    self,
    query: str,
    *,
    semantic: bool = False,
    jurisdiction: str = "cal",
    max_results: int = 15,
    published_only: bool = True,
) -> List[CaseResult]:
```

- When `semantic=True`, add `semantic: "true"` to the GET params; otherwise keyword as today.
- Always filter to `court=CA_COURTS` (the existing constant) and, when `published_only`, precedential status published.
- For each argument, run **both** a semantic pass and a keyword pass per query and **union the candidates by `cluster_id`** (hybrid recall). Target ~15–25 unique candidates per argument.
- `CaseResult` already carries `cluster_id`, `name`, `citation`, `date`, `court`, `snippet`, `url`.

### 3c. Opinion text fetch

For the top ~6–8 candidates per argument (ranked by CourtListener relevance + the soft good-law signal below), fetch full opinion text via the existing `get_opinion_text(cluster_id)`, reusing the on-disk `.cache/opinions/{cluster_id}.json` cache. Candidates whose text cannot be retrieved are dropped from consideration (a citation we cannot read is one we will not ground a draft on).

### 3d. Re-rank + select (`rerank_select` prompt — new)

A workbench-editable prompt receives the argument's proposition and the fetched candidate texts (truncated per token budget) and selects the **best 3–5 cases**. For each selected case it returns:

- `cluster_id` and the citation **as it appears in CourtListener metadata** (never model-generated),
- `case_name`,
- `supports`: one sentence stating the proposition this case supports,
- `passage`: a **verbatim** quote from the opinion text that establishes the holding.

Output is strict JSON. Any selected case whose `passage` is not found (substring, normalized) in the fetched opinion text is dropped — this prevents the re-ranker from fabricating a supporting quote.

### 3e. Empty-result handling

If an argument yields zero usable candidates after re-rank, retry **once** with a single broadened query (drop the most specific term). If still empty, the argument proceeds to drafting with **no case authority** and is marked so the drafter can emit the soft no-authority marker (see Stage 4). Deeper multi-round agentic query refinement is **out of scope v1**.

### 3f. Soft good-law signal

True KeyCite/Shepard's good-law checking is **out of scope** (CourtListener exposes no clean "overruled" flag). As a cheap proxy:

- During candidate ranking (3c), prefer cases with higher citation counts (available from CourtListener) — well-cited cases are less likely to be obscure or bad.
- Surface **citation count** and **most-recent-citing-year** (from `get_citing_cases`, capped to a small fetch) on the selected cases, carried through to the cite-detail panel as a "freshness" hint.

This catches "obscure / likely-stale" but **not** "formally overruled" — that remains a manual attorney check, consistent with the shipped design's out-of-scope list.

### Output model

`research_arguments` returns a structure consumable by the drafter and persistable for the wizard. Proposed dataclass in `models.py`:

```python
@dataclass
class RetrievedAuthority:
    argument_id: str = ""          # outline node id / argument index
    argument_text: str = ""        # heading or proposition researched
    cluster_id: str = ""
    case_name: str = ""
    citation: str = ""             # from CourtListener metadata, not generated
    supports: str = ""             # one-sentence proposition this case supports
    passage: str = ""              # verbatim opinion quote
    opinion_url: str = ""
    citation_count: int | None = None
    latest_citing_year: str = ""
```

The full retrieval result is `list[RetrievedAuthority]` (multiple per argument).

---

## Stage 4 — `draft_memorandum` (rewritten)

`icharlotte_core/opposition/drafter.py` gains a `retrieved_authorities: list[RetrievedAuthority]` input and a rewritten prompt.

### Labeled authority block

The drafter prompt assembles retrieved authorities into one block, grouped by argument:

```
AUTHORITY POOL (cite ONLY from this list; use the citation exactly as written):

For "Argument 1 heading":
  - <case_name>, <citation>
    Supports: <supports>
    Holding (verbatim): "<passage>"
  - ...

For "Argument 2 heading":
  - ...
```

### Prompt rules (replacing the memory-based citation rules)

- **Cite cases ONLY from the authority pool.** Use each citation exactly as written; do not alter, abbreviate, or invent reporter cites. Never cite a case not in the pool.
- For each case cited, ground the in-text parenthetical/summary in the provided `passage` — do not assert a holding the passage does not support.
- **No-authority handling:** if no pooled case supports a proposition, argue it from the controlling statute and the logic of the motion's own admissions, and append the inline marker `[no case authority retrieved for this point]` so the attorney can supply one. Never invent a case to fill the gap.
- **Statutes** are drafted from knowledge as today (CourtListener does not index CA statutes; statute hallucination is rarer) and remain independently verified against leginfo by the existing `StatuteVerifier`.
- Existing hardening (wrong-side detection, forbidden-output detection, prompt-injection resistance, JSON-only output) is retained unchanged.

### Post-draft pool-membership check

This runs as a deterministic step in the verifier orchestration, after `citation_parser` extracts cites and **before** the network verification fans out (so off-pool cites short-circuit without an API call). The parser stays extraction-only; the pool join lives in the orchestration layer alongside `OppositionVerifier`.

Each extracted case cite is matched against the retrieved pool by normalized citation / cluster lookup:

- **In pool** → carries forward to verification normally (expected to verify SUPPORTED).
- **Not in pool** → immediately assigned `verdict="NOT_FOUND"` with note `"Cited a case that was not in the researched authority pool — likely model-introduced; verify or replace."` This is a deterministic anti-hallucination backstop that does not depend on the LLM verifier.

The soft `[no case authority retrieved for this point]` markers are surfaced in the output page as a distinct (gray) inline flag, not a red citation.

---

## Stage 5 — verification (existing, unchanged)

`citation_parser → OppositionVerifier` runs exactly as shipped: case path against CourtListener, statute path against leginfo, dedup + bounded ThreadPool, verdict-colored underlines, cite-detail right panel. With grounding in place this stage should be almost entirely SUPPORTED; remaining flags are rare and high-signal.

The existing (ungrounded) `find_replacement` button stays as-is for v1. A future enhancement can re-point it at `research_arguments` so replacements are grounded too; not required here.

---

## Workbench / Config Integration

### New prompts (workbench-editable, seeded by `PromptManager.seed_pipeline_prompts()`)

Added to `icharlotte_core/opposition/prompts.py` and the `oppose_motion` prompt directory:

- `research_queries_current.txt` — argument → 1–2 search queries (JSON).
- `rerank_select_current.txt` — candidates → best 3–5 with verbatim passage (JSON).

`draft_memorandum_current.txt` is updated to the cite-only-from-pool version above.

### LLMConfig

`agent_oppose_motion` gains two passes: `research_queries`, `rerank_select`. Per-pass model preference:

- `research_queries`, `rerank_select` → fast/cheap model (e.g. Gemini Flash).
- `draft_memorandum` → stronger model (unchanged).

### Caching

- Opinion text: reuse existing `.cache/opinions/{cluster_id}.json`.
- Search results: optional new `.cache/search/{md5(query+filters)}.json` to avoid repeat searches within a session (TTL or session-scoped; gitignored).

---

## UI / Wizard Changes

`icharlotte_core/ui/wizard/pages/oppose_motion_page.py`:

- **Status page** gains a "Researching authorities (N arguments)…" phase between outline and drafting, emitting per-argument progress lines:
  ```
  Researching authorities (4 arguments)...
    ✓ Discovery cutoff — 4 cases found
    ✓ Good cause standard — 3 cases found
    ⚠ Sanctions — no on-point authority retrieved
  ```
- **Output page** is structurally unchanged. Two additions:
  - Pool-check red flags appear in the existing verification summary and underlines.
  - `[no case authority retrieved for this point]` markers render as a distinct gray inline note.
  - The cite-detail right panel shows the soft good-law hint (citation count + latest-citing-year) for grounded cases.

No change to the Settings page or the three-page flow structure.

---

## Module Summary

### New
- `icharlotte_core/opposition/argument_research.py` — query-gen, hybrid search, fetch, re-rank, parallel orchestration; returns `list[RetrievedAuthority]`.

### Modified
- `icharlotte_core/legal_research/sources/courtlistener.py` — `search_opinions(..., semantic=bool, published_only=bool)`; small helper for citation count / latest-citing-year.
- `icharlotte_core/opposition/drafter.py` — accept `retrieved_authorities`; rewritten prompt; load from `PromptManager`.
- `icharlotte_core/opposition/prompts.py` — new `RESEARCH_QUERIES_PROMPT`, `RERANK_SELECT_PROMPT`; updated `DRAFT_MEMORANDUM_PROMPT`.
- `icharlotte_core/opposition/verifier.py` — pool-membership check (deterministic NOT_FOUND for off-pool cites) as a pre-verification step; `citation_parser.py` stays extraction-only.
- `icharlotte_core/opposition/models.py` — add `RetrievedAuthority`.
- `icharlotte_core/llm_config.py` — register `research_queries`, `rerank_select` passes.
- `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` — research phase progress; no-authority + good-law-hint rendering.

### Unchanged
- `case_verifier.py`, `statute_verifier.py`, `verifier.py` (verification stays the safety net).

---

## Testing Strategy

### Unit (mocked CourtListener / LLM)
- `tests/test_opposition/test_argument_research.py` — query-gen JSON parsing; hybrid union by `cluster_id`; re-rank selection; **verbatim-passage drop** when quote not in fetched text; empty-result single retry; parallel orchestration returns per-argument results.
- `tests/test_opposition/test_courtlistener_semantic.py` — `semantic=true` param added on semantic path; published filter; keyword path unchanged.
- `tests/test_opposition/test_drafter.py` (extend) — labeled authority block assembly; cite-only-from-pool prompt content; `RetrievedAuthority` plumbing; rejection paths preserved.
- `tests/test_opposition/test_pool_check.py` — off-pool case cite → deterministic NOT_FOUND; in-pool cite passes through; no-authority marker handling.
- Existing parser/verifier/page tests stay green.

### Integration (gated; skipped in CI)
- End-to-end against the existing Pinscreen MTC motion: drive analyze → outline → research → draft → verify. Assert (a) each argument receives retrieved authorities or a logged no-authority result, (b) every case cite in the draft is in the pool, (c) the verification report shows a large majority SUPPORTED. Runs only when `COURTLISTENER_API_TOKEN` and `GEMINI_API_KEY` are set.

All tests run with the venv interpreter `C:\geminiterminal2\.venv\Scripts\python.exe` (system Python lacks bs4 / PySide6).

---

## Rollout Notes

- Additive: new passes/prompts are seeded on first run; existing `oppose_motion` workbench entry already exists.
- The app runs from the `C:\geminiterminal2\` main checkout, **not** this worktree — edits must be applied there (or merged) and iCharlotte restarted to take effect.
- No data migration: the task persists only the wizard preview .docx, which is unchanged in shape.

---

## Out of Scope (v1)

- Multi-round agentic query refinement (only a single broadening retry on empty results).
- Statute retrieval grounding (statutes stay knowledge-drafted + leginfo-verified).
- True citator / "is this still good law" (only a soft citation-count + recency hint).
- Pin-cite page-number verification.
- Federal authority, local court rules, treatises, California Constitution.
- Re-pointing the `find_replacement` button at grounded retrieval (kept ungrounded for v1).
- Path A (download CourtListener's ~2 TB bulk embeddings into a local index) — hosted `semantic=true` makes it unnecessary for this use case.
