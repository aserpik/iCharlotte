# Agentic Legal Deep Research Orchestrator - Design

**Date:** 2026-06-06
**Status:** Draft for user review
**Primary package:** `icharlotte_core/legal_research/`
**Consumers:** Chat, Word Assistant, Oppose a Motion, Generate a Motion

## Goal

Build a shared, Westlaw-Deep-Research-inspired legal research backend for
iCharlotte. The goal is not to clone Westlaw or replace KeyCite. The goal is to
make iCharlotte's legal research behave more like a careful legal researcher:
plan the work, search trusted sources, refine when results are thin, surface
conflicting authority, cite only retrieved material, and show the user how the
answer was built.

The new backend should produce a reusable `ResearchRun` / `ResearchPacket`
object that downstream workflows can inject into chat answers, motion drafts,
brief outlines, and Word Assistant prompts.

## Inputs From The White Paper

The Thomson Reuters white paper does not disclose implementation internals, but
it does clarify the product architecture we should emulate:

- Deep research is an iterative loop: set goal, create plan, apply tools,
  review, refine, synthesize, cite sources.
- Trust comes from source grounding plus a visible research log.
- Negative treatment, conflicting authority, and competing interpretations are
  treated as first-class outputs.
- Statutory issues use a specialized path: statute text, annotations or related
  authority, and cases applying the statute.
- The output is strategy-facing, not just a list of search results.

iCharlotte should translate those ideas into a source-transparent research
orchestrator over the sources we actually have: local CA case law, CourtListener,
CA LegInfo, CA courts, and firm-brief authority.

## Existing Context

iCharlotte already has the main building blocks:

- `icharlotte_core/legal_research/engine.py` handles query planning, parallel
  source search, relevance filtering, opinion enrichment, synthesis, LLM
  verification, and deterministic citation checks.
- `icharlotte_core/legal_research/local_corpus/` provides an offline California
  corpus with FTS5 BM25 plus semantic rerank and soft authority signals.
- CourtListener parentheticals, once ingested into the local corpus, should be
  treated as a separate secondary signal from opinion text. They are useful for
  ranking, issue matching, and explaining why a case matters, but they should
  not be overweighted or mislabeled as verbatim opinion holdings.
- `icharlotte_core/legal_research/sources/courtlistener.py` provides live
  CourtListener search, citation lookup, opinion fetch, and California reporter
  citation preference.
- `icharlotte_core/opposition/argument_research.py` already performs
  proposition-level retrieval, semantic/keyword search union, LLM reranking, and
  verbatim quote checks.
- `icharlotte_core/firm_briefs/` provides firm-brief citation reuse and style
  indexing.
- `icharlotte_core/opposition/verifier.py`, `case_verifier.py`,
  `local_case_verifier.py`, and `statute_verifier.py` already verify cited
  authorities after drafting.
- `docs/superpowers/specs/2026-06-05-chat-legal-research-design.md` already
  defines a Chat-specific source selector and research packet.

This design should not replace those pieces. It should make them available
through one shared headless orchestration layer.

## Approaches Considered

### Option A: Extend `LegalResearchEngine` in place

This is the smallest change because the engine already performs a linear
research pipeline. It is also the riskiest long-term path: the current engine
would become a large mixed-purpose module handling search, planning, source
selection, iterative refinement, logging, conflict checks, and consumer-specific
packet formatting.

### Option B: Build separate deep-research flows per surface

Chat, Word Assistant, Oppose Motion, and Generate Motion could each get their
own improved flow. This keeps each UI surface simple in the short term, but it
would duplicate source-selection, quote verification, adverse-authority logic,
and research logs. It also increases the chance that one surface is more
trustworthy than another.

### Option C: Add a shared agentic orchestrator with adapters

Create a Qt-free `legal_research.deep_research` package that owns research
planning, iterative retrieval, conflict detection, statutory analysis,
verification, logging, and packet construction. Existing surfaces call it
through thin adapters. This is the recommended path because it keeps the
research policy in one place while allowing each workflow to shape the same
verified packet for its own output.

## Scope

The first version should support California civil litigation research for:

- discrete chat questions;
- Word Assistant research augmentation;
- Generate Motion authority support;
- Oppose Motion authority support;
- future case assessment or memo development workflows.

The orchestrator must be headless and testable without Qt. UI surfaces should
own controls, progress display, and final rendering only.

## Non-Goals

This design does not:

- provide Westlaw, Practical Law, KeyCite, or Key Number access;
- claim Shepard's or KeyCite equivalence;
- scrape Westlaw or any licensed proprietary database;
- change local corpus build mechanics;
- replace the firm-brief ingestion system;
- remove existing motion-specific citation review;
- allow the model to cite authority from memory.

## Architecture

Add a package:

```text
icharlotte_core/legal_research/deep_research/
  __init__.py
  models.py
  orchestrator.py
  planning.py
  sources.py
  retrieval.py
  ranking.py
  conflicts.py
  statutes.py
  verification.py
  packets.py
  cache.py
```

### Main API

```python
def run_deep_research(
    request: DeepResearchRequest,
    llm_callback: LLMCallback,
    source_registry: SourceRegistry,
    status_callback: StatusCallback | None = None,
) -> ResearchRun:
    ...
```

The orchestrator returns a full `ResearchRun`, not just a memo. UI surfaces can
then derive:

- a prompt-ready authority packet;
- a user-facing research log;
- a citation review payload;
- a drafting authority pool;
- warnings and failure reasons.

## Core Models

### `DeepResearchRequest`

Fields:

- `surface`: `chat`, `word_assistant`, `oppose_motion`, `generate_motion`, or
  `case_assessment`.
- `task_type`: `discrete_question`, `motion_argument`, `brief_section`,
  `statutory_interpretation`, or `mixed`.
- `jurisdiction`: default `California`.
- `side`: optional litigation side, such as `moving`, `opposition`, or
  `neutral`.
- `question`: the user question or proposition to research.
- `matter_context`: optional facts, motion metadata, claims, defenses, or
  selected document text.
- `source_policy`: selected sources and fallback rules.
- `freshness_policy`: whether current law is required.
- `max_questions`: default 5.
- `max_iterations`: default 2 for chat/Word, 3 for motion workflows.
- `fail_closed`: default true when a workflow claims research-backed output.

### `ResearchRun`

Fields:

- `run_id`;
- `request`;
- `status`: `complete`, `partial`, `failed`;
- `plan`;
- `questions`;
- `steps`;
- `searches`;
- `candidates`;
- `selected_authorities`;
- `statutory_materials`;
- `adverse_authorities`;
- `treatment_signals`;
- `conflict_analysis`;
- `citation_audit`;
- `synthesis`;
- `warnings`;
- `packet`;
- `diagnostics`.

### `ResearchPlan`

The plan should be structured and visible. It should include:

- core legal questions;
- required source types: cases, statutes, rules, firm authority, and future
  secondary-source adapters if iCharlotte later gains a licensed or local
  secondary-source corpus;
- search strategy per question;
- jurisdiction and date constraints;
- whether current-law fallback is needed;
- expected output format.

### `ResearchStep`

Each step records:

- `phase`: planning, search, ranking, refinement, conflict_check,
  statute_analysis, synthesis, verification;
- `input`;
- `tool_or_source`;
- `output_summary`;
- `decision`;
- `warnings`;
- timestamp and duration.

This becomes the research log.

### `AuthorityCandidate`

Fields:

- `candidate_id`;
- `source`: `firm`, `local_corpus`, `courtlistener`, `ca_leginfo`,
  `ca_courts`;
- `case_name`, `citation`, `year`, `court`;
- `cluster_id` or local `case_uid`;
- `source_url` or local reference;
- `snippet`;
- `full_text_available`;
- `proposition_match`;
- `retrieval_score`;
- `semantic_score`;
- `recency_score`;
- `authority_signal_score`;
- `citation_count`;
- `latest_citing_year`;
- `negative_signal`;
- `parentheticals`: short CourtListener parentheticals tied to this authority;
- `parenthetical_match_score`: similarity between parentheticals and the
  researched proposition, capped as a secondary ranking signal;
- `treatment_signals`: parenthetical or citation-edge records from later cases
  describing how this authority was used;
- `provenance`.

### `SelectedAuthority`

Fields:

- all stable citation metadata;
- `supports`: one-sentence proposition;
- `verbatim_quote`;
- `quote_location` if available;
- `selection_reason`;
- `limitations`;
- `adverse_or_distinguishable`: boolean;
- `parenthetical_summary`: optional CourtListener parenthetical selected as a
  research note;
- `parenthetical_source`: the citing case or bulk-data source for that
  parenthetical;
- `verification_status`;
- `alternatives`.

The system must drop any selected authority whose quote cannot be found in the
retrieved source text after whitespace normalization, unless it is explicitly
marked as unverified firm-brief material and excluded from verified-citation
prompt packets.

### `TreatmentSignal`

CourtListener parentheticals should be modeled as treatment or explanation
signals, not as ordinary snippets. A single case can have many parentheticals
from later citing decisions, and those parentheticals may characterize the case
in different ways.

Fields:

- `signal_id`;
- `source`: `courtlistener_parenthetical`, `citation_edge`, or future source;
- `described_case_uid` or `described_cluster_id`;
- `described_citation`;
- `citing_case_uid` or `citing_cluster_id`;
- `citing_case_name`;
- `citing_citation`;
- `citing_year`;
- `citing_court`;
- `parenthetical_text`;
- `depth` or citation-edge weight, if available;
- `classification`: `supporting`, `limiting`, `distinguishing`, `contrary`,
  `background`, or `unknown`;
- `confidence`;
- `provenance`.

The first implementation can classify parentheticals with deterministic keyword
rules plus an LLM fallback for ambiguous cases. Classification should be treated
as advisory unless independently verified by opinion text.

## Source Registry

Introduce a `SourceRegistry` that creates and normalizes source clients.

Sources:

- `FirmAuthoritySource`: wraps `FirmAuthorityProvider`.
- `LocalCorpusSource`: wraps `LocalCaseCorpus`.
- `CourtListenerSource`: wraps `CourtListenerClient`.
- `CALegInfoSource`: wraps `CALegInfoClient`.
- `CACourtsSource`: wraps `CACourtsClient`.

Each source adapter should expose:

```python
search(question: ResearchQuestion, query: SearchQuery) -> list[AuthorityCandidate]
fetch(candidate: AuthorityCandidate) -> AuthorityCandidate
fetch_treatment(candidate: AuthorityCandidate) -> list[TreatmentSignal]
verify(candidate: AuthorityCandidate) -> VerificationSignal
```

Adapters should not know about Qt or drafting.

## Source Policy

Use explicit source settings instead of implicit fallback.

Recommended defaults:

- firm authority: on when index exists;
- local corpus: on when corpus exists and is fresh enough;
- CourtListener: fallback/current-law;
- CA LegInfo: on for statutory references;
- CA courts recent opinions: fallback/current-law;
- fail closed: on for research-backed legal answers.

CourtListener modes:

- `off`: never call it.
- `fallback_current_law`: call only when local results are stale, thin, missing,
  or the question asks for current law.
- `always_search`: search live for each question.

If a selected source is unavailable, the run records a warning. If no verified
source remains and `fail_closed` is true, the run fails before producing a
research-backed answer.

## Research Flow

### Phase 1: Normalize The Question

Use the LLM to convert the user prompt and matter context into focused legal
questions. The output must be JSON:

- `questions`;
- `jurisdiction`;
- `material_facts`;
- `statutory_refs`;
- `desired_side`;
- `excluded_noise_terms`;
- `clarifying_questions` if the prompt is too ambiguous.

The planner should remove party names and irrelevant facts unless they are
legally material, matching the Westlaw guidance that excess names and facts can
distort search.

### Phase 2: Build A Research Plan

For each question, decide:

- search source mix;
- case-law queries;
- statutory queries;
- firm-authority query;
- current-law need;
- whether adverse authority is required;
- whether follow-up search is likely.

The plan is shown in the research log.

### Phase 3: Initial Retrieval

Run source searches in bounded parallelism:

- local corpus semantic search first for natural-language propositions;
- keyword or Boolean-style searches where useful;
- firm-brief citation/proposition search;
- live CourtListener according to policy;
- CA LegInfo for statutes;
- CA courts recent opinions when current-law fallback is triggered.

Deduplicate by:

- CourtListener cluster id;
- local corpus case uid;
- normalized reporter citation;
- normalized case name plus year.

Preserve every source that found the authority.

### Phase 4: Candidate Enrichment

Fetch full opinion text or reliable excerpts for the best candidates. Record
failures in the log, but do not cite candidates without source text unless they
are clearly marked unverified and excluded from verified packets.

When CourtListener parentheticals are available in the local corpus, enrichment
should also fetch the best parentheticals for each candidate. Parentheticals
should be stored as structured treatment or explanatory signals:

- parenthetical text;
- cited authority;
- citing case, if known;
- citing case date and court, if known;
- citation depth or citation-edge weight, if available;
- whether the parenthetical appears to support, limit, distinguish, or merely
  describe the authority.

Parentheticals are not the same as opinion text. The packet can display them as
research notes and use them for ranking, but the final LLM must not present a
parenthetical as a quote from the cited opinion.

### Phase 5: Ranking And Selection

Use deterministic scoring plus LLM reranking.

Deterministic ranking signals:

- semantic similarity to the proposition;
- keyword match;
- jurisdiction;
- court level;
- publication/citable status;
- recency when current law matters;
- citation count and latest citing year;
- parenthetical match to the researched proposition;
- number and quality of later-case parentheticals describing the authority;
- firm-prior usage;
- source count, meaning the authority was found by multiple sources;
- negative or stale signals.

Parenthetical weighting rules:

- Parenthetical signals are advisory and should be capped at a small share of
  the deterministic candidate score. The default cap should be 10 percent of
  the deterministic score before LLM reranking.
- Parentheticals may break ties, help choose candidates for the reranker, and
  explain why an authority deserves closer review.
- Parentheticals may not be the sole reason a case is selected as cited
  authority.
- A candidate with matching parentheticals but no direct opinion-text support
  must not outrank a candidate with verified on-point opinion text.
- Parenthetical classification confidence should reduce, not increase, weight
  when the classification is ambiguous or based only on weak keyword rules.
- Multiple near-duplicate parentheticals should be deduplicated or dampened so a
  heavily repeated boilerplate parenthetical does not swamp better direct
  authority.

The LLM reranker must return:

- selected candidate id;
- supported proposition;
- verbatim quote;
- reason selected;
- limitation or caveat;
- whether it is adverse, distinguishable, or supporting.

The verifier rejects non-verbatim quotes.

### Phase 6: Review And Refine

After initial selection, the orchestrator evaluates gaps:

- no authority for a required proposition;
- only old authority when current law is requested;
- only trial-level, unpublished, or weak authority;
- no statute when statute is central;
- no adverse authority for motion practice;
- selected authority has negative/stale signals.

If gaps exist and `max_iterations` remains, generate follow-up searches and log
why the refinement occurred.

### Phase 7: Conflict And Adverse Authority Pass

Run an explicit pass for:

- cases reaching a contrary result;
- cases distinguishing or limiting the selected authority;
- negative treatment signals;
- parentheticals from later cases that characterize the selected authority in a
  limiting, distinguishing, or contrary way;
- more recent cases citing the selected authority;
- competing interpretations of the statute or doctrine.

Outputs:

- `adverse_authorities`;
- `distinguishing_notes`;
- `risk_flags`;
- `conflict_summary`.

This is a required phase for motion workflows and optional but recommended for
chat/Word answers.

### Phase 8: Statutory Analysis Pass

When a statute, rule, or code section is central, run a statute-specific path:

1. Fetch the statute text from CA LegInfo.
2. Extract relevant subsections.
3. Search for cases construing the statute.
4. Prefer cases interpreting the exact phrase or subsection.
5. Record amendment/history limitations if current source data cannot support
   them.
6. Emit statute-specific warnings when no annotations or legislative history are
   available.

This path should not pretend to provide Westlaw-style annotated statutes. It
should explicitly say what was and was not checked.

### Phase 9: Synthesis

The synthesizer receives only verified selected authorities, statutes, adverse
authority notes, and the research log summary.

The output format should be structured:

- answer or draft-support memo;
- rule;
- best supporting authority;
- adverse/conflicting authority;
- application or argument use;
- risk/warnings;
- follow-up research questions;
- research basis.

The synthesizer may not cite authority outside the packet.

### Phase 10: Citation Audit

After synthesis or draft generation:

- parse all case, statute, and rule citations;
- verify each citation against the selected packet or source verifier;
- mark off-packet citations as unsupported;
- fail closed for research-backed output if unsupported citations remain.

For motion workflows, the existing citation review panel should continue to be
the user-facing review gate.

## Research Log

Every run should produce a user-readable log:

- original question;
- normalized questions;
- plan;
- sources searched;
- queries run;
- authorities selected and rejected;
- why follow-up searches ran;
- adverse authority found;
- current-law and source warnings;
- citation audit outcome.

The log should be compact enough for UI display but backed by a full JSON
diagnostic object for debugging.

## Prompt Packet

`ResearchPacket` should expose:

```python
packet.to_prompt_block()
packet.to_research_basis_markdown()
packet.to_citation_review_items()
packet.to_authority_pool()
packet.known_case_names()
packet.known_reporter_citations()
```

Prompt block requirements:

- list only verified citable authorities;
- include exact quotes;
- include CourtListener parentheticals as labeled research notes, not as
  verbatim opinion quotations;
- include statutes and relevant text;
- include adverse/conflicting authority;
- include warnings;
- tell the downstream LLM not to cite outside the packet.

Firm-brief authorities marked `unverified_firm` may appear in the research log
and citation review, but not in the verified prompt block unless independently
resolved by local corpus or CourtListener.

## Consumer Integration

### Chat

The existing Chat legal research design remains valid. The Chat-specific service
can initially exist as a consumer adapter, then migrate its retrieval internals
to the shared orchestrator.

Chat should display:

- concise answer;
- Research Basis section;
- source warnings;
- citations with short quotes.

### Word Assistant

The Word Assistant legal research checkbox should call the orchestrator before
the final LLM request. It should inject `packet.to_prompt_block()` and keep the
run log available for review. If research fails closed, it should not produce a
research-backed legal answer.

### Oppose Motion

The existing `research_arguments(...)` path can become a specialized adapter for
motion argument questions. The orchestrator should reuse its verbatim quote and
firm-authority logic, not bypass it.

Opposition drafting should require:

- authority for each substantive argument;
- adverse authority pass;
- citation audit;
- user review gate before final use.

### Generate Motion

Generate Motion should use the same motion adapter as Oppose Motion, with
`side="moving"`. The citation review panel should show source provenance,
alternatives, and adverse authority notes.

## Failure Handling

The orchestrator must fail closed when the user requested research-backed legal
analysis and no verified basis exists.

Fail-closed conditions:

- no selected source is usable;
- only CourtListener is selected but no token is available;
- local corpus is stale and live fallback is required but unavailable;
- no selected authority supports a required proposition;
- selected authority quote is not found in source text;
- final answer contains off-packet citations;
- statute-central question lacks statute text.

Partial success is allowed only when the packet clearly marks unsupported
questions and the downstream answer states that selected sources did not provide
support.

## Caching

Cache at two levels:

- source result cache keyed by source, query, source metadata, and freshness;
- full `ResearchRun` cache keyed by normalized question, source policy, corpus
  signature, firm-index signature, prompt fingerprints, and date.

Do not reuse cached runs when:

- local corpus metadata changed;
- firm-brief index changed;
- prompt templates changed;
- source policy changed;
- freshness policy requires current law and cache is stale.

## Observability

Log:

- run id;
- source policy;
- query count;
- source timings;
- candidate counts;
- selected authority count;
- failed quote verifications;
- off-packet citation count;
- warnings.

Avoid logging privileged matter context beyond what the existing app already
stores. Diagnostic JSON should stay local.

## Testing

Add focused tests under `tests/test_legal_research/test_deep_research/`.

Model and planning tests:

- request normalization;
- source policy normalization;
- plan JSON parsing;
- ambiguous prompt handling;
- noise-term exclusion.

Retrieval tests:

- local-only search;
- firm-only search;
- CourtListener off never calls live client;
- fallback mode calls live only for stale/thin/current-law conditions;
- dedup across source identities;
- provenance preservation.

Quote and ranking tests:

- LLM-selected quote must appear verbatim;
- non-verbatim quote is rejected;
- parentheticals improve candidate ranking when they match the researched
  proposition, but cannot overcome missing direct opinion-text support;
- parenthetical score contribution is capped and duplicate parentheticals are
  dampened;
- parentheticals are included as labeled research notes and are not treated as
  opinion quotes;
- firm authority is preferred only when on point;
- old or weak authority is down-ranked when current law is required.

Conflict tests:

- adverse authority pass runs for motion workflows;
- conflicting authority is included in packet and research basis;
- limiting or distinguishing parentheticals become conflict/treatment signals;
- negative signals become warnings.

Statutory tests:

- statute text fetched for statute-central questions;
- cases construing statute are searched;
- missing statute text fails closed.

Consumer tests:

- Chat packet compatibility;
- Word Assistant prompt injection;
- Oppose Motion authority pool compatibility;
- Generate Motion citation review compatibility.

Verification should stay focused. Avoid broad unrelated legal-research suites
when optional dependencies are missing.

## Phasing

### Phase 1: Shared models and packet contract

Create the deep-research package, models, source policy normalization, research
log, and packet formatting. Add compatibility adapters so Chat and motion flows
can consume the packet without changing retrieval yet.

### Phase 2: Source registry and retrieval unification

Wrap firm authority, local corpus, CourtListener, CA LegInfo, and CA courts in
source adapters. Add candidate deduplication, provenance preservation, and
parenthetical/treatment-signal retrieval from the local corpus when available.

### Phase 3: Iterative planning and refinement

Add plan generation, gap detection, follow-up query generation, and bounded
multi-iteration runs.

### Phase 4: Conflict and statutory passes

Add adverse-authority/conflict detection and statute-centered research flow.

### Phase 5: Consumer migration

Move Chat, Word Assistant, Oppose Motion, and Generate Motion to call the shared
orchestrator through adapters. Keep existing UI surfaces and citation panels.

### Phase 6: Quality evaluation

Build a small gold set of legal research questions with expected authority,
adverse authority, and unacceptable citations. Use it to regression-test search
quality after ranking or prompt changes.

## Acceptance Criteria

- A research run records a visible plan, source searches, selected authority,
  adverse authority, warnings, and citation audit.
- The final prompt packet contains only verified retrieved authorities.
- Non-verbatim LLM-selected quotes are rejected.
- CourtListener parentheticals are used for ranking and research explanation
  but are capped as secondary signals and labeled separately from opinion text
  in prompt packets.
- Off-packet citations are detected and blocked in fail-closed mode.
- Current-law questions trigger live fallback when configured and available.
- Statute-central questions include statute text or fail closed.
- Motion workflows include adverse/conflict authority notes.
- Existing Chat, Oppose Motion, Generate Motion, and Word Assistant behavior can
  adopt the orchestrator incrementally.

## Open Decisions For User Review

1. Should Phase 1 target Chat first, or the motion workflows where citation
   review already exists?
2. Should the first version always run the adverse-authority pass for Chat, or
   only when the user asks for motion/argument strategy?
3. Should unverified firm-brief authorities be visible by default in Chat, or
   hidden unless the user expands the research log?
4. Should current-law mode require live CourtListener for every run, or only
   when the local corpus metadata is stale/thin?
