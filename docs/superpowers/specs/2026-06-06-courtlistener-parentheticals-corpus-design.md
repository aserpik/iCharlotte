# CourtListener Parentheticals Local Corpus Design

**Date:** 2026-06-06
**Status:** Approved design; pending implementation plan
**Related:** `icharlotte_core/legal_research/local_corpus/`, `docs/superpowers/specs/2026-05-29-ca-caselaw-local-corpus-design.md`

## Problem

The local California case-law corpus currently indexes opinion-text passages from
CAP and recent CourtListener bulk opinions. That gives the legal-research stack a
local, rate-limit-free source of primary case text, but it misses a high-signal
kind of case-law evidence: CourtListener parentheticals.

CourtListener parentheticals are descriptions written in one opinion about
another opinion. They are often compact statements of what the described case
stands for. Those summaries are useful for retrieval because they connect legal
issues and propositions to the case that other courts thought was worth citing.
They should improve search recall and authority selection, but they must not be
confused with the described case's own text.

CourtListener exposes parentheticals as bulk data, not through the REST API. The
bulk snapshot is a full dataset, so refreshes should be idempotent and should
avoid re-streaming the large opinions file when a reusable opinion-id map already
exists.

Source references:

- CourtListener bulk data docs: https://www.courtlistener.com/help/api/bulk-data/
- CourtListener case law API docs: https://www.courtlistener.com/help/api/rest/case-law/
- CourtListener bulk export fields: https://github.com/freelawproject/courtlistener/blob/main/scripts/make_bulk_data.sh

## Goal

Bulk-ingest CourtListener parentheticals and attach them as high-signal passages
to cases already present in the local California corpus. Retrieval should be able
to find a California case through the way later opinions describe it, while
verification and drafting retain a clear distinction between:

- primary text from the described case, and
- secondary parenthetical text written by a citing/describing opinion.

## Non-Goals

- Do not expand the corpus beyond cases already present in the local California
  corpus.
- Do not append parentheticals to `cases.full_text`.
- Do not treat parentheticals as quote support from the described opinion.
- Do not add Shepard's, KeyCite, or true negative-treatment analysis.
- Do not add a live CourtListener REST fallback for parentheticals.
- Do not download or permanently store the full CourtListener opinions bulk file.

## Key Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Attachment scope | Attach only to cases already in the local CA corpus | Keeps research CA-focused while improving recall for existing authorities |
| Storage model | Store as tagged `passages` rows plus provenance metadata | Reuses current FTS/vector retrieval while preserving source distinction |
| Primary-text boundary | Keep `cases.full_text` unchanged | Prevents parenthetical text from polluting quote verification |
| Mapping strategy | Map `described_opinion_id -> cluster_id`, then attach to `cl:<cluster_id>` or a CAP case matched by citation | Parentheticals point to opinions, while the corpus addresses cases/clusters |
| Refresh strategy | Cache opinion-id mapping metadata and ingest full parenthetical snapshots idempotently | Avoids unnecessary 54 GB opinion-stream passes on routine parenthetical refresh |
| Ranking | Give parenthetical passages a modest retrieval boost, not proof status | They are high-signal retrieval evidence, not the case's own holding text |

## Architecture

The feature extends the existing `local_corpus` package instead of adding a
separate research index.

```text
CourtListener opinions bulk
    -> opinion id map cache
       opinion_id -> cluster_id

CourtListener citations / clusters bulk
    -> cluster/citation map
       cluster_id -> CA citation metadata

CourtListener parentheticals bulk
    -> parenthetical_loader
       filters described_opinion_id to local corpus cases
       normalizes text and score
       emits tagged PassageRecord rows

CorpusIndexer
    -> writes passages, FTS rows, vectors
       passage_type = parenthetical
       provenance columns populated

LocalCaseCorpus.search_opinions()
    -> BM25 + semantic over opinion and parenthetical passages
       modest parenthetical boost
       CaseResult remains case-level

Verification and drafting
    -> may use parentheticals for retrieval context
       must use full_text/opinion passages for quoted case support
```

## Data Model

Additive schema changes keep existing corpora readable:

- `passages.passage_type TEXT DEFAULT 'opinion'`
- `passages.source TEXT DEFAULT ''`
- `passages.parenthetical_id TEXT DEFAULT ''`
- `passages.parenthetical_score REAL`
- `passages.described_opinion_id TEXT DEFAULT ''`
- `passages.describing_opinion_id TEXT DEFAULT ''`
- `passages.describing_cluster_id TEXT DEFAULT ''`

The existing `PassageRecord` gains matching optional fields. Existing callers can
continue constructing opinion passages with only `passage_uid`, `case_uid`,
`ordinal`, `text`, `page_label`, and `vec_row`.

Add a small cache table:

- `courtlistener_opinion_map`
  - `opinion_id TEXT PRIMARY KEY`
  - `cluster_id TEXT NOT NULL`
  - `snapshot_date TEXT NOT NULL`

Add ingest metadata:

- `corpus_meta.parentheticals_snapshot_date`
- `corpus_meta.parentheticals_count`
- `corpus_meta.parentheticals_min_score`
- `corpus_meta.parentheticals_max_per_case`

## Ingest Flow

1. Open the existing corpus DB and vector sidecar.
2. Ensure additive schema columns and cache tables exist.
3. Build or refresh the opinion-id map:
   - Prefer existing `courtlistener_opinion_map` rows for the requested snapshot.
   - If missing, stream the CourtListener opinions bulk file and store only
     `id -> cluster_id`.
   - Do not store opinion text during this pass.
4. Build cluster-to-local-case attachment:
   - Direct hit: if the local corpus has `case_uid = cl:<cluster_id>`, attach
     there.
   - CAP overlap hit: use CourtListener citations metadata to map the cluster's
     preferred/parallel CA citation to an existing CAP case by normalized
     citation.
   - No local case hit: skip the parenthetical.
5. Stream the CourtListener parentheticals bulk file.
6. For each parenthetical:
   - Require `score >= min_score`.
   - Resolve `described_opinion_id` through the opinion-id map.
   - Resolve that cluster to a local case.
   - Normalize text.
   - Deduplicate by `parenthetical_id`.
   - Respect `max_per_case` by score order when requested.
7. Append parentheticals through `CorpusIndexer` so FTS rows and vector rows stay
   aligned.
8. Store corpus metadata and checkpoint the snapshot.

Default parameters:

- `min_score = 0.5`
- `max_per_case = 25`
- `embed = false` by default for a fast keyword-only append, with an explicit
  `--embed-parentheticals` option to generate semantic vectors during ingest.

If `embed = false`, parenthetical passages still get vector rows with zero
placeholder vectors, matching the current CL append pattern. A later re-embed
operation can fill those vectors without changing schema.

## Search Behavior

Parenthetical passages participate in normal FTS search and optional semantic
ranking. `LocalCaseCorpus` groups all matching passages back to case-level
results exactly as today.

Ranking changes:

- Opinion-text passages remain the default authority source.
- Parenthetical passages receive a modest score boost when they match the query,
  because they are curated descriptions from citing opinions.
- The `CaseResult.snippet` should prefer the best matching passage, but display
  metadata should indicate when the snippet is a parenthetical.

The research prompt can use parenthetical snippets as "described by later cases"
context, but final quoted support must come from the described case's own text
unless a user specifically asks to quote the citing case.

## Verification Boundary

Parentheticals improve retrieval and triage. They do not prove that the
described case contains the parenthetical wording.

Verification rules:

- `get_opinion_text(case_uid)` continues returning only the described case's
  full text.
- Local citation verification compares propositions against `cases.full_text`
  and opinion passages.
- Parenthetical snippets may be shown as research context with provenance.
- A citation cannot be marked supported merely because a parenthetical matched.

This preserves the existing fail-closed legal-writing behavior while allowing
parentheticals to guide the system toward better cases.

## CLI and Operations

Extend the local corpus build CLI:

```powershell
python -m icharlotte_core.legal_research.local_corpus.build --source parentheticals --cl-date 2026-03-31
```

Supported options:

- `--parentheticals-min-score FLOAT`
- `--parentheticals-max-per-case INT`
- `--embed-parentheticals`
- `--refresh-opinion-map`

Operational notes:

- Latest visible snapshot checked during design: `2026-03-31`.
- The `parentheticals-2026-03-31.csv.bz2` object is about 286 MB compressed.
- The matching `opinions-2026-03-31.csv.bz2` object is about 54 GB compressed,
  so the map-cache path matters.
- Routine refresh after the opinion map exists should need only the
  parentheticals snapshot plus lightweight citation/cluster metadata.

## Testing

Add focused tests under `tests/test_legal_research/test_local_corpus/`:

- Schema migration creates parenthetical columns and opinion-map cache without
  rewriting existing cases.
- Loader parses parenthetical rows and filters by score.
- Loader maps `described_opinion_id` through `opinion_id -> cluster_id`.
- Loader attaches to direct `cl:<cluster_id>` cases.
- Loader attaches to CAP cases by normalized citation.
- Loader skips parentheticals for cases absent from the local CA corpus.
- Ingest is idempotent when re-run with the same parenthetical id.
- Search returns a case when only its parenthetical passage matches the query.
- `get_opinion_text()` does not include parenthetical text.
- Verification ignores parenthetical text as direct quoted support.
- Metadata records snapshot date, count, min score, and max-per-case settings.

Focused verification commands:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus -q
python -m py_compile icharlotte_core/legal_research/local_corpus/*.py icharlotte_core/legal_research/local_corpus/loaders/*.py
```

## Risks and Mitigations

- The opinion-id map may require one large opinions stream. Mitigation: cache it
  by snapshot and reuse it for parenthetical-only refreshes.
- Parentheticals can overstate, simplify, or frame a case for a citing court's
  purpose. Mitigation: use them for retrieval context only, not direct support.
- Low-score parentheticals can add noise. Mitigation: default score threshold
  and per-case cap.
- Parenthetical volume can inflate vector storage. Mitigation: keyword-only
  append by default, optional embedding, and `max_per_case`.
- CourtListener bulk field names may drift. Mitigation: tests use fixture CSVs
  with the current exported fields: `id`, `text`, `score`,
  `described_opinion_id`, `describing_opinion_id`, `group_id`.

## Acceptance Criteria

- A user can run a parenthetical ingest against an existing local corpus without
  rebuilding CAP or CL case text.
- Parentheticals attach only to local corpus cases.
- Search can discover a case through a matching parenthetical.
- Parenthetical provenance is stored and available for display/debugging.
- Opinion text and citation verification remain free of parenthetical pollution.
- Re-running the same snapshot does not duplicate parenthetical passages.
