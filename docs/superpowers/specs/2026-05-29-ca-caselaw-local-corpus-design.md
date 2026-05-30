# CA Case Law Local Corpus — Design

**Date:** 2026-05-29
**Status:** Approved (design); pending implementation plan
**Related:** [[2026-05-27-oppose-motion-citation-grounding-design]], `oppose_motion_redesign` memory, `courtlistener_semantic_search` memory

## Problem

The Oppose-a-Motion wizard grounds its draft in real CA authority by calling the
CourtListener (CL) live API for retrieval **and** citation verification. CL throttles
standard tokens to **5 requests/minute** on both the `/search/` and
`/clusters/`+`/opinions/` endpoints. The research step fires a parallel burst that
stampedes this limit; even with request retry + a 1 s client-side throttle (already
shipped), a 5-argument brief needs ~100 calls and cannot complete in reasonable time —
it hangs 10+ minutes and returns zero authorities, producing an opposition with no case
cites. A different free-tier token does not raise the limit (confirmed: it is an
account-tier throttle, not per-token). Raising it requires Free Law Project granting an
elevated throttle to the account.

## Goal

Replace the live-API dependency with a **local, source-agnostic CA case-law corpus** that
serves retrieval **and** verification entirely offline — unlimited, fast, deterministic,
no rate limit anywhere in the pipeline. Full coverage (≈1850–present) assembled **from
bulk data only** (no live API): Harvard Caselaw Access Project (CAP) for the
≤~2020 backbone, CourtListener bulk for the 2020+ gap.

## Non-Goals

- True good-law / negative-treatment verification (Shepard's/KeyCite). Neither source has
  it; we provide only a **soft** staleness signal (inbound citation count + latest citing
  year). The manual attorney good-law check remains, exactly as today.
- Federal, 9th-Circuit, sister-state, or statutory authority. CA case law only. (Statutes
  continue to come from leginfo; federal reporters are a possible future extension.)
- A live-API fallback bridge. v1 is bulk-only by deliberate decision.

## Key Decisions (locked during brainstorming)

| Decision | Choice | Rationale |
|---|---|---|
| Retrieval engine | FTS5 (BM25) **+** local semantic rerank, both in v1 | Match the hosted semantic engine's vocab-mismatch recall |
| Embedding deps | Lightweight ONNX (`fastembed`, BGE-small), **behind a swappable `Embedder` interface** | No PyTorch/CUDA; avoids DLL-conflict risk next to Qt/PyMuPDF; upgradeable later |
| Coverage | CAP (≤~2020) **+** CL bulk-filtered (2020+), merged, **no phasing** | Full coverage in v1, all from bulk → zero live-API dependency |
| Vector store | Exact brute-force cosine over a **`np.memmap` float16 sidecar**; no ANN | At ~650k passages, exact recall is cheap (~200–300 ms); ANN buys unneeded scale at the cost of a dependency + recall. Isolated scorer → sqlite-vec/hnswlib is a drop-in upgrade if federal expands the corpus |
| Good-law signal | Soft only (inbound citation count + latest citing year), **explicit plan item** | No clean overruled flag exists in either source |
| Pin-cites | Parse CAP HTML `page-label` anchors → per-passage `page_label`, **explicit plan item** | JSON plain text lacks page anchors; briefs need pin-cites |
| Verification | Move **local** too (`LocalCaseVerifier`) | Leaving verification on the live API would just relocate the 5/min hang |

> Disk note: the machine now has >25 GB free, so disk is no longer tight. The lightweight
> dependency choices stand regardless — the primary reason to avoid the torch stack is
> Windows DLL-conflict risk alongside Qt/PyMuPDF, not disk.

## Architecture

Source-agnostic local corpus. Bulk loaders normalize CAP and CL records into one schema;
a single `LocalCaseCorpus` owns search + rerank and exposes the **same interface** as
today's `CourtListenerClient`, making it a drop-in for `argument_research._hybrid_search`.
Adding a source later = one more loader; nothing else changes.

```
            ┌──────────── BUILD TIME (one-time / quarterly) ────────────┐
 CAP ZIPs  ─► cap_loader ────────┐
 (static.case.law)               ├─► normalize ─► SQLite (FTS5 + metadata)
 CL bulk CSV ─► cl_bulk_loader ──┘                       + vectors.f16 (memmap)
 (stream-filtered: CA courts, date > CAP)        ▲
                                        onnx_embedder (fastembed; swappable)
                                        good-law signal roll-up (cites_to / citation-map)

            └─────────────── QUERY TIME (per argument) ────────────────┘
 argument ─► LocalCaseCorpus.search_opinions()
                 │ 1. FTS5 BM25  ─────────────► candidate cases
                 │ 2. embed query, exact cosine over memmap ─► candidate cases
                 │ 3. Reciprocal Rank Fusion
                 └► CaseResult[]  (drop-in for existing _hybrid_search)
            LocalCaseCorpus.get_opinion_text() ─► full local text (no API)
            LocalCaseVerifier.verify_all()     ─► SUPPORTED/NOT_FOUND (no API)
```

### New module: `icharlotte_core/legal_research/local_corpus/`

| File | Responsibility |
|---|---|
| `schema.py` | SQLite DDL + connection helper |
| `models.py` | Normalized `CaseRecord` / `PassageRecord` |
| `loaders/cap_loader.py` | CAP ZIP → normalized records (incl. HTML page-label parse) |
| `loaders/cl_bulk_loader.py` | CL bulk CSV stream-filter → normalized records |
| `embedder.py` | `Embedder` protocol + `OnnxEmbedder` (fastembed) |
| `indexer.py` | Build FTS5; compute + write passage vectors to `vectors.f16` |
| `authority_signals.py` | Build inbound citation counts / latest-citing-year |
| `corpus.py` | `LocalCaseCorpus`: `search_opinions` / `get_opinion_text` / `get_authority_signals` |
| `verifier.py` | `LocalCaseVerifier`: pool-aware, local-text verification |
| `build.py` | CLI orchestrator (`--source cap|cl|all --data-dir PATH`) |
| `README.md` | Build/refresh steps, embedder-swap how-to, known limits |

## Data Model

**`CaseRecord`:** `case_uid` (source-prefixed string, e.g. `cap:269732`, `cl:4408734`),
`source`, `name`, `name_abbreviation`, `citation` (preferred CA reporter cite),
`parallel_citations`, `court`, `decision_date`, `year`, `docket_number`, `url`,
`full_text`, `citation_count`, `latest_citing_year`, `cites_to` (outbound edges).

**`PassageRecord`:** `passage_uid`, `case_uid`, `ordinal`, `text`, `page_label`
(pin-cite anchor from CAP HTML), `vec_row` (row index into `vectors.f16`).

### SQLite schema (single DB file + sidecar vector file)

- `cases` — metadata + `full_text`; index on `citation` (dedup + cite lookup).
- `passages` — passage rows + `page_label` + `vec_row`.
- `passages_fts` — FTS5 virtual table over passage `text` (BM25 at passage granularity,
  grouped back to cases).
- `citation_edges` — outbound edges; inbound counts rolled up onto `cases` at build end.
- `vectors.f16` — sidecar contiguous float16 memmap; row `vec_row` = passage embedding.

## Ingest Pipeline

### `cap_loader` (backbone, ≤~2020)
Enumerate CA reporter volumes (`cal`, `cal-2d/3d/4th/5th`, `cal-app`, `cal-app-2d/3d/4th/5th`,
`cal-rptr-3d`, `cal-unrep`) → download each `https://static.case.law/{rep}/{vol}.zip`
(parallel, **idempotent/resumable** — skip already-ingested) → from `json/*.json`: metadata
+ opinion texts (by type: majority/concurring/dissenting) + `cites_to` → from paired
`html/*.html`: **parse `class="page-label">*NN` anchors into a char-offset→reporter-page
map** so each passage carries a `page_label` → normalize (fix `§` encoding, whitespace,
footnote markers) → chunk into ~512-token passages on paragraph boundaries, assign
`page_label` from the offset map → dedup by citation.

### `cl_bulk_loader` (2020+ gap, stream-filtered)
Download `courts` CSV (tiny) → CA court ids. Stream `opinion-clusters` CSV (~2.3 GB) →
build the set of cluster_ids that are CA **and** dated after the CAP cutoff (with a small
overlap buffer for dedup), retaining cluster metadata (citation, court, date, name). Stream
the ~50 GB `opinions` CSV decompressed → keep only rows whose `cluster_id` is in that set
(reusing the existing multi-field text-priority logic from `courtlistener.py`), **discarding
all other rows so the 50 GB never lands on disk**. Cross-source dedup by normalized official
citation, **CAP authoritative on overlap** (cleaner canonical text). CL pagination is
unreliable → recent-slice cites may carry coarser pin-cites (known limit). Build is
idempotent (re-running skips already-written clusters); a mid-stream failure re-streams
(true byte-offset resume is infeasible over bz2 — accepted limitation).

### Good-law soft-signal builder (explicit item)
After both sources ingest: invert CAP `cites_to` + fold in CL `citation-map` → per-case
inbound `citation_count` and `latest_citing_year`, stored on `cases`, surfaced via
`get_authority_signals()` exactly like today's CL method so the drafter's existing good-law
hint works unchanged. **Soft staleness signal only — not Shepard's.**

## Query & Rerank Flow

`LocalCaseCorpus` mirrors `CourtListenerClient` (drop-in; no logic change to
`argument_research.py`):

- **`search_opinions(query, *, semantic, max_results, published_only)`** — true hybrid:
  1. BM25 over `passages_fts` → candidate cases.
  2. Embed query (fastembed), exact cosine over memmap'd vectors → candidate cases
     (runs independently of BM25 → catches vocab-mismatch cases keyword misses).
  3. Fuse by `case_uid` via Reciprocal Rank Fusion → top `max_results` as `CaseResult`
     (`snippet` = best-matching passage, `cluster_id` = `case_uid`).
  - `_hybrid_search` calls this with `semantic=True` then `False`; the `False` call is the
    BM25-only view. Works unchanged.
- **`get_opinion_text(case_uid)`** → full local text. No API.
- **`get_authority_signals(case_uid)`** → `{citation_count, latest_citing_year}`. No API.

**Integration change (explicit):** `argument_research._opinion_text` currently calls
`cl_client.get_opinion_text(int(cluster_id))`. Make it **id-type-agnostic** (pass the id
through as-is; the CL client keeps its own internal `int()`), so a string `case_uid` works.
One-line change; both clients keep working.

## Verification Path

Today `build_opposition_verifier` verifies every citation against CL at 5/min. Leaving it
live would relocate the hang to the verification step, so verification moves local:

- **`LocalCaseVerifier`** mirrors the verifier's `verify_all(...) -> CitationVerification[]`:
  - **In-pool cites** (from `LocalCaseCorpus`, already passed the verbatim-passage check in
    `select_authorities`) → **SUPPORTED by construction**, carrying the stored passage +
    `page_label` pin-cite. Zero API, zero LLM.
  - **Proposition-level / ambiguous checks** → LLM compares the proposition against the
    case's local full text. No API.
  - **Off-pool cites** → flagged **NOT_FOUND**, exactly as today.
- Net: research + drafting + verification are **all API-free and unlimited**.
- Wiring: `oppose_motion_page.py` selects `LocalCaseVerifier` instead of
  `build_opposition_verifier`. The CL verifier stays in the tree (untouched) for other callers.

## Build, Dependencies & Operations

- **`build.py` CLI:** `python -m icharlotte_core.legal_research.local_corpus.build
  --source {cap|cl|all} --data-dir PATH`. Stages: download/stream → normalize → chunk →
  embed → index → roll up good-law signals. Idempotent/restartable; logs progress + a final
  coverage summary (case counts by reporter + year range). Free-disk preflight warning.
- **New deps** (`requirements.txt`): `fastembed` (pulls `onnxruntime` + `tokenizers`,
  ~330 MB, no torch/CUDA) and `numpy`. No faiss, no torch.
- **Data location:** new `CASELAW_DATA_DIR` in `config.py`, default
  `C:\geminiterminal2\.gemini\caselaw\` (add to `.gitignore`), **relocatable to a roomier
  drive**. Holds the SQLite DB + `vectors.f16`. CAP ZIPs download to a scratch subdir and
  are deleted after ingest; the CL 50 GB stream never lands.
- **Refresh:** quarterly manual re-run, documented in module README (CAP frozen → re-fetches
  nothing; CL re-streams the newest snapshot). Scheduled-task wrapper = future nicety.
- **Embedder swap:** `Embedder` protocol; `OnnxEmbedder` default. README documents dropping
  in a stronger model later (re-embed only; schema/index/search untouched).

## Testing

- **Loaders:** synthetic fixture ZIP (2–3 fake CAP cases, paired JSON+HTML) + small CL
  fixture CSVs → assert normalization, **HTML page-label → pin-cite offset mapping**,
  chunking, `§`/whitespace cleanup, cross-source dedup (CAP wins overlap).
- **Search/rerank:** temp DB seeded with a few cases → BM25 hits, semantic hits (via a
  **deterministic fake embedder** so CI needs no model), RRF ordering, `CaseResult` shape.
- **Good-law signal:** inbound count / latest-citing-year roll-up from a small `cites_to` graph.
- **Drop-in compatibility:** run `research_argument` against a fake `LocalCaseCorpus`; assert
  no logic change needed in `argument_research.py`; assert the `_opinion_text` id-agnostic tweak.
- **Local verification:** in-pool → SUPPORTED with passage+pin-cite; off-pool → NOT_FOUND.
- **One gated integration test** exercises real `fastembed` on a 2-case corpus (skipped if
  the model isn't present), mirroring the existing gated-test pattern.
- All via venv python; under `tests/test_legal_research/` + `tests/test_opposition/`.

## Risks & Known Limits (carried forward, not forgotten)

1. **No true good-law check** — soft signal only; manual attorney step remains (no regression
   vs. live API).
2. **Recency gap** — CAP frozen ~2020; CL bulk is quarterly (≈2 months stale). Freshest 0–3
   months absent; currency depends on quarterly rebuilds.
3. **CA-only** — no federal/sister-state/statutes in the pool.
4. **Semantic quality < hosted** — BGE-small < CL's fine-tuned legal model; mitigated over
   time via the swappable embedder.
5. **Recent slice = lowest-quality text** — CL-bulk 2020+ text is multi-source (OCR/HTML/
   syllabus-only) vs. CAP's canonical text.
6. **One-time ~52 GB network stream** for the CL slice; mid-stream failure re-streams.
7. **Quarterly rebuild + multi-step pipeline** — ongoing maintenance the user owns.
8. **Implementation cost** — a real subsystem (ETL, embeddings, vector store, two loaders,
   local verifier); larger build than the shipped rate-limit fix.
