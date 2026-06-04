# Firm Brief Library — Authority Reuse & Style Selection (Design)

**Date:** 2026-06-04
**Status:** Approved design, pending spec review
**Tasks affected:** `oppose_motion`, `generate_motion` (wizard mode)

## Goal

Use the firm's sorted library of past briefs (the `5800_AMTRUST_Pleadings_PDFs`
folder tree, and future libraries like 3800) as a dual-purpose resource for the
**Generate a Motion** and **Oppose a Motion** wizard tasks:

1. **Style / style-guide samples** — feed the drafter the most relevant past briefs
   of the same motion type & side as voice/format exemplars.
2. **Authority reuse** — harvest the case law cited in those past briefs and reuse
   it (where appropriate) *in addition to* the existing local-corpus research, so the
   drafter prefers the firm's previously-relied-upon authorities for a given point.

The feature is **purely additive**: when the firm-brief index is absent or errors,
both tasks fall back to today's corpus-only behavior unchanged.

## Decisions (from brainstorming)

- **Authority weighting:** *Prefer firm, but flag both.* The drafter leads with the
  firm's previously-cited authority for each point; corpus alternatives for the same
  point are surfaced in the citation panel for one-click swap.
- **Unverifiable firm cites:** *Try live CourtListener, then flag.* Resolve order is
  local corpus → live CourtListener verifier → if neither confirms, include the cite
  flagged "⚠ from firm brief — not independently verified," using the passage as it
  appeared in the firm brief.
- **Style selection:** *Auto — most similar to the current matter*, by embedding a
  distilled issue-profile of each sample and cosine-matching to the current motion.
- **Updatability (explicit requirement):** the index must support incremental
  re-indexing over time — drop new PDFs in, re-run, only new/changed files process.

## Architecture

New shared package `icharlotte_core/firm_briefs/`, consumed by both tasks, plus thin
integration shims at existing seams. Nothing in the current
research → rerank → verify → draft pipeline is replaced; we add a second (preferred)
candidate source and a smarter style selector.

```
            ┌──────────────── OFFLINE (incremental) ────────────────┐
 sorted     │ ingest.py: extract text (OCR-aware) → harvest          │
 library  → │  citations+propositions → extract argument headings →  │
 roots      │  compose issue profile → embed → UPSERT into index     │
            │  (unchanged files skipped via content hash)            │
            └───────────────────────────┬───────────────────────────┘
                                         │
              FirmBriefIndex  (SQLite firm_briefs.db + profiles.f16 sidecar)
                                         │
        ┌────────────────────────────────┴───────────────────────────┐
   STYLE consumer                                            AUTHORITY consumer
   style.select_exemplars(type, side,              FirmAuthorityProvider.candidates_for(
     motion_metadata, k=3)                            proposition, type, side)
   embed current motion profile →                  proposition hybrid-search →
   cosine vs same type+side →                       resolve each firm cite to opinion
   top-3 trimmed excerpts →                         text (local corpus → CL fallback) →
   drafter style feed                               inject as PREFERRED candidates into
                                                    research_argument() rerank+verify
```

**Folder → metadata reuse:** the folder each PDF was sorted into maps directly to
`(motion_type, side)` — `Oppositions/Motion to Compel` → `(compel, opposition)`,
`Motion - Summary Judgment` → `(msj, moving)`, `Replies/Demurrer` → `(demurrer, reply)`,
`Pleadings - Complaint` → `(complaint, pleading)`. Ingestion derives metadata from the
path; no manual tagging. The index is firm-wide (client-agnostic); each library is a
configured root.

### Components (each independently testable)

- `ingest.py` — walks configured root(s); per file: extract → harvest → profile →
  embed → upsert; incremental + resumable.
- `index.py` — `FirmBriefIndex`: SQLite + vector sidecar, **thread-local connections**
  (mandatory — the local-corpus pipeline fans out over a ThreadPoolExecutor and a
  single shared sqlite connection raises cross-thread errors that get swallowed into
  empty results). Query methods for style and authority.
- `citation_harvest.py` — reuses `icharlotte_core/opposition/citation_parser.py` to
  pull `Citation` records (case/statute, with the sentence-window proposition and the
  passage the brief quoted) from sample text.
- `profile.py` — composes the issue-profile string (relief + argument headings +
  propositions); optional cheap LLM distill for garbled OCR headings.
- `embedding.py` — thin reuse of the corpus's fastembed BGE-small embedder (384-dim,
  no torch). No new dependency.
- Integration shims: `style.select_exemplars(...)` (feeds the existing exemplar
  channel) and `FirmAuthorityProvider` (feeds `research_argument`).

## Index schema & incremental ingestion

Storage under `FIRM_BRIEFS_DATA_DIR` (default `.gemini/firm_briefs/`, gitignored):
`firm_briefs.db` (SQLite) + `profiles.f16` (np.memmap vector sidecar, row-aligned).

```sql
briefs(
  id INTEGER PK,
  path TEXT UNIQUE, content_hash TEXT,      -- sha1(path|mtime|size)
  motion_type TEXT, side TEXT,              -- from folder
  heading TEXT, profile TEXT,               -- embedded issue-profile string (kept for debug/LLM)
  vec_row INTEGER,                          -- row in profiles.f16 (-1 = not embedded)
  char_len INTEGER, ocr_ratio REAL,         -- style-quality signals
  ingested_at TEXT, status TEXT             -- ok | text_failed | stale
)
citations(
  id INTEGER PK, brief_id INTEGER REFERENCES briefs(id),
  case_name TEXT, reporter_cite TEXT, year TEXT,
  norm_cite TEXT,                           -- normalized "vol reporter page" (indexed)
  proposition TEXT,                         -- what the cite is offered for
  quoted_passage TEXT,                      -- what the brief quoted (CL-fallback/unverified passage)
  prop_vec_row INTEGER                      -- optional per-proposition embedding
)
ingest_runs(...)                            -- run log: added/updated/skipped/failed
```
- Index on `citations.norm_cite` (avoid the corpus's linear-scan lookup). FTS5 over
  `citations.proposition` for keyword+vector authority lookup.

Incremental ingestion (`ingest.py`), modeled on the corpus's resumable build:
1. Walk root(s); compute `content_hash` per PDF.
2. Skip if `path` present with same hash and `status='ok'`.
3. New/changed → extract → harvest → profile → embed → **upsert** (replace brief +
   its citations; reuse `vec_row` if present else append to sidecar).
4. Removed file (in DB, gone on disk) → `status='stale'` (excluded from queries; rows
   kept so vec alignment holds; `--compact` rewrites the sidecar to reclaim space).
5. Crash-safety: embed → fsync sidecar → commit brief row (sidecar-before-DB), so a
   kill mid-run just reprocesses the in-flight file.

Embedding throughput: profiles are short strings (~hundreds of chars) → ~800/sec on
this CPU-only box; ~10k samples is a one-time minutes job, incremental near-instant.

## Draft-time authority merge (core)

Plugs into the existing `research_argument()` per-proposition flow. Per
argument/section-leaf proposition, candidates now come from two sources through the
*same* rerank + verbatim-verify gate.

`FirmAuthorityProvider.candidates_for(proposition, motion_type, side)`:
1. Search over `citations.proposition`, filtered to the motion type. **Baseline is
   FTS5 keyword** matching of the argument proposition against the stored
   propositions; **per-proposition vector rerank is an optional enhancement** used
   only when `prop_vec_row` embeddings have been built (Phase 1 ships FTS5-only; the
   brief-level profile vector already exists for style). **Authority matching ignores
   side** (a meet-and-confer case is good law no matter who cited it); side strictly
   filters only **style**.
2. Resolve each firm cite to opinion text for the verbatim gate:
   - `corpus.lookup_by_citation(norm_cite)` → opinion text → candidate `source='firm'`,
     `verification='local'`; runs the normal LLM rerank + verbatim-passage check.
   - Not in corpus → live CourtListener `case_verifier` on the proposition. Verified →
     `verification='courtlistener'`, passage = brief's quoted passage. Neither confirms
     → include flagged `verification='unverified_firm'`, passage = brief's quoted
     passage (skips the corpus-text gate by design).
3. **Merge & prefer:** firm candidates ordered first and tagged in the rerank prompt
   ("prefer a firm-tagged candidate when on-point"). Corpus candidates the rerank also
   liked for the same proposition are attached to the chosen authority as
   `alternatives` ("flag both").

`RetrievedAuthority` (in `icharlotte_core/opposition/models.py`) gains: `source`
(`firm`|`corpus`), `verification` (`local`|`courtlistener`|`unverified_firm`),
`source_brief` (provenance — which firm brief), `alternatives: list[RetrievedAuthority]`.
Drafter pool leads with firm authorities; gap markers only when *neither* source has a
case (unchanged behavior).

Both tasks share the provider: oppose_motion `side='opposition'`, generate_motion
`side='moving'`. Provider wrapped so an index miss / CL timeout degrades gracefully to
corpus-only (never blocks a draft). CL-fallback lookups are bounded and reuse the
existing 1-req/s CL throttle + cache.

## Style selection & citation-panel UI

`style.select_exemplars(motion_type, side, motion_metadata, k=3)`:
- Compose current motion issue-profile from `MotionMetadata` (relief + principal
  arguments), embed (same model), cosine vs `briefs` where `motion_type` matches,
  **`side` matches strictly**, `status='ok'`.
- Quality guard: down-rank low `char_len` / high `ocr_ratio` so a garbled scan never
  wins as the style model; near-duplicate collapse (`_1`/`_2` copies, versions).
- Returns trimmed **argument-section** excerpts (token budget), cached by content hash
  like today's exemplar cache.
- Feeds the existing drafter exemplar channel; manually-pinned Workbench exemplars
  still merge and win (backward-compat).

Citation panel — extends the existing `CitationDetailPanel` (right panel, no popup):
- Source badge: "From your brief: *<label>*" (click → open source PDF) vs "Corpus."
- Verification tier drives underline color (reuse current scheme): `local` &
  `courtlistener` → verified; `unverified_firm` → amber "⚠ from firm brief — not
  independently verified."
- Alternatives section: corpus options for the same point, each with "Use this
  instead" → reuses existing `find_replacement` swap machinery.

## Configuration, build/refresh, management

- `config.py`: `FIRM_BRIEFS_DATA_DIR` (default `.gemini/firm_briefs/`),
  `FIRM_BRIEFS_ROOTS` (list; seeded with `5800_AMTRUST_Pleadings_PDFs`). Optional LLM
  distill uses a cheap model (Flash); ingestion otherwise needs no API.
- CLI: `python -m icharlotte_core.firm_briefs.ingest --root <path> [--rebuild] [--compact]`
  — incremental by default, backgroundable.
- Workbench **"Sample Library" tab:** configured roots (add/remove), **"Re-index"**
  button (ingest in a QThread w/ progress), live stats, and existing Style-Example
  pins. This is the "update over time" surface.

## Testing (TDD, `.venv` python, pytest)

- ingest: folder→`(type,side)` mapping; hash-skip; upsert on change; `stale` on
  removal; sidecar-before-DB crash ordering.
- index: **thread-local connections under a ThreadPoolExecutor**; style filters by
  side; authority ignores side; `norm_cite` lookup.
- provider: corpus-resolve; CL-fallback (mock verifier); `unverified_firm` flag;
  graceful degrade on miss/timeout.
- research integration: firm candidates lead; `alternatives` attached; gap-marker only
  when both empty.
- style: similarity ranking; side filter; OCR/dup down-rank; cache.
- UI: panel source badge / verification tier / alternatives-swap (pytest-qt,
  `importorskip("PySide6")`, `assert not isHidden()`).
- regression: existing oppose/generate tests stay green; **corpus-only path unchanged
  when the index is absent**.

## Phasing (each independently shippable)

1. **Ingestion + index + authority merge** — the high-value core; works headlessly.
2. **Style auto-selection.**
3. **Citation-panel source/alternatives/swap UI + Workbench Sample Library tab** —
   makes "flag both" interactive.

## Non-goals / known limits

- No true Shepard's/good-law beyond the corpus's soft inbound-citation signal and the
  CL verifier's existence check.
- OCR'd scans yield noisier citation harvests; the verbatim/CL gates catch most bad
  cites, and `unverified_firm` flags the rest for human review.
- Statutes/rules cited in firm briefs: harvested and reusable, but verification stays
  on the existing statute path (leginfo), out of scope for the case-law merge here.
