# Firm Brief Library — Phase 3: Polish & Surfacing (Design)

**Date:** 2026-06-04
**Status:** Approved design (user approved 3.1–3.7; 3.8 excluded), pending spec review.
**Builds on:** firm_briefs Phase 1 (authority), Phase 2 (style), Phase 2.5 (taxonomy) — all merged to main.

## Goal

Make the firm-library feature's value **visible and complete**: surface provenance in the
citation panel ("from your brief" + swap to alternatives), let the user manage the library
from the Workbench, remove the image-only style gap, sharpen authority retrieval, and pay
down small debts. Seven independent pieces (3.1–3.7); 3.8 (3800 build) is out of scope.

## 3.1 — Citation panel "flag both" UI

**Now:** `RetrievedAuthority` carries `source`/`verification`/`source_brief`/`alternatives`
(Phase 1), and the parsed `Citation` already gained a `source_brief` field; but the
`CitationDetailPanel` doesn't render provenance, and the draft's parsed citations aren't
linked to the research pool.

**Change:**
- **Plumbing:** after the draft is parsed into `Citation`s and verified, join each Citation
  to the `RetrievedAuthority` pool by **normalized reporter cite** and copy provenance
  (`source`, `verification`, `source_brief`, `alternatives`) onto the Citation (extend the
  `Citation` model with the missing fields). Persist this with the saved output so reopened
  tasks keep provenance.
- **Panel (`citation_review.py`):** `_citation_header_html`/`_citation_body_html` render a
  **source badge** ("From your brief: *<label>*", `source_brief` basename; click → open that
  PDF) vs "Corpus"; a **verification tier** line (`local`/`courtlistener` = verified;
  `unverified_firm` = amber "⚠ from firm brief — not independently verified"); and an
  **Alternatives** section listing each `alternatives` entry with a **"Use this instead"**
  button that swaps via the existing `_run_find_replacement` machinery (generalized to accept
  a target replacement authority, not just the auto-found one).

**Out of scope:** changing verdict colors for the existing verification path (reuse).

## 3.2 — Workbench "Sample Library" tab

Mirror the existing `StyleExamplesTab`/`MotionTypesTab` pattern (`dialogs.py`
`_refresh_*_tab`, separate `dialogs_*.py` widget files). New `dialogs_sample_library.py`
`SampleLibraryTab` + `_refresh_sample_library_tab()` in `dialogs.py`:
- list `FIRM_BRIEFS_ROOTS` with add/remove (persisted to a small JSON config so it survives
  restarts — `Scripts/prompts/firm_briefs/roots.json`, falling back to `config.FIRM_BRIEFS_ROOTS`);
- a **"Re-index" button** running ingestion in a `QThread` with a progress bar/log;
- **stats** (briefs / citations / per-(type,side) breakdown / last run) from `index.stats()`;
- an **"OCR image-only briefs"** checkbox (runs the fitz→tesseract fill after the native pass).
- Programmatic API for tests (no live Qt event loop needed): `set_roots`, `stats_text`, etc.

## 3.3 — Store full text so image-only briefs model style

**Now:** style excerpts are re-extracted OCR-off at draft time, so image-only briefs yield
no style sample.

**Change:** add a `full_text TEXT` column to `briefs`; ingest stores the extracted text
(including OCR'd text). `style.select_exemplars` reads the stored `full_text` (trimmed to the
argument section) instead of re-extracting; falls back to on-the-fly extraction only when
`full_text` is empty (older rows). Requires a one-time re-index (+ OCR-fill) of the real
library. Index growth is modest (~tens of MB).

## 3.4 — Per-proposition vector rerank for authority

**Now:** `authority_candidates` is FTS5 keyword-only.

**Change:** embed each harvested citation's proposition at ingest into a **second vector
sidecar** (`prop_vectors.f16`), recording `citations.prop_vec_row`. `authority_candidates`
gains an optional semantic rerank: FTS5 keyword pre-filter, then cosine of the argument's
embedded proposition against candidate prop-vectors (reciprocal-rank fusion), when prop
vectors exist; pure-FTS5 fallback otherwise. The research path passes the argument text so
the index can embed the query. ~3,349 propositions embed in minutes at build.

## 3.5 — `--compact` / `--rebuild` CLI flags

`icharlotte_core.firm_briefs.__main__`: `--rebuild` drops the DB + sidecars and re-ingests
from scratch; `--compact` rewrites the vector sidecar(s) keeping only live (`status='ok'`)
rows and re-points `vec_row`/`prop_vec_row`, reclaiming space from upsert/stale churn.

## 3.6 — Shared-module refactor

Extract `_make_firm_provider`, `_make_local_corpus`, `_research_targets`,
`_firm_style_exemplars`, and the `normalize_motion_type` match-time hop into a new
`icharlotte_core/ui/wizard/pages/_motion_research_support.py`. Both `oppose_motion_page` and
`generate_motion_page` import from there (generate stops importing oppose's privates).
Behavior-preserving; existing tests must stay green.

## 3.7 — Harvest cosmetic cleanup

In `citation_harvest`, strip leading signal words ("See ", "See, e.g., ", "In ", "Accord ",
"Cf. ", "See also ") and collapse embedded whitespace/newlines in `case_name` before storing.
Matching is unaffected (keys on reporter cite); this cleans the provenance display (3.1). A
one-off backfill updates existing `citations.case_name` in place.

## Data flow (after)

```
Ingest:  extract (full text, incl OCR) -> store full_text + harvest cites (clean names)
         -> embed profile vec + per-proposition vecs -> index
Draft:   research (RetrievedAuthority pool) -> draft -> parse Citations
         -> JOIN by normalized cite -> copy provenance onto Citations -> verify -> output page
Panel:   per Citation: source badge + verification tier + alternatives (swap via find_replacement)
Style:   select_exemplars reads stored full_text (works for image-only briefs)
Authority: authority_candidates = FTS5 + per-proposition semantic rerank
Manage:  Workbench Sample Library tab -> add roots / re-index (QThread) / stats
```

## Testing

- 3.1: provenance-join function (cite→RetrievedAuthority by normalized cite); panel renders
  badge/tier/alternatives (pytest-qt, `importorskip("PySide6")`, assert `not isHidden()`);
  swap-to-specific-alternative calls find_replacement with the chosen target.
- 3.2: `SampleLibraryTab` programmatic API (roots add/remove persist; stats_text reflects a
  fake index); re-index invokes ingest on a QThread (mock ingest).
- 3.3: ingest stores `full_text`; `select_exemplars` returns an excerpt for an image-only
  brief (no on-the-fly extraction needed); fallback when `full_text` empty.
- 3.4: prop-vector sidecar built; `authority_candidates` semantic rerank orders a
  paraphrase-matching cite above a keyword-only match; FTS5 fallback when no prop vectors.
- 3.5: `--rebuild` rebuilds from scratch; `--compact` shrinks the sidecar and keeps queries
  correct (vec rows realigned).
- 3.6: both pages import from the shared module; full oppose/generate/firm regression green.
- 3.7: harvester strips signal words + newlines; backfill updates existing rows; idempotent.
- Full regression: firm_briefs + opposition + motion_generation + wizard suites green.

## Execution after merge
- Re-index the real library (`--rebuild`) to populate `full_text` + prop vectors + clean names,
  then OCR-fill + re-tag; restart iCharlotte.

## Out of scope
- 3.8 (3800 library build).
- New verdict semantics; new draft engine behavior.
