# Firm Brief Library — Phase 2.5: Motion-Type Taxonomy & Normalizer (Design)

**Date:** 2026-06-04
**Status:** Approved design, pending spec review
**Builds on:** firm_briefs Phase 1 (authority) + Phase 2 (style), both merged to main.

## Problem

Firm sample/authority matching is keyed on an **exact** `motion_type` id, but the three
places that produce/consume that id don't agree:

- **Ingest** derives the id from the *folder* name (path_meta `_TYPE_ALIASES`). The
  `Motions - Other` folder (and `Oppositions|Replies / Other` subfolders) are tagged
  `other` wholesale, so 76 substantive briefs (Leave-to-Amend, IME, Dismiss, GFS,
  Consolidate, Protective Order, Trial Preference, etc.) are lumped into `other`.
- **Oppose-a-Motion** auto-detects the type via the LLM analyzer, which emits a
  *freeform* label ("Motion to Compel Further Responses"). That never equals the index
  id `compel`, so firm matching usually does not fire.
- **Generate-a-Motion** dropdown offers only `compel`/`demurrer`/`strike` (+ "Other
  (specify…)" → `generic`). MSJ/Ex Parte/IME/etc. are unselectable, and `generic` has
  no briefs in the index — so those firm samples are unreachable from Generate.

## Goal

One canonical motion-type vocabulary, used everywhere, so authority + style matching
fires reliably across both tasks and the "other" bucket resolves into useful types.

## Decisions (from brainstorming)

- **Normalizer mechanism:** static, ordered keyword/alias map (deterministic, testable,
  no LLM latency/cost). Unknown → `other`.
- **"other" classification:** by **filename** (the names are descriptive). No
  content/first-page reading. Re-tag the existing index **in place** (no re-extraction).
- **Generate registration:** **light** — register the common types reusing the existing
  generic drafter with a per-type `legal_standard_hint` + basic section plan. No bespoke
  per-type prompts. Existing bespoke compel/demurrer/strike untouched.

## Components

### 1. `icharlotte_core/firm_briefs/motion_taxonomy.py` (single source of truth)
- `CANONICAL_TYPES`: ordered list of `(id, display_name, [keyword regexes])`, most-specific
  first (so "summary judgment" → `msj` before generic "motion"). Absorbs and supersedes
  path_meta's `_TYPE_ALIASES`.
- `normalize_motion_type(text: str) -> str`: lowercases, strips, returns the first matching
  canonical id, else `"other"`. Handles freeform labels, folder labels, and dropdown names.
- `display_name(id) -> str`.
- Canonical ids (initial set): `msj, compel, demurrer, strike, in_limine, quash, sanctions,
  relieve_counsel, continue_trial, ex_parte, leave_to_amend, ime, gfs, dismiss, consolidate,
  protective_order, set_aside_default, reconsider, other`. (Pleadings ids stay in path_meta
  for completeness but are out of the motion index scope.)

### 2. Wire the normalizer into ingest / oppose / generate
- **path_meta.py:** delegate type resolution to `normalize_motion_type`. Side logic stays.
  For the `Motions - Other` folder and `Oppositions|Replies / Other` subfolders, normalize
  the **filename** (fall back to `other`) so subtypes get real ids on future ingests.
- **oppose_motion_page.py:** keep the editable "Motion type" field showing the analyzer's
  freeform label (human-readable), but normalize it (`normalize_motion_type(metadata.motion_type)`)
  at the point it's passed to `_make_firm_provider`, `_firm_style_exemplars`, and
  `research_arguments(motion_type=…)`.
- **generate_motion_page.py:** dropdown ids are already canonical; additionally normalize the
  "Other (specify…)" custom name so typing "MSJ" there resolves to `msj` and matches.

### 3. Re-tag the existing index in place
- One-off script `retag_firm_index.py`: for every brief with `motion_type='other'`, set
  `motion_type = normalize_motion_type(<original filename>)`; UPDATE only (no re-extraction,
  no re-embed — the profile vector and citations are unchanged). Reports the before/after
  distribution. Idempotent.

### 4. Register common types for Generate (light)
- Add `MotionTypeConfig` entries to `config.BUILTIN_SEED` for: `msj, ex_parte, ime, gfs,
  dismiss, leave_to_amend, consolidate, quash, sanctions, continue_trial, protective_order`.
  Each: `type_id` (== canonical id), `display_name`, a `legal_standard_hint` (the governing
  CCP/CRC standard), and a basic `section_plan`; analyzer/drafter reuse the **generic**
  engine (no new prompts). `get_motion_config` keeps returning `generic` for unknowns.
- These then appear in `list_motion_types()` → the dropdown, and their ids match the index.

## Data flow (after)

```
Ingest:   path -> path_meta(folder->side; folder|filename -> normalize -> type) -> index
Oppose:   analyze_motion -> freeform motion_type (shown, editable)
                         -> normalize_motion_type() at match -> provider/style/research (canonical id)
Generate: dropdown id (canonical)  OR  Other->normalize(custom name) -> provider/style/research
```

## Testing

- `normalize_motion_type`: a freeform→id table (msj/compel/demurrer/strike/in_limine/quash/
  sanctions/relieve_counsel/continue_trial/ex_parte/leave_to_amend/ime/gfs/dismiss/consolidate/
  protective_order/set_aside_default/reconsider, plus unknown→other; plus realistic phrasings
  like "Defendant's Notice of Motion and Motion for Summary Judgment").
- `path_meta`: existing folder mappings still hold; `Motions - Other/<file>` now subclassifies
  by filename; `_Support`/`_Other` still excluded.
- re-tag script: on a temp index seeded with `other` briefs, the IME/Leave/Dismiss/etc. rows
  get reclassified; truly-generic ones stay `other`; idempotent on re-run.
- config: new types in `list_motion_types()`; `get_motion_config(id)` returns them; unknown →
  generic; existing compel/demurrer/strike unchanged.
- oppose/generate wiring: the normalized id (not the freeform label) is what's passed to the
  firm provider/style helpers (unit test the normalization hop).
- Full regression: firm_briefs + opposition + wizard suites green.

## Execution after merge
- Run `retag_firm_index.py` once against the real index; report the new distribution.
- Restart iCharlotte to load the new dropdown types + normalization.

## Out of scope
- Content/first-page classification (filename-only suffices).
- Bespoke per-type analyzer/draft prompts.
- Phase 3 citation-panel UI + Workbench Sample Library tab.
- 3800 library build.
