# Generate Motion — Citation Review Parity

**Date:** 2026-06-04
**Status:** Approved (design)
**Branch (proposed):** `feature/generate-motion-citation-review`

## Problem

The **Draft a Motion** (Generate Motion) wizard task lacks the citation-review
surface that **Oppose a Motion** has: the verdict **summary banner**, the
**color-coded clickable citation underlines** in the draft body, and the
**right-side `CitationDetailPanel`** (with find-replacement and open-source
actions).

The data needed for that review *already exists*: `GenerateMotionWorker`
already extracts citations, runs `pool_membership_check`, builds the local/
CourtListener verifier, runs `verify_all`, sorts by body offset, and calls
`enrich_with_pool_signals` — populating `draft.citations` with the same
`CitationVerification` objects oppose produces. The only consumer that throws
that data away is `GenerateMotionOutputPage`, which renders
`draft.body_text` as **plain text** and ignores `draft.citations`.

A secondary gap: the generate worker researches only the **top-level grounds**
(`metadata.principal_arguments`), not each selected outline subsection, so
fewer propositions get grounded in real authority → fewer citations to review.
Oppose researches subsection granularity via `_research_targets(metadata, plan)`.

## Goals

1. Generate Motion's output page shows the **same** citation-review UI as
   Oppose: summary banner, color-coded clickable cites, right-side detail
   panel, save-with-red-flag warning.
2. The generate worker researches at the **same granularity** as oppose, so the
   review surface is as substantive.
3. The two tasks **stay** in parity: a future improvement to one should reach
   the other without a second port. (This drift is what created the request.)
4. No regression to the shipped Oppose a Motion feature.

## Non-goals

- Persisting verification verdicts across close/reopen. A reopened `.docx`
  carries no verdict metadata, so the panel shows its empty-state — **identical
  to oppose's current behavior**. Out of scope.
- Adding a DEV-only "Re-verify" button to generate. Oppose's is dev cruft
  slated for removal; generate does not get one.
- Changing the verification engine, parser, or verdict semantics. They are
  already shared and already run in the generate worker.
- Word-output validation changes: `assemble_motion_preview` already validates
  internally via `validate_opposition_docx` and raises on errors, which the
  worker surfaces. No change needed.

## Architecture (Approach A — shared toolkit + base output page)

All citation-review UI lives in `oppose_motion_page.py` today by historical
accident; it is entirely side-agnostic. We extract it into a shared module and
have both task pages subclass a common base.

### 1. New module: `icharlotte_core/ui/wizard/pages/citation_review.py`

Moved verbatim from `oppose_motion_page.py` (logic unchanged):

- **Body-render helpers:** `_HORIZONTAL_RULE_RE`, `_MD_HEADING_RE`,
  `_MD_ITALIC_RE`, `_VERDICT_COLORS`, `_color_for_verdict`,
  `_render_draft_html`, `_build_citation_index`, `_format_inline_html`.
  - One tweak: `_render_draft_html`'s title fallback becomes generic
    (`draft.title or "Memorandum"`) instead of hardcoding
    `"Opposition Memorandum"`. The drafter sets `draft.title` in both tasks, so
    this only affects the rare empty-title path.
- **Panel helpers:** `_VERDICT_HEADER_COLORS`, `_VERDICT_LABELS`,
  `_citation_header_html`, `_citation_body_html`, `_run_find_replacement`.
  - `_run_find_replacement` already uses `agent_id="agent_oppose_motion"`;
    generate reuses that same agent, so it is correct for both unchanged.
- **Widgets:** `CitationDetailPanel`, `CitationDetailDialog`.
- **New base class `CitationReviewOutputPage(QWidget)`** holding the common
  page glue currently inside `OpposeMotionOutputPage`:
  - widgets: `summary_banner` (QLabel), `editor` (QTextBrowser, links off,
    `anchorClicked` → `_on_anchor_clicked`), `detail_panel`
    (`CitationDetailPanel`), laid out editor:panel = 2:1 under the banner.
  - methods: `show_result(draft)`, `show_citation(index)`,
    `_on_anchor_clicked(url)` (parses `citation:N`), `_refresh_summary_banner()`,
    `load_output(path)` (reconstruct `DraftDocument` from `.docx`, no
    citations), `output_path` property, `default_save_dir(preview_path)`
    (static), and a parameterized `save_as()` with the red-flag warning.
  - overridable seams:
    - `default_title: str` — class attr (`"Memorandum"` default).
    - `empty_citations_message() -> str` — returns the placeholder shown in the
      panel when `draft.citations` is empty.
    - `_build_action_buttons(row: QHBoxLayout) -> None` — subclass adds its own
      buttons (Save, Open in Word, etc.); base provides a default Save button
      via a small helper so subclasses can call/extend it.
  - The base constructor builds banner + editor + panel + an action-button row,
    calling `_build_action_buttons(row)` so each subclass populates it.

### 2. `oppose_motion_page.py` (refactor, behavior preserved)

- `OpposeMotionOutputPage(CitationReviewOutputPage)`:
  - `default_title = "Opposition Memorandum"`.
  - `empty_citations_message()` returns the existing opposition copy.
  - `_build_action_buttons()` adds the **DEV-only Re-verify** button (and its
    `_ReverifyWorker` handlers, which stay in `oppose_motion_page.py`) plus the
    **Save** button. `save_as()` keeps opposition's suggested filename.
  - `show_result` override (if needed) only to keep the empty-state message;
    otherwise inherits.
- **Backward-compat re-exports** at module level so existing imports/tests that
  do `from ...oppose_motion_page import _render_draft_html` (and
  `CitationDetailPanel`, `CitationDetailDialog`, `_citation_header_html`,
  `_citation_body_html`, `_run_find_replacement`, `_color_for_verdict`,
  `_VERDICT_COLORS`, etc.) keep resolving:
  `from .citation_review import (...)  # noqa: F401`.
  - Tests that depend on this today:
    `tests/test_wizard/test_oppose_motion_output_page_verdicts.py`
    imports `_render_draft_html`, `OpposeMotionOutputPage`,
    `CitationDetailDialog`, `CitationDetailPanel`, and patches
    `oppose_motion_page._ReverifyWorker` — all must keep working.

### 3. `generate_motion_page.py`

**Output page** — replace the bare `GenerateMotionOutputPage` with:

- `GenerateMotionOutputPage(CitationReviewOutputPage)`:
  - `default_title = "Generated Motion"`.
  - `empty_citations_message()` returns motion-specific copy (e.g. "No
    citations were detected in this motion. If California case-law research
    returned no results, the motion was drafted without case citations…").
  - `_build_action_buttons()` adds a **Save** button (parity, with the
    red-flag warning inherited from the base) **and** keeps **Open in Word**
    (`QDesktopServices.openUrl(QUrl.fromLocalFile(preview_path))`).
  - `load_output(path)` inherited from base; override only to re-enable the
    Open-in-Word button (toggle on valid preview path).
  - Preserves the existing `output_path` property and `load_output(path)`
    signature required by `iCharlotte.py` reopen wiring
    (`_snapshot_open_task_tabs` reads `tab.output_page.output_path`;
    `_restore_task_tabs_for_case` calls `load_output`).

**Worker** — `GenerateMotionWorker.run()` research parity:

- Import and use `_research_targets` from `oppose_motion_page`
  (precedent: `_make_local_corpus` is already imported from there). Replace the
  `metadata.principal_arguments` argument to `research_arguments` with
  `_research_targets(metadata, plan)` (args ∪ selected section-plan leaves,
  structural sections skipped, capped at 24). `plan` is already computed.
- Pass `cache_dir` to `research_arguments`:
  `Scripts/prompts/generate_motion/.cache/opinions` (mirrors oppose's
  `Scripts/prompts/oppose_motion/.cache/opinions`; gitignored).
- Drop `max_workers` 4 → 2 for `research_arguments` (oppose's rate-limit
  lesson: 4 parallel workers burst the LLM/CourtListener throttle).
- Verification block, assembler, preview path, and `task_completed` emission
  are unchanged.

> Note: `_research_targets` and `_make_local_corpus` currently live in
> `oppose_motion_page.py`. They are worker/research helpers, not UI, but moving
> them is out of scope here; importing them (as already done for
> `_make_local_corpus`) keeps this change small. A future cleanup could relocate
> them to a research-helpers module.

## Data flow (unchanged spine, now consumed)

```
intake → analyze_target → merge → outline → [Generate]
  → GenerateMotionWorker:
      extract context
      research_arguments(_research_targets(metadata, plan), cache_dir, max_workers=2)
      draft_motion(...)               → draft.body_text
      extract_citations + verify_all  → draft.citations  (already present)
      assemble_motion_preview(...)    → draft.preview_path (validates internally)
  → GenerateMotionOutputPage.show_result(draft)
      _render_draft_html(draft)       → colored clickable cites   (NEW consumer)
      _refresh_summary_banner()       → verdict counts            (NEW)
      detail_panel.set_citation(...)  → on anchor click           (NEW)
```

## Testing

New `tests/test_wizard/test_generate_motion_output_page.py`, mirroring the
oppose verdict tests:

- `_render_draft_html` colors: SUPPORTED→green, NOT_SUPPORTED→red,
  PARTIAL→yellow, UNVERIFIED→gray (reuse the shared helper).
- `summary_banner` counts per verdict after `show_result`.
- `GenerateMotionOutputPage` exposes a working `detail_panel`; clicking a
  `citation:N` anchor calls `show_citation` and updates the panel.
- `save_as` warns when a citation is flagged red (monkeypatch
  `QMessageBox.question` + `QFileDialog.getSaveFileName`).
- Empty-citations path shows the motion-specific placeholder, not the
  opposition copy.

Worker test (extend `tests/test_wizard/test_generate_motion_page.py` or a new
worker test): given metadata + an outline with selected subsection leaves,
`_research_targets(metadata, plan)` yields more than just the top-level
arguments (assert a subsection leaf is included, structural sections excluded).

Regression: run the full `tests/test_wizard/` suite (esp.
`test_oppose_motion_output_page_verdicts.py`) to confirm the extraction +
re-exports did not break oppose.

Environment: tests use the venv `C:\geminiterminal2\.venv\Scripts\python.exe`
(`pytest.importorskip("PySide6")` / `"pytestqt"`). Stop the running iCharlotte
before a full collection — a launched app instance breaks PySide6 import in
pytest collection. Per `worktree_vs_main_checkout`, edits land in the main
checkout `C:\geminiterminal2\`; restart iCharlotte to see them live.

## Risks & mitigations

- **Refactoring the shipped oppose page** → guarded by its existing verdict
  test file; keep all public names + behaviors; re-export moved symbols.
- **Re-export omissions** breaking an import → enumerate the moved names and
  re-export the full set; run the oppose tests.
- **Research call-volume / rate limits** with subsection granularity on the
  live CourtListener fallback → mitigated by `max_workers=2` and the existing
  process-wide throttle in `courtlistener.py`; the local corpus path is
  unaffected (offline, unlimited).
