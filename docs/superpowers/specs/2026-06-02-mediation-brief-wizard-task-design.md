# Mediation Brief — Wizard Task

**Date:** 2026-06-02
**Status:** Approved design, pending implementation plan
**Scope:** Port the existing defense-side Mediation Brief *generation* feature (today reachable only via the Chat tab → Templates dropdown in Advanced mode) into a first-class Wizard Mode task card. Generation only; the existing `MediationBriefGenerator` engine is reused unchanged.

## Problem

The Mediation Brief generator is a powerful, well-tuned feature, but it is buried in the Chat tab's Templates dropdown and only available in Advanced mode. Wizard Mode is now the default landing mode for a loaded case, and every other major drafting workflow (Oppose a Motion, Generate a Motion, Respond to Discovery, Separate) is a Wizard card. The Mediation Brief should be discoverable and runnable the same way: pick documents → generate → review → save.

## Goals

- Add a **"Mediation Brief"** Wizard task card under **"Motions & Drafting"**.
- Drive the **unmodified** `MediationBriefGenerator` engine from a Wizard `Settings → Status → Output` flow.
- Let the user select source documents (multi-folder) and confirm/override the caption template before generating.
- Save the finished brief to a predictable default location, with a "Save a Copy As…" escape hatch.
- Mirror the most robust sibling task (`GenerateMotionTaskTab`) for worker lifecycle and close-safety, because generation takes minutes.

## Non-Goals (v1)

- **No conversational refinement** in the Wizard. Section refinement is already provided by the Word AI Assistant (Win+V → "Mediation Brief: Refine Section") on the generated document.
- **No in-Wizard "Add Quotes."** Grounded deposition-quote insertion is already provided by the Word AI Assistant (Win+V → "Add Quotes") on the generated document.
- **No changes to the Chat tab feature.** The existing Templates → "Mediation Brief" entry (with its conversational refinement + Add Quotes) stays exactly as-is. Both entry points coexist.
- **No changes to `MediationBriefGenerator`, its workers, or the prompts.** The Wizard adds a new driver only.

## Decisions (from brainstorming)

- **Scope:** Generation only. Refinement/quotes live in the Word assistant.
- **Chat entry:** Keep both. Non-destructive — the Chat tab flow is untouched.
- **Default save target:** `<case>/NOTES/AI OUTPUT/MEDIATION/` (auto-saved), plus a "Save a Copy As…" button for a user-chosen location.

## Task Wiring

In-process task (`script_name=""`), wired exactly like the other in-process builders.

- **`registry.py`** — new `TaskSpec`:
  - `task_id="mediation_brief"`, title **"Mediation Brief"**, category **"Motions & Drafting"**.
  - `icon_glyph` = 🤝 (`"\U0001F91D"`).
  - `keywords=["mediation","brief","settlement","mediator","defense","confidential"]`.
  - `default_folders=[]` (the settings page manages its own multi-folder selection; no pre-Settings picker).
  - No `_settings_page_cls_factory` needed — in-process builders construct their own pages.
  - The launcher card appears automatically (`wizard_tab.py` builds cards from `list_tasks()`); no launcher edit required.
- **`task_routing.py`** — add `"mediation_brief": "build_mediation_brief_tab"` to `_IN_PROCESS_TASK_BUILDERS`. This routes `_open_task_tab` through the generic in-process branch (no pre-Settings file picker — the page owns source selection).
- **`in_process_task_tab.py`** — add a thin `build_mediation_brief_tab(spec, case_path, file_number, parent)` wrapper that delegates to the new page module (mirrors the existing `build_generate_motion_tab` / `build_separate_tab` wrappers).

## New Module: `ui/wizard/pages/mediation_brief_page.py`

Contains everything specific to the task: the settings page, the worker, the output page, the tab, and the builder. Page-index constants `TASK_PAGE_SETTINGS=0`, `TASK_PAGE_STATUS=1`, `TASK_PAGE_OUTPUT=2`.

### `MediationBriefSettingsPage(QWidget)` — emits `run_requested(dict)`

- Helper text describing the task.
- **Source documents**: a `QListWidget` + "Add Files…" / "Remove" buttons, using `ContextFilesDialog.get_files(parent, title=…, start_dir=case_root, file_filter="Documents (*.pdf *.docx *.doc *.txt *.msg);;All files (*.*)")`. Multi-folder accumulation lets the user pull transcripts, pleadings, and records together. De-dupe on add.
- **Caption template**: a field pre-filled by `MediationBriefGenerator().find_caption_template(case_path)`, with a "Browse…" button to override (`.docx` filter, start at case root). Shown in `theme.ERROR` red when empty/missing, per the wizard design-system affordances.
- **Generate Mediation Brief** primary button → validates (≥1 source file AND a caption path that exists), then emits:
  ```python
  {
    "files": [...],                 # absolute paths
    "caption_path": "<.docx>",
    "save_default_dir": "<case>/NOTES/AI OUTPUT/MEDIATION",
    "suggested_filename": "Defendant's Confidential Mediation Brief.docx",
  }
  ```
- `to_dict()` / `from_dict()` for persistence parity (files + caption path).

### `MediationBriefWizardWorker(QThread)`

Signals: `progress = Signal(str)`, `finished_result = Signal(bool, object)` (payload = output path on success, error string on failure). Method `cancel()` sets a flag checked between sections.

`run()` pipeline (each step emits a `progress` line):

1. **Read source documents** → a single `document_content` string. Table-aware extraction is mandatory for fidelity (mediation briefs lean on tabular chronologies / depo summaries): `.docx` via `extract_docx_text` (document-order, table-aware), `.pdf` via fitz, `.txt` direct, `.doc` via read-only Word COM (never `Quit`/`Visible`), `.msg` via Outlook COM. This mirrors the Chat tab's `read_files_content` so the Wizard feeds the engine identical input. Per-file failures are reported via `progress` and skipped.
2. Bail with `finished_result(False, …)` if no content was extracted.
3. Construct `MediationBriefGenerator()`; set `caption_template_path` (from settings) and `document_content`.
4. `get_style_excerpts()` (loads/builds the style cache).
5. `run_planning_pass()` (after a cancel check).
6. Loop `GENERATION_ORDER`: `progress("Generating <Heading> (i of n)…")` (heading via `SECTION_HEADINGS`), cancel-check, `generate_section(name)`, store into `generator.sections`.
7. Set `generator.is_active = True`.
8. **Assemble + save** to the default target (see below); emit `finished_result(True, saved_path)`.
9. Any exception → `finished_result(False, str(exc))`. Cancellation between steps → `finished_result(False, "Generation cancelled.")` and stop.

### Output location & file-lock safety

- Target: `<case>/NOTES/AI OUTPUT/MEDIATION/Defendant's Confidential Mediation Brief.docx` (folder created if absent).
- `assemble_document(caption_path, target)` already runs `validate_report` (satisfies the mandatory Word-validation rule).
- **Re-run while the brief is open in Word**: assemble to a temp file, then move it into place. If the destination is locked (`PermissionError`), fall back to a counter-suffixed name (`…Brief (2).docx`) and return the actual path used — never silently fail, never close the user's Word (global safety rule).

### `MediationBriefOutputPage(QWidget)`

Read-only review + save affordances (bespoke, because the save model differs from the base `OutputPage`'s single Save button).

- Read-only rendered preview via `load_docx_as_html` (reused from `ui/wizard/docx_io.py`).
- A **"Saved to: NOTES/AI OUTPUT/MEDIATION/…"** banner so the location is obvious.
- Buttons:
  - **Open in Word** — `os.startfile` the saved file.
  - **Save a Copy As…** — `QFileDialog` defaulting to the MEDIATION folder; copies the saved file to the chosen path (the canonical file stays in MEDIATION).
  - **Re-run** / **Edit Settings & Re-run** — emit `rerun_requested` / `edit_settings_requested` for parity.
- Exposes `output_path` property + `load_output(path)` + `show_result(path)` so reopen and `_snapshot_open_task_tabs` (which reads `tab.output_page.output_path`) work.

### `MediationBriefTaskTab(WizardTaskContainer)`

Mirrors `GenerateMotionTaskTab`:

- Pages: `MediationBriefSettingsPage`, `StatusPage`, `MediationBriefOutputPage`.
- `settings_page.run_requested → _on_run`: reset status, indeterminate progress bar, switch to Status, start the worker **parented to `None`** (so tab deletion can't orphan a running QThread).
- `status_page.cancel_requested → worker.cancel()` (a working Cancel; the engine stops at the next section boundary).
- `worker.progress → status_page.on_status`; `worker.finished_result → _on_worker_finished` with `sender()`/`_finishing_worker` guards (the generate_motion pattern).
- On success: `output_page.show_result(path)`, switch to Output, emit `task_completed` (entry dict: task_id, title, files, settings, output_path, completed_at).
- On failure: `status_page.on_status("FAILED: …")` (or "Cancelled.").
- `closeEvent` guard refuses to close while the worker runs.
- Public API parity: `spec` property, `files` property (current source files).
- Builder `build_mediation_brief_tab(spec, case_path, file_number, parent)` returns the tab on its Settings page.

## iCharlotte.py wiring

- **Open** (`_open_task_tab`): no change — the generic `in_process_builder_name` branch already calls `build_mediation_brief_tab` (no pre-Settings picker).
- **Reopen** (`_on_reopen_recent_task`) and **Restore** (`_restore_task_tabs_for_case`): no change — the generic "in-process builder (non-oppose)" branch already creates a fresh tab on the Settings page. (Same v1 behavior as Separate / Generate Motion: settings/output are **not** restored on reopen; the user re-picks and re-runs. Acceptable for v1.)
- **Close guard** (`_on_tab_close_requested`): add `"MediationBriefTaskTab"` to the class-name tuple (currently `("OpposeMotionTaskTab", "GenerateMotionTaskTab", "SeparateTaskTab")`) so closing the tab mid-generation shows "wait for it to finish" instead of tearing down a running thread. This is the actual protection, because that handler uses `removeTab` + `deleteLater` (it does not call `tab.close()`, so the tab's own `closeEvent` would not fire on the X click).

## Components / Boundaries

- `registry.py` — task registration only.
- `task_routing.py` — one dict entry.
- `in_process_task_tab.py` — one thin builder wrapper.
- `ui/wizard/pages/mediation_brief_page.py` (new) — owns the settings page, worker, output page, tab, and builder. Depends on `mediation_brief.MediationBriefGenerator` (+ `GENERATION_ORDER`, `SECTION_HEADINGS`), `ContextFilesDialog`, `StatusPage`, `WizardTaskContainer`, `theme`, `docx_io.load_docx_as_html`, and `document_processor.extract_docx_text`.
- `iCharlotte.py` — one tuple entry (close guard).
- **Unchanged:** `mediation_brief.py`, the Chat tab, the Word AI Assistant.

## Testing

`tests/test_wizard/test_mediation_brief_page.py` (Qt binding is **PySide6**; gate Qt tests with `pytest.importorskip("pytestqt")`):

- **Registry / routing (pure logic, no Qt):**
  - `mediation_brief` is registered, category `"Motions & Drafting"`, has keywords, `script_name == ""`.
  - `get_in_process_task_builder_name("mediation_brief") == "build_mediation_brief_tab"`; the attribute exists on `in_process_task_tab`.
  - `filter_tasks` still groups everything (existing `test_task_categories.py` stays green — assertions are dynamic against `list_tasks()` and Discovery's count is unaffected).
- **Settings page (pytest-qt):**
  - Caption auto-detect pre-fills from a fake case folder containing a `*Caption*.docx`.
  - Add/Remove files; de-dupe; `to_dict`/`from_dict` round-trip.
  - Validation: Generate is blocked (warning) with no files or no caption; emits `run_requested` with the expected payload otherwise.
- **Worker (pytest-qt or direct, mocked generator):**
  - Patch `MediationBriefGenerator` (and document extraction) so no real LLM/IO runs; assert the progress sequence (read → style → planning → each section → assemble), the final `finished_result(True, path)` with the path under `NOTES/AI OUTPUT/MEDIATION`, that `cancel()` stops before remaining sections, and that an engine exception yields `finished_result(False, msg)`.
  - Lock fallback: when the destination move raises `PermissionError`, the worker returns a counter-suffixed path and still reports success.
- **Output page (pytest-qt):**
  - `show_result(path)` sets `output_path`, renders, shows the banner; "Save a Copy As…" copies to a chosen path while leaving the original in place.

## Risks / Notes

- **Text-extraction fidelity** is the main correctness risk. The Wizard must feed the engine the same text the Chat tab does; `.docx` extraction must be table-aware (`extract_docx_text`), not a `doc.paragraphs` loop (silent data loss on chronologies/summaries). Confirm during implementation whether to reuse a shared helper vs. `opposition.extract_context_bundle` (verify the latter's docx table-awareness before adopting; it omits `.doc`).
- **Deposition-quote source restriction** is enforced inside the engine's prompts (quotes only from real transcripts). No Wizard-side handling needed, but note it when documenting.
- **Style cache** is read from `C:\AI\Mediation Briefs`; if absent, `get_style_excerpts()` degrades gracefully (same as the Chat flow).
- **Caption discovery** relies on a `*caption*.docx` in the case folder; the Browse override is the fallback (surfaced on the settings page, not mid-run).
- **Reopen does not restore prior output** in v1 (fresh tab), consistent with sibling in-process tasks.
- **Long runs**: generation is multi-minute; the indeterminate progress bar + per-section status lines + working Cancel + close guard are deliberate.
