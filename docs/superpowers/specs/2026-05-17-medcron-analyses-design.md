# Med-Cron Multi-Analysis Design

**Date:** 2026-05-17
**Status:** Approved for planning

## Summary

Extend the Med-Cron agent from a single "rewrite the narrative chronology" task into a multi-analysis task. After a user selects a medical chronology file, the wizard presents a list of selectable analyses — both a curated catalog (Rewrite Chronology, Inconsistency Check, Treatment Gap Detector, etc.) and user-typed custom analyses. Selected analyses run in parallel and each produces its own docx.

The existing Rewrite Chronology behavior is preserved as one of the selectable analyses, using only the narrative-filtered text. Other analyses use both the narrative AND the tables from the original chronology document.

## Goals

- Add a "Med Chron Analysis" task to the wizard following the two-phase pattern used by Summarize Depositions.
- Support a curated catalog of analyses defined in code, plus user-added custom analyses defined per-session.
- Run selected analyses in parallel inside a single Phase 2 process.
- One docx output per analysis.
- Preserve the existing `python med_chron.py <file>` CLI behavior (rewrite-only) so the older IndexTab agent runner keeps working.

## Non-goals

- Cross-chronology analysis (analysis spanning multiple chronology files at once). Each chronology runs independently in its own task tab.
- DocumentRegistry registration for analysis outputs in v1 (current Med-Cron does not register; adding it is a separate concern).
- Cache management UI (cleaning up `.med_chron/` cache folders is a future feature).
- LLM-discovered analyses. The catalog is curated; only custom analyses are dynamic.

## Architecture

One subprocess per chronology file (managed by the wizard). Inside Phase 2, a `ThreadPoolExecutor` fans out the selected analyses concurrently because `LLMCaller.call` is I/O-bound (network requests).

```
[Wizard "Med Chron Analysis" task]
    └─ User selects chronology file(s) — one TaskTab per file
        └─ MedChronSettingsPage opens
            ├─ launches Phase 1 subprocess speculatively
            └─ shows analysis picker UI

Phase 1 subprocess (med_chron.py --phase=prep <file>):
    extract narrative-only text + full-with-tables text → write
    session JSON listing catalog → print AWAITING_INPUT → exit

[MedChronSettingsPage]
    user picks analyses + custom rows → writes user_config back to
    session JSON, sets phase=ready_to_run → emits phase2_requested

Phase 2 subprocess (med_chron.py --phase=run <session.json>):
    load both cached texts → build run list → ThreadPoolExecutor fans
    out one LLM call per analysis → each writes its own docx
```

## Components

### Scripts/med_chron.py — two-phase agent

Reworked from its current single-pass form into three CLI modes:

| Invocation | Mode | Behavior |
|---|---|---|
| `med_chron.py <file>` | Legacy | Runs only Rewrite, writes `med_chron_<file>.docx`. Preserves existing IndexTab path. |
| `med_chron.py --phase=prep <file>` | Phase 1 | Extracts text twice (narrative + full), writes session JSON, prints `AWAITING_INPUT:<path>`. |
| `med_chron.py --phase=run <session.json>` | Phase 2 | Reads session, runs selected analyses in parallel, writes one docx per analysis. |

### Scripts/MED_CHRON_ANALYSES/catalog.py — curated catalog

A Python module (not JSON) so multi-paragraph prompts can be authored as triple-quoted strings with full IDE support.

```python
@dataclass(frozen=True)
class AnalysisDef:
    id: str            # stable slug for filenames + session JSON
    title: str         # shown in checkbox UI
    description: str   # short tooltip under the checkbox
    uses_tables: bool  # True → full text; False → narrative only
    prompt_file: str   # filename in MED_CHRON_ANALYSES/prompts/
    default_selected: bool = False

CATALOG: list[AnalysisDef] = [
    AnalysisDef(
        id="rewrite_chronology",
        title="Rewrite Chronology (readable narrative)",
        description="Reformats the pre/post-injury summary into a clean narrative.",
        uses_tables=False,
        prompt_file="rewrite_chronology.txt",
        default_selected=True,
    ),
    AnalysisDef(
        id="inconsistencies",
        title="Inconsistency Check",
        description="Flags contradictions between narrative and table entries.",
        uses_tables=True,
        prompt_file="inconsistencies.txt",
    ),
    AnalysisDef(
        id="treatment_gaps",
        title="Treatment Gap Detector",
        description="Identifies unexplained gaps in treatment dates.",
        uses_tables=True,
        prompt_file="treatment_gaps.txt",
    ),
    # … additional analyses authored later
]
```

### Scripts/MED_CHRON_ANALYSES/prompts/ — prompt files

- `rewrite_chronology.txt` — existing `Scripts/MED_CHRON_PROMPT.txt` content, moved here.
- `inconsistencies.txt`, `treatment_gaps.txt`, … — authored later.
- `_custom_wrapper.txt` — generic template for user-typed analyses. Contains `{user_instruction}` placeholder. Phase 2 substitutes the user's text and feeds the full chronology.

### icharlotte_core/ui/wizard/pages/med_chron_settings_page.py

`MedChronSettingsPage(SettingsPage)`. Mirrors `DepositionSettingsPage`:

- `QStackedWidget` with two pages: a "Preparing chronology…" indicator (Phase 1 in flight) and the config form (after `AWAITING_INPUT`).
- `attach_worker(worker)` wires `worker.awaiting_input` → swap stack to form; `worker.failed` → show error.
- Proceed button disabled until Phase 1 completes.
- `_on_proceed` calls `form.commit_user_config()` then emits `phase2_requested(session_path)`.

### icharlotte_core/ui/med_chron_config_form.py

`MedChronConfigForm` — embeds inside the settings page.

- Reads session JSON, builds checkbox list from `session.catalog`. Each entry: `QCheckBox` + small description label.
- Rewrite Chronology pre-checked (only one with `default_selected=True`).
- "Custom analyses" panel: vertical list of `CustomAnalysisRow` widgets, each with label `QLineEdit` + instruction `QPlainTextEdit` + remove button. "Add custom analysis" button below.
- Banner shown when `session.narrative_missing == true`: "Narrative text not found in this document — Rewrite Chronology will be skipped."
- `commit_user_config()` validates:
  - At least one selection (curated or custom) — otherwise inline red error, no proceed.
  - Each custom row must have both label and non-empty instruction. Empty rows silently dropped.
  - On success, writes `user_config` to session JSON, sets `phase=ready_to_run`, returns True.

### icharlotte_core/ui/wizard/registry.py — new task entry

Add to `TASK_REGISTRY`:

```python
"med_chron_analysis": TaskSpec(
    task_id="med_chron_analysis",
    title="Med Chron Analysis",
    description="Run selectable analyses on a medical chronology.",
    icon_glyph="\U0001FA7A",  # 🩺
    script_name="med_chron.py",
    default_folders=["NOTES/AI OUTPUT", "RECORDS"],
    _settings_page_cls_factory=_med_chron_settings_page_cls,
),
```

## Data flow

### Phase 1 (prep)

1. Resolve case output dir using the existing case-folder detection logic in `med_chron.py`.
2. Compute session paths:
   ```
   <output_dir>/.med_chron/<file_hash>/
       narrative.txt   ← cached narrative-only text
       full.txt        ← cached narrative+tables text
       session.json    ← phase state + user_config
   ```
   `<file_hash>` = sha1(`abspath(input) + str(mtime_ns(input))`), truncated to 12 hex chars. Touching the source file invalidates the cache. The source file's mtime is captured at the start of Phase 1 and never mutated by the agent.
3. Extract narrative-only text: reuse existing `extract_text()` (PDF: pypdf + OCR fallback; .docx: paragraphs only). Run existing `filter_content()` to slice the PRE/POST-INJURY synopses. Write to `narrative.txt`.
4. Extract full-with-tables text:
   - `.docx` → `icharlotte_core.document_processor.extract_docx_text()` (the canonical full-text extractor that includes tables as pipe-separated rows).
   - `.pdf` → same `extract_text()` result as step 3 (PDFs don't have the paragraph-vs-table distinction in extraction).
   - `.doc` → Word COM read helper following `ChatTab._extract_doc_text` pattern (no `word.Quit()`, no `word.Visible` flip).
   - Write to `full.txt`.
5. Detect provider name from filename via existing `extract_provider_from_filename`.
6. Write session JSON:
   ```json
   {
     "version": 1,
     "phase": "awaiting_input",
     "input_path": "...",
     "narrative_text_path": "...narrative.txt",
     "full_text_path": "...full.txt",
     "narrative_missing": false,
     "provider_name": "...",
     "file_number": "1234.567",
     "catalog": [
       {"id": "rewrite_chronology", "title": "...", "description": "...",
        "uses_tables": false, "default_selected": true},
       ...
     ],
     "user_config": null
   }
   ```
7. Print `AWAITING_INPUT:<session.json path>`, exit 0.

If `filter_content()` returns None (no PRE/POST-INJURY headings), `narrative.txt` is written empty and `narrative_missing: true` is set. Phase 1 still succeeds — the user can run non-rewrite analyses against `full.txt`.

If full text extraction fails (corrupt docx), Phase 1 fails (exit 1). Without full text, no table-using analyses work.

### Phase 2 (run)

1. Load session.json. Bail if `phase != "ready_to_run"`.
2. Read `narrative.txt` and `full.txt` from disk.
3. Resolve catalog entries from `user_config.selected_catalog_ids` by importing `catalog.py` and looking up each id (source of truth for `prompt_file` and `uses_tables` is Python, not the JSON snapshot).
4. Build run list. `safe_basename` and `slug()` both use the same sanitizer as today's `med_chron.py`: `re.sub(r"[^a-zA-Z0-9_\-]", "_", value).strip("_")`, lowercased for slugs.
   ```python
   safe_basename = sanitize(os.path.splitext(os.path.basename(input_path))[0])

   for cat_id in user_config.selected_catalog_ids:
       definition = CATALOG_BY_ID[cat_id]
       runs.append(RunSpec(
           id=cat_id,
           title=definition.title,
           prompt_text=read(prompts/<prompt_file>),
           input_text=narrative if not definition.uses_tables else full,
           output_filename=f"med_chron_{cat_id}_{safe_basename}.docx",
       ))
   for i, c in enumerate(user_config.custom_analyses, 1):
       label_slug = slug(c.label)  # e.g. "Left-knee mentions" → "left_knee_mentions"
       runs.append(RunSpec(
           id=f"custom_{i}_{label_slug}",
           title=c.label,
           prompt_text=CUSTOM_WRAPPER.replace("{user_instruction}", c.instruction),
           input_text=full,
           output_filename=f"med_chron_custom_{i}_{label_slug}_{safe_basename}.docx",
       ))
   ```
   The custom output filename includes the index `i` so two custom rows with identical labels can't collide on disk.
5. Skip-with-warning: if `rewrite_chronology` is in runs but `narrative.txt` is empty, drop it and log "Skipping Rewrite Chronology — no pre/post-injury synopsis headings found."
6. Parallel execution:
   ```python
   with ThreadPoolExecutor(max_workers=min(len(runs), 4)) as ex:
       futures = {ex.submit(_run_one, r, llm_caller, logger): r for r in runs}
       for f in as_completed(futures):
           result = f.result()
           logger.progress(...)
   ```
   `max_workers` capped at 4 to stay safe with provider rate limits. `_run_one` wraps the LLM call in try/except; failures return `RunResult(success=False, error=str(e))` so siblings continue.
7. Each successful run:
   - `save_to_docx(content, output_path, provider_name, run.title)` reusing the existing `add_markdown_to_doc` + `docx_writer.get_docx_lock` pattern from `summarize_deposition.save_to_docx`.
   - `CaseDataManager.save_variable(file_num, var_key=f"med_chron_{cat_id}_{safe_provider}", content=summary, source="med_chron_agent", extra_tags=["Evidence", "Medical Records", run.title])`.
8. Progress message: `<N> of <M> analyses complete (<X> failed)`.
9. Exit 0 if any run succeeded; exit 1 only if all runs failed.

Session JSON is kept after Phase 2 (matching deposition behavior) so the user can re-open the popup and run additional analyses without re-extraction.

## LLM call discipline

Per the cross-cutting lesson in MEMORY.md: every `_run_one` call passes `prompt` and `text` as separate arguments to `LLMCaller.call(prompt=..., text=...)`. Never pre-concatenated. This applies to both catalog analyses and the custom-wrapper-prefixed instruction.

## Error handling

**Phase 1:**
- Text extraction returns None / empty → log error, exit 1. Wizard `failed` signal fires.
- `filter_content` returns None → session.json written with `narrative_missing: true`; not a failure.
- Full text extraction fails → log error, exit 1.

**Phase 2:**
- Per-run failure → caught in `_run_one`, logged with analysis id, `RunResult(success=False)`. Siblings keep running.
- All runs fail → exit 1.
- At least one succeeds → exit 0; status page shows per-analysis breakdown.
- File-lock contention on output docx → reuse `docx_writer.get_docx_lock` auto-version-bump from `summarize_deposition.save_to_docx`.
- LLMCaller already does multi-provider fallback (Gemini → Claude → OpenAI). No retry layer on top.

## Testing

```
tests/test_med_chron/
    test_phase1_prep.py
        - test_phase1_writes_session_with_both_text_caches
        - test_phase1_narrative_only_excludes_table_content
        - test_phase1_full_text_includes_table_rows
        - test_phase1_missing_synopsis_marks_narrative_missing
        - test_phase1_prints_awaiting_input_token_on_success
        - test_phase1_reuses_cache_on_unchanged_input
    test_catalog.py
        - test_catalog_ids_are_unique
        - test_every_catalog_entry_has_prompt_file_on_disk
        - test_only_rewrite_is_default_selected
    test_phase2_runner.py
        - test_run_one_uses_narrative_when_uses_tables_false
        - test_run_one_uses_full_text_when_uses_tables_true
        - test_per_run_failure_does_not_abort_siblings
        - test_all_runs_failed_exits_nonzero
        - test_partial_success_exits_zero
        - test_skips_rewrite_when_narrative_missing
        - test_custom_analysis_wraps_user_instruction_in_template
        - test_output_filenames_use_analysis_id_slug
        - test_max_workers_capped_at_4
    test_legacy_cli_compatibility.py
        - test_no_phase_arg_runs_rewrite_only
        - test_no_phase_arg_writes_existing_filename_pattern

tests/test_wizard/
    test_med_chron_settings_page.py
        - test_proceed_disabled_until_phase1_completes
        - test_form_appears_after_awaiting_input_signal
        - test_proceed_requires_at_least_one_selection
        - test_empty_custom_rows_dropped_on_commit
        - test_custom_row_requires_label_and_instruction
        - test_narrative_missing_banner_shown_when_marker_set
    test_med_chron_registry.py
        - test_med_chron_analysis_task_registered
        - test_task_uses_med_chron_settings_page
```

Wizard tests use `pytest.importorskip("pytestqt")` (no underscore — per MEMORY.md, the underscored form silently skips the whole file).

## Backward compatibility

`python med_chron.py <file>` (no `--phase` flag) remains valid. In this mode the script runs only the Rewrite analysis using narrative-only text and writes to the existing `med_chron_<file>.docx` filename. Existing IndexTab callers and any external scripts unaffected.

## Open items deferred to implementation

- Exact icon glyph for the wizard task card (placeholder 🩺).
- Final wording on the "narrative missing" banner.
- Whether to surface per-analysis progress in the status page or only an aggregate `<N> of <M>` message. Start with aggregate; per-analysis can be added if it feels needed during testing.
