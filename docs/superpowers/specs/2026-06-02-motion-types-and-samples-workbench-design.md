# Editable Motion Types + Samples in the Prompts Workbench

**Date:** 2026-06-02
**Status:** Approved design, pending implementation plan
**Scope:** Make Generate Motion's motion types, style samples, and prompt templates editable through the Prompt Engineering Workbench (`icharlotte_core/ui/dialogs.py`), so users can add new motion types and improve generation without code changes.

## Problem

Generate Motion's motion types are a hardcoded Python dict (`MOTION_TYPE_CONFIGS`), its drafter ignores style samples (passes `style_exemplars=[]`), and its prompt templates live in code. Users cannot add motion types, attach sample documents to steer drafting, or tune prompts from the UI.

## Goals

- Add new motion types (full fidelity) and edit existing ones from the Workbench.
- Attach style sample documents per motion type to improve drafting.
- Edit the Generate Motion prompt templates (draft + analyzer) from the Workbench.
- No code changes required for any of the above once shipped.

## Non-Goals

- No changes to the Oppose a Motion workbench behavior (only generalize shared hooks).
- No new sample-management UI — reuse the existing `StyleExamplesTab`.

## Decisions (from brainstorming)

- **Full fidelity** custom types: display name, legal-standard text, section plan, placeholder attachments, plus advanced analyzer/grounds prompts.
- **Full migration**: a JSON registry is the single source of truth for motion types, seeded from the current built-ins, all editable, with a **Restore Defaults** safety net.
- **Separate Style Examples tab** reusing `StyleExampleRegistry` (not embedded in the type editor).
- **Prompt editing included**: the Generate Motion draft + analyzer templates become Workbench-editable passes.

## 1. Types registry (data model)

New `icharlotte_core/motion_generation/types_registry.py`:

- `MotionTypeRegistry` mirroring `opposition/style_examples.py:StyleExampleRegistry`:
  - `path` → `Scripts/prompts/generate_motion/motion_types.json`
  - `load(path)` — if file missing, seed from `_BUILTIN_SEED` and save; else parse.
  - `save()`, `list_types()`, `get(type_id)` (fallback to `generic`), `add/update/remove(type_id)`, `restore_defaults()` and `restore_default(type_id)` (re-seed from `_BUILTIN_SEED`).
- `MotionTypeConfig` gains `to_dict()` / `from_dict()` for JSON.
- `config.py`: keep the four current configs as `_BUILTIN_SEED`. `get_motion_config(type_id)` is rewritten to read a **module-level cached singleton** `MotionTypeRegistry`, with `reload_motion_types()` to refresh the singleton after Workbench edits. The canonical registry path is resolved once (project `Scripts/prompts/generate_motion/`).
- **Backward compatibility**: all existing callers (`analyzer`, `drafter`, `assembler`, `generate_motion_page`) keep calling `get_motion_config(type_id)` unchanged; only its internals change.

### JSON schema (per type)
```json
{
  "type_id": "compel",
  "display_name": "Motion to Compel Further Responses",
  "target_doc_guidance": "…",
  "legal_standard_hint": "…CCP 2030.300 / 2031.310…",
  "section_plan": ["Introduction", "Statement of Facts", "Legal Standard", "Argument", "Conclusion"],
  "placeholder_attachments": ["Meet and Confer Declaration", "Separate Statement (…rule 3.1345)"],
  "analyzer_prompt": "…",
  "grounds_prompt": "…"
}
```
`generic` remains a seeded entry and the fallback for unknown ids.

## 2. Samples wiring

- New `Scripts/prompts/generate_motion/style_examples.json` managed by the existing `StyleExampleRegistry` (samples are `.docx` paths tagged by motion-type `type_id`, with an active flag).
- `GenerateMotionWorker.run` loads the registry, calls `matches_for_motion_type(type_id)` + `extract_exemplar_text(path, cache_dir=…)` (cached), and passes the resulting exemplars into `draft_motion(..., style_exemplars=exemplars)` — replacing today's empty list. This mirrors `OpposeMotionWorker` (oppose_motion_page.py).

## 3. Prompt-template editing

- Register two Generate Motion passes in the prompt registry (`Scripts/prompts/registry.json`) seeded from code defaults:
  - `generate_motion:draft_motion` ← `motion_generation/prompts.py:MOTION_DRAFT_PROMPT`
  - `generate_motion:analyze_target` ← a templatized version of the analyzer's user prompt (convert `analyzer._build_user_prompt` to a single template string with named placeholders: `motion_type`, `analyzer_prompt`, `grounds_prompt`, `legal_standard`, `target_text`, `context_text`).
- `drafter.draft_motion` uses `get_prompt("generate_motion", "draft_motion") or MOTION_DRAFT_PROMPT`.
- `analyzer.analyze_target` uses `get_prompt("generate_motion", "analyze_target") or DEFAULT_ANALYZE_TEMPLATE`.
- Seeding happens idempotently on first Workbench load (the Workbench already runs `_migrate_if_needed`); create versions via `PromptManager.create_version` only if the pass is absent.

## 4. Workbench UI

In `icharlotte_core/ui/dialogs.py`:
- Add `generate_motion` to `_populate_agents` and `WORKBENCH_AGENT_MAP` (`"generate_motion": "agent_generate_motion"`; the agent reuses the oppose_motion model sequence by default in `llm_preferences.json`).
- The Editor / A-B / History / Model Defaults tabs work for generate_motion via the two registered passes (section 3).
- Generalize `_refresh_style_examples_tab` to show the **Style Examples** tab for both `oppose_motion` and `generate_motion`, choosing the registry path by agent.
- New `_refresh_motion_types_tab`: when `generate_motion` is selected, show a **Motion Types** tab.

New `icharlotte_core/ui/dialogs_motion_types.py`:
- `MotionTypesTab(registry_path)` — a `QTableWidget` of types (id, display name, # sections, # placeholders, # samples) with **Add / Edit / Remove / Restore Defaults** buttons, backed by `MotionTypeRegistry`. Save → `registry.save()` then `reload_motion_types()`.
- `_MotionTypeEditDialog` — fields: type id (read-only on edit), display name, target-doc guidance, legal-standard text (multiline), section plan (multiline, one per line), placeholder attachments (multiline, one per line), and a collapsible "Advanced" group for analyzer/grounds prompts (multiline, defaulted).

## 5. Intake dropdown

`GenerateMotionSettingsPage` builds its type combo from `MotionTypeRegistry.list_types()` (display name + type_id) instead of the hardcoded `_CONFIGURED_TYPES`, so user-added types appear automatically. The "Other (specify…)" entry remains for one-off custom names routed through `generic`. `_on_type_changed` reads guidance via `get_motion_config`.

## Components / Boundaries

- `motion_generation/config.py` — built-in seed + `MotionTypeConfig` (+ JSON (de)serialization) + `get_motion_config` / `reload_motion_types` (singleton accessors).
- `motion_generation/types_registry.py` — the editable registry (load/seed/save/CRUD/restore).
- `opposition/style_examples.py` — reused unchanged for samples.
- `ui/dialogs_motion_types.py` — the Motion Types editor tab + dialog.
- `ui/dialogs.py` — agent registration + tab show/hide hooks.
- `motion_generation/prompts.py` + `analyzer.py` + `drafter.py` — templatized prompts read via `get_prompt` with code-default fallback.

## Testing

- **Pure logic:** registry seeds from built-ins when file absent; `to_dict`/`from_dict` round-trip; `add/update/remove/restore_defaults`; `get_motion_config` reflects registry edits after `reload_motion_types`; unknown id → generic.
- **Samples:** a fake-LLM `draft_motion` asserts matched exemplar text appears in the draft prompt; `matches_for_motion_type` selects by type_id; empty registry → no exemplars (drafts fine).
- **Prompts:** `draft_motion` / `analyze_target` use a registered override when present and fall back to the code default when absent (monkeypatched `get_prompt`).
- **pytest-qt:** Motion Types tab add → new type persists and appears; edit → fields round-trip; Restore Defaults → built-ins return; the intake dropdown lists a newly added type after `reload_motion_types`.

## Risks / Notes

- A user could delete or corrupt a built-in type; **Restore Defaults** (per-type and global) is the mitigation, and `_BUILTIN_SEED` always lives in code.
- `get_motion_config` moving from a static dict to a cached singleton: callers in worker threads must see edits — `reload_motion_types()` is called after any Workbench save; the singleton load is process-wide.
- Registry path resolution must work both in-app and under pytest (resolve project `Scripts/prompts/generate_motion/`, creating the dir on first save).
