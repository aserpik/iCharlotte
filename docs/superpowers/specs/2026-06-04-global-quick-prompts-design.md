# Global Quick Prompts — Design

**Date:** 2026-06-04
**Status:** Approved
**Branch / worktree:** `feature/global-quick-prompts` (`.claude/worktrees/global-quick-prompts`, based on `main`)

## Problem

Custom chat "quick prompts" (templates) added via **Chat tab → Templates → Manage Templates** are
stored **per-case** in `.gemini/case_data/{file_number}_chat.json` through
`ChatPersistence.add_quick_prompt()`. The "Manage Templates" dialog
(`PromptTemplateDialog`) and the template dropdown (`ChatTab.update_template_menu`) both read and
write custom templates through the current case's `persistence` object. Consequences:

- A template added in one case never appears in any other case.
- Templates do not carry across to a freshly opened case.
- The dialog refuses to save with "No case loaded. Cannot save prompts." when no case is open.

The user wants custom quick prompts to **persist globally** — across sessions and across cases —
so the same templates are available everywhere.

## Goal

Custom (non–built-in) quick prompts persist in a single global store shared by every case and every
session. Built-in prompts (Transcribe, Mediation Brief) remain hardcoded and read-only, exactly as
today.

## Non-goals

- No Prompt Engineering Workbench integration / version history for templates (a heavier alternative
  that was explicitly set aside in favor of the dedicated store).
- No change to built-in prompts or to the Mediation Brief special-case flow.
- No change to how conversations, messages, settings, or attached files are stored (those stay
  per-case in `ChatPersistence`).

## Approach (chosen)

A dedicated, single-purpose global store backed by one JSON file. This was chosen over routing
templates through `PromptManager` because it is a smaller, focused change with a near drop-in
interface, it keeps chat templates decoupled from the prompt-engineering system, and it avoids
encoding template metadata into unrelated version fields.

## Components

### 1. `GlobalQuickPromptStore` (new)

- **File:** `icharlotte_core/chat/global_prompts.py`
- **Backing file:** `<.gemini>/global_quick_prompts.json`, i.e. one level **above** the per-case
  `case_data/` folder (`os.path.dirname(GEMINI_DATA_DIR)`). Living above `case_data/` signals that
  it is app-global, not case data, and keeps it out of the `case_data/*_chat.json` glob used by
  migration.
- **JSON shape:**
  ```json
  {
    "version": "1.0",
    "migrated_from_per_case": true,
    "quick_prompts": [
      {"id": "...", "name": "...", "prompt": "...", "category": "Custom", "is_builtin": false}
    ]
  }
  ```
  Entries reuse `QuickPrompt.to_dict()` / `QuickPrompt.from_dict()` from
  `icharlotte_core/chat/models.py`.
- **Public API** (names mirror the `ChatPersistence` quick-prompt methods so callers change
  minimally):
  - `get_quick_prompts() -> List[QuickPrompt]` — returns the stored custom prompts (all
    `is_builtin=False`). Built-ins are **not** returned; callers add `BUILTIN_PROMPTS` themselves,
    exactly as they do today.
  - `add_quick_prompt(name, prompt, category='Custom') -> str` — returns the new prompt ID.
  - `update_quick_prompt(prompt_id, name=None, prompt=None, category=None)`.
  - `delete_quick_prompt(prompt_id)`.
- **Behavior:**
  - Reads fresh from disk on every call. The file is tiny (a handful of templates), so re-reading is
    negligible, and it keeps multiple chat tabs — and even a second app instance — consistent
    without a cache-invalidation scheme.
  - Writes atomically: serialize to a temp file in the same directory, then `os.replace()` onto the
    target, so a crash mid-write cannot corrupt the store.
  - Creates the parent directory if missing.
- **Constructor:** `GlobalQuickPromptStore(path=None, case_data_dir=None, auto_migrate=True)`.
  `path` and `case_data_dir` default to the real locations; tests pass temp dirs. `auto_migrate`
  lets tests construct a store without triggering migration.
- **Singleton accessor:** `get_global_quick_prompt_store()` returns a module-level shared instance,
  mirroring `get_prompt_manager()`.

### 2. One-time migration (in the store)

On first construction, if `migrated_from_per_case` is not set in the JSON:

1. Scan `case_data/*_chat.json`.
2. From each file's `quick_prompts`, take every entry with `is_builtin == False`.
3. Deduplicate by `name` (first occurrence wins, including against any names already in the global
   store).
4. Add the survivors to the global store.
5. Set `migrated_from_per_case = true` and save.

Properties:

- **Non-destructive:** per-case files are never modified. Existing templates are *copied forward*;
  the UI simply stops reading the per-case copies. Fully reversible.
- **Idempotent:** the flag guards against re-running, so templates a user later deletes globally are
  not resurrected on the next launch. (If the global JSON is deleted, migration re-runs from the
  per-case files — acceptable recovery behavior.)
- **No-op safe:** with no per-case templates, migration just sets the flag.

### 3. Wiring changes

- **`PromptTemplateDialog`** (`icharlotte_core/ui/chat_dialogs.py`):
  - `load_prompts`, `save_prompt`, `delete_prompt` use `get_global_quick_prompt_store()` instead of
    `self.persistence` for custom-template CRUD.
  - The `persistence` constructor parameter is kept (made optional, default `None`) for backward
    compatibility but is no longer used for templates.
  - The "No case loaded. Cannot save prompts." guard is removed — saving now always works.
  - Built-in prompts continue to come from `BUILTIN_PROMPTS` and stay read-only.
- **`ChatTab.update_template_menu`** (`icharlotte_core/ui/tabs.py`): custom prompts come from the
  global store. Additionally, rebuild the menu on the button's `aboutToShow` signal so a template
  added in one tab/case becomes available in every chat tab without a restart.
- **`ChatTab.open_template_manager`** (`icharlotte_core/ui/tabs.py`): opens the dialog without
  depending on per-case persistence (passing `self.persistence` is harmless and may be kept).
- **`ChatPersistence`** per-case quick-prompt methods (`get_quick_prompts`, `add_quick_prompt`,
  `update_quick_prompt`, `delete_quick_prompt`) are left in place for backward compatibility and to
  preserve existing data, but are no longer called by the chat UI.

## Data flow

```
Add/edit/delete template (Manage Templates dialog)
    -> GlobalQuickPromptStore.{add,update,delete}_quick_prompt()
    -> atomic write to <.gemini>/global_quick_prompts.json

Open template dropdown (any case, any session)
    -> ChatTab.update_template_menu()
    -> BUILTIN_PROMPTS (hardcoded) + GlobalQuickPromptStore.get_quick_prompts()
    -> menu actions insert prompt text into the chat input
```

## Testing

- **`tests/test_global_quick_prompts.py`** — pure Python, no Qt (avoids the known PySide6
  pytest-collection issue):
  - add / get / update / delete round-trip.
  - persistence across two separate store instances pointed at the same file (simulates a restart →
    proves cross-session persistence).
  - empty state returns `[]`.
  - atomic write produces valid JSON.
  - migration from sample `*_chat.json` files: imports custom prompts, skips built-ins, dedupes by
    name, and runs only once (flag honored on a second construction).
- **UI wiring** — verified with an offscreen (`QT_QPA_PLATFORM=offscreen`) standalone smoke script
  that constructs `PromptTemplateDialog`, adds a template, and confirms it lands in the global store
  and survives a new store instance. (Per project convention, the running app breaks pytest's
  PySide6 import, so UI checks use a standalone offscreen script, not pytest.)
- **Real app** — launch iCharlotte from the main checkout (after porting the changes there) and
  confirm a template added in one case appears in another, per the project's "always test after a
  change" rule.

## Delivery / git

- Built in the `feature/global-quick-prompts` worktree (based on `main`) for isolation from the
  concurrent generate-motion session in the main checkout. `main` has no committed changes to the
  chat files versus the generate-motion branch, so the work is independent and the branch is cleanly
  mergeable.
- The finished changes are also applied to the main checkout (`C:\geminiterminal2`) and the app
  restarted so the user can use the feature immediately ("apply to both" workflow for this repo).
