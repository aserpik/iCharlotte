# Wizard Separate → persistent, viewable document Index

- **Date:** 2026-06-04
- **Status:** Design — awaiting user review
- **Approach:** A (reveal the existing advanced `IndexTab` singleton; add persistence)

## Problem

When a user runs the **Separate Documents** task in **Wizard Mode**, the detected
document map (the table of page-ranges / dates / titles / merge-groups) is shown in
an embedded `SeparatorWorkbench` and then discarded. Nothing is written to the
per-case index store, so once the task tab closes the analysis is gone. Advanced
Mode's `IndexTab` *does* persist (via `run_separator_path` → `IndexTab.add_pdf`),
but Wizard Mode has no equivalent and no way to browse past runs.

The user wants Wizard-mode Separate output to be **accessible in the future**, via a
small button on the bottom-right of the **Separate Documents launcher card** that
opens an "Index" tab — essentially the same Index tab that exists in Advanced Mode
(left: list of processed source PDFs; middle: editable table of separated docs;
right: PDF viewer + Mark/Process controls).

## Goals

1. Wizard-mode Separate runs persist to the same per-case store Advanced Mode reads,
   so a run in either mode is visible in both (single source of truth).
2. A corner button on the Separate Documents launcher card opens the Index for the
   current case, available at any time without re-running the task.
3. In Wizard Mode, the revealed Index tab can be dismissed by clicking its "x"
   (which **re-hides** the singleton, never destroys it).

## Non-goals

- No new Wizard-themed Index UI (that was Approach B/C). We reveal the existing
  classic `IndexTab`.
- No change to how splitting/merging works.
- No refactor of `IndexTab`'s own persistence (kept untouched to minimize regression
  risk in working Advanced-mode code). The new store module simply writes the same
  on-disk format `IndexTab` already reads.
- No auto-reload of the Advanced `IndexTab`'s in-memory state when toggling modes
  mid-session; the reveal path always reloads from disk, which covers the feature.

## Current state (verified)

- **Store:** `{file_number}_index.json` in `GEMINI_DATA_DIR` (`.gemini/case_data`,
  relative to cwd). Shape: `{pdf_path: [doc, ...]}`. Each `doc` = `{id, title, date,
  start, end}`. Written/read by `IndexTab.save_data`/`load_data`/`add_pdf`
  (`icharlotte_core/ui/tabs.py`).
- **Advanced `IndexTab`** is a persistent singleton created at startup
  (`iCharlotte.py:874-875`), loaded per-case via `load_data(file_number)`, and merely
  *hidden* in Wizard Mode — `"Index"` is in `_WIZARD_HIDDEN_TABS`
  (`iCharlotte.py:1025-1032`). Layout = `[Processed-PDFs QListWidget] |
  [SeparatorWorkbench]` + "Save Edits to Index".
- **Wizard `SeparateTaskTab`** (`icharlotte_core/ui/wizard/pages/separate_page.py`)
  runs `SeparateAnalysisWorker` → `_on_analysis_finished` → `workbench.load_docs(...)`;
  `_on_processing_complete` emits `task_completed` but writes nothing to the store.
  It already holds `self._file_number` and `self._pdf_path`.
- **Launcher** (`icharlotte_core/ui/wizard/wizard_tab.py`) builds one `TaskCard`
  (`task_card.py`, 280×140 `QFrame`, whole surface clickable → `clicked(task_id)` →
  `WizardTab.task_requested` → `iCharlotte._open_task_tab`). The Separate spec id is
  `"separate"` (`registry.py`).
- **Tab close mechanics:** `setTabsClosable(True)` global (`iCharlotte.py:606-607`);
  `_hide_fixed_close_buttons()` hides the "x" on every tab lacking a `wizard_task_id`
  property; `_on_tab_close_requested(index)` returns early for such tabs. So fixed
  tabs (incl. Index) currently show no "x" and can't be closed.

## Design

### 1. Shared store module — `icharlotte_core/case_index_store.py` (new)

Small, UI-free module so both modes and tests use one writer:

```python
from icharlotte_core import config  # read GEMINI_DATA_DIR at call time

def index_path(file_number: str) -> str: ...
def load_index(file_number: str) -> dict: ...        # {} on missing/corrupt
def upsert_pdf(file_number: str, pdf_path: str, docs: list) -> None:
    # load_index, data[pdf_path] = docs, makedirs, json.dump(indent=4)
```

- Reads `config.GEMINI_DATA_DIR` via attribute access (not a frozen import binding)
  so tests can monkeypatch `icharlotte_core.config.GEMINI_DATA_DIR`.
- Matches `IndexTab`'s exact format: same path, same `{pdf_path: [docs]}` shape,
  `json.dump(..., indent=4)`, UTF-8.
- `IndexTab` is **not** modified to call this module (Approach A minimal-risk). Both
  resolve the same path in production, so they stay consistent on disk.

### 2. Wizard Separate task writes to the store

In `separate_page.py` `SeparateTaskTab`, add a private helper and call it at two
points (both guarded on a truthy `self._file_number`):

- `_persist_to_index(docs)`: `case_index_store.upsert_pdf(self._file_number,
  self._pdf_path, docs)`, wrapped in try/except + `log_event` on failure (never break
  the task flow on a persistence error).
- **`_on_analysis_finished(success, payload)`** — on success, after
  `workbench.load_docs(self._pdf_path, docs)`, call `_persist_to_index(docs)`. Captures
  the run even if the user analyzes then closes without splitting.
- **`_on_processing_complete(summary)`** — before emitting `task_completed`, re-read
  the current (possibly edited) table rows from the workbench and persist them, so the
  store reflects edits the user made before splitting:

  ```python
  wb = self.workbench
  edited = [wb._get_doc_from_row(r) for r in range(wb.doc_table.rowCount())]
  self._persist_to_index(edited)
  ```

  (Mirrors `IndexTab.save_table_to_index`. Rows with unparsable page ranges yield
  `start/end = None`; we still store them — the Index tab tolerates and lets the user
  fix them, same as Advanced mode after an edit.)

### 3. Launcher card corner button

- **`TaskSpec`** (`registry.py`) gains three optional fields (default `None`):
  `card_action_id`, `card_action_glyph`, `card_action_tooltip`. Adding fields (not a
  task) does not affect `test_registry.py::test_initial_tasks_registered`.
- Only the `"separate"` spec sets them:
  `card_action_id="open_separate_index"`, `card_action_glyph="🗂"`,
  `card_action_tooltip="Open the document Index for this case"`.
- **`TaskCard`** (`task_card.py`): when `spec.card_action_id` is set, render a small
  `QToolButton` (flat, the glyph as text, the tooltip) pinned bottom-right (a footer
  `QHBoxLayout` with a leading stretch). New signal `action_requested = Signal(str)`
  emitting `spec.card_action_id`. Because a `QToolButton` accepts its own mouse press,
  the card's `mousePressEvent`/`clicked` does **not** fire for clicks on the button
  (verified by test). Cards without a `card_action_id` are visually unchanged.
- **`WizardTab`** (`wizard_tab.py`): new signal `card_action_requested = Signal(str)`;
  in `_build_ui`, also `card.action_requested.connect(self.card_action_requested.emit)`
  for each card. Spec-driven — no hardcoded task-id in launcher logic.

### 4. Reveal the Index tab

- In `iCharlotte._build...` (where `wizard_tab.task_requested` is wired,
  ~`iCharlotte.py:656`), add
  `self.wizard_tab.card_action_requested.connect(self._on_card_action)`.
- New `_on_card_action(action_id: str)` dispatches; for `"open_separate_index"` →
  `_reveal_index_tab()`.
- New `_reveal_index_tab()`:
  ```python
  if not self.case_path:
      QMessageBox.information(self, "No case loaded",
          "Open a case from the Master List first.")
      return
  idx = self._index_of_tab("Index")
  if idx < 0:
      return
  if self.file_number:
      self.index_tab.load_data(self.file_number)   # refresh from disk → shows wizard run
  self.tabs.setTabVisible(idx, True)
  self.tabs.setCurrentIndex(idx)
  self._hide_fixed_close_buttons()                 # (re)show the Index "x" in wizard mode
  ```

### 5. Re-hide via the tab "x" (wizard mode only)

The Index tab is the singleton `self.index_tab`; closing must hide, not destroy.

- **`_hide_fixed_close_buttons()`**: a tab shows its "x" when
  `is_task_tab OR (widget is self.index_tab AND self.mode_controller.is_wizard AND
  self.tabs.isTabVisible(i))`. Use `getattr(self, "index_tab", None)` for a defensive
  identity check (helper may run before `index_tab` exists during init).
- **`_on_tab_close_requested(index)`**: before the existing `wizard_task_id` early
  return, add:
  ```python
  if widget is getattr(self, "index_tab", None) and self.mode_controller.is_wizard:
      self.tabs.setTabVisible(index, False)
      wiz = self._index_of_tab("Wizard")
      if wiz >= 0:
          self.tabs.setCurrentIndex(wiz)
      return
  ```
  Never `removeTab`/`deleteLater` the singleton.
- **`_apply_mode_visibility(mode)`**: append a `self._hide_fixed_close_buttons()` call
  at the end so the Index "x" is suppressed when toggling back to Advanced mode (where
  Index is a permanent, non-closeable tab).

## Testing

Per repo convention (PySide6; UI verified via offscreen `MainWindow()` smoke tests —
running the app breaks pytest's PySide6 import, see memory). Use
`QT_QPA_PLATFORM=offscreen`.

1. **`tests/test_case_index_store.py`** (new): `load_index` empty/missing/corrupt →
   `{}`; `upsert_pdf` round-trips; file is readable by the same `json.load` shape
   `IndexTab.load_data` uses; `GEMINI_DATA_DIR` monkeypatch is honored.
2. **`tests/test_task_card_action.py`** (new): a card built from a spec **with**
   `card_action_id` shows the button and emits `action_requested(action_id)`; clicking
   the button does **not** emit `clicked`; a card **without** the field shows no button
   and still emits `clicked` on body press.
3. **`WizardTab`** test: `card_action_requested` re-emits the separate card's action id.
4. **Separate persistence** test: drive `SeparateTaskTab._on_analysis_finished(True,
   docs)` and `_on_processing_complete({...})` with a temp store; assert
   `case_index_store.load_index(file_number)[pdf_path]` equals the docs / edited rows.
   Guard path: empty `file_number` → no write, no crash.
5. **Reveal + re-hide** smoke test (offscreen `MainWindow()`): with a case loaded in
   wizard mode, `_on_card_action("open_separate_index")` makes the Index tab visible
   and current and gives it a visible "x"; simulating `_on_tab_close_requested(idx)`
   hides it (tab still exists, count unchanged) and selects the Wizard tab.

Run: `python -m pytest tests/test_case_index_store.py tests/test_task_card_action.py
-q` plus the smoke tests; then a full `python -m pytest tests/ -q` sanity pass.

## Risks / edge cases

- **Stale Advanced in-memory state:** if the user runs a wizard Separate, then toggles
  to Advanced without the reveal path, the Advanced tab's in-memory `index_data` is
  stale until its next `load_data`. Out of scope (reveal always reloads). Noted as a
  known minor limitation.
- **No case loaded:** reveal shows an info dialog and no-ops.
- **Concurrent writes:** in Wizard mode the hidden Advanced tab isn't being edited, and
  reveal reloads before display, so no clobber. If the user edits in the revealed tab
  and clicks "Save Edits to Index", `IndexTab` writes its in-memory map (which already
  includes the loaded wizard run) — consistent.
- **`QToolButton` click isolation** is the one behavior to verify explicitly (test 2).

## Rollout / process notes

- **Shared checkout hazard:** `C:\geminiterminal2` is a live shared checkout and the
  current branch (`feature/generate-motion-detailed-outline`) is unrelated to this work,
  with many uncommitted changes from other sessions. Implementation should happen on a
  dedicated branch — preferably an isolated **git worktree** (per the Separate→Wizard
  precedent) — to avoid disrupting a concurrent session. Confirm branch strategy with
  the user before committing code. The running app uses the main checkout, so after
  implementing, edits must be present in `C:\geminiterminal2` and iCharlotte restarted
  to verify live.
- Per global rule: commit only when the user asks.
