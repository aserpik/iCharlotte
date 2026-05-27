# Wizard Mode — Design Spec

**Date:** 2026-05-17
**Status:** Approved design, ready for implementation planning

## Summary

iCharlotte today exposes ~12 fixed tabs (Master List, Case View, Status, Index, Chat, Email, Email Update, Depositions, Discovery, Liability & Exposure, Templates / Resources, Logs). Power users navigate efficiently; new users (and the primary user in many workflows) want a guided "pick a task → answer some questions → see the result" experience.

This spec adds a new **Wizard Mode** alongside the existing UI (renamed **Advanced Mode**). Both modes use the same underlying engines (LLM caller, master DB, document processor, existing task agents); only the surface changes.

## Goals

- A toggle between Advanced Mode (current UI, unchanged) and Wizard Mode (new guided UI).
- In Wizard Mode, tasks are selected from a card grid and run inside dynamically created "task tabs" rather than scattered across pre-existing tabs.
- Each task tab walks the user through: pick files → configure settings → watch progress → review/edit/save output.
- Open task tabs, recent task history, and the chosen mode all survive app restarts.

## Non-goals

- No changes to the underlying task agents (`Scripts/summarize.py`, `Scripts/summarize_discovery.py`, `Scripts/summarize_deposition.py`, `Scripts/med_record.py`). They are wrapped by thin worker shims.
- No changes to where outputs land on disk (`<case>/NOTES/AI Output/`).
- No redesign of any Advanced-Mode tab. Wizard Mode is additive.
- Task-specific Settings-page content is out of scope for this spec. Each task gets a placeholder Settings page; per-task settings will be defined in follow-up specs.

## Architecture

### New module: `icharlotte_core/ui/wizard/`

```
icharlotte_core/ui/wizard/
├── __init__.py
├── mode_controller.py     # ModeController (QSettings-backed)
├── wizard_tab.py          # WizardTab: header + card grid + Recent Tasks
├── task_card.py           # TaskCard widget (clickable card)
├── task_tab.py            # TaskTab: QStackedWidget with three pages
├── pages/
│   ├── settings_page.py   # SettingsPage (placeholder for now)
│   ├── status_page.py     # StatusPage (progress bar + log + Cancel)
│   └── output_page.py     # OutputPage (mammoth editor + actions)
├── registry.py            # TASK_REGISTRY: id → metadata + classes
├── persistence.py         # WizardStatePersistence (load/save JSON)
├── runners.py             # Thin wrappers around existing agents
└── file_picker.py         # Default-folder resolution helper
```

### Touched files

- `iCharlotte.py` — orchestrate tab visibility per mode, add `_open_task_tab()`, modify `load_case_by_number()` for state snapshot/restore, modify close handling, remove Change File corner button.
- `icharlotte_core/ui/master_case_tab.py` — add mode toggle inside the Master List content header.
- `requirements.txt` — add `mammoth`.

### Modes

- **Advanced Mode**: current UI, unchanged.
- **Wizard Mode**: only Master List + Wizard + any open task tabs are visible. Status, Index, Chat, Email, Email Update, Depositions, Discovery, Liability & Exposure, Templates / Resources, Logs, and Case View are hidden. The tabs and their state are preserved — switching back to Advanced restores them as they were.

Mode is global (one setting for the whole app), persisted via `QSettings("iCharlotte", "iCharlotte")` under key `app/mode`. Default on first run: `"wizard"`. `ModeController` emits `mode_changed(str)`; `MainWindow` listens and applies tab visibility.

## UI Components

### Master List mode toggle

Inside the Master List tab content (top of the panel), a right-aligned segmented control:

```
                                                  ┌─────────────────┬─────────────────┐
                                                  │  Advanced Mode  │ ▣ Wizard Mode   │
                                                  └─────────────────┴─────────────────┘
```

- Implemented with `QButtonGroup` of two checkable `QPushButton`s, styled as a segmented control.
- Click → `ModeController.set_mode(...)` → `mode_changed` signal.
- The existing **Change File** corner-widget button at `iCharlotte.py:922-929` is removed in both modes. The Win+C hotkey at `iCharlotte.py:464` is retained.

### Wizard tab

```
What would you like to do?

┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
│ 📄  Summarize   │  │ 📋  Summarize   │  │ 🎙  Summarize   │
│     Documents   │  │     Discovery   │  │     Depositions │
│                 │  │                 │  │                 │
│ Short purpose…  │  │ Short purpose…  │  │ Short purpose…  │
└─────────────────┘  └─────────────────┘  └─────────────────┘

┌─────────────────┐
│ 🏥  Medical     │
│     Records     │
│                 │
│ Short purpose…  │
└─────────────────┘

───────────────────────────────────────────────────────────
Recent Tasks
• Summarize Documents  — 2026-05-15 10:42   [Reopen]
• Medical Records       — 2026-05-14 16:08   [Reopen]
```

- Header `QLabel("What would you like to do?")`, large, light weight.
- `TaskCard(QFrame)`: rounded border, hover state, fixed ~280×140, icon tile + title + one-line description. Emits `clicked(task_id)`.
- Cards laid out in a `QGridLayout` (3 columns), wrapped in `QScrollArea`.
- Cards come from `TASK_REGISTRY` — adding a future card requires only a new registry entry plus its `SettingsPage` subclass and `runner`. No `WizardTab` changes.
- **Recent Tasks** section below: divider, label, vertical list of up to 20 completed runs for the current case (newest first). Each row: title, timestamp, [Reopen] button. Reopen creates a new task tab directly on the Output Page bound to the saved `output_path`. Missing file → offer to re-run with saved settings.

### Task tab

`TaskTab(QWidget)` with a `QStackedWidget` containing three pages.

State machine:

```
[card click]
     │
     ▼
[QFileDialog]  ──cancel──► (nothing happens, no tab created)
     │ ok
     ▼
[Settings]  ──Proceed──►  [Status]  ──finished──►  [Output]
     ▲                       │                       │
     └── Edit Settings & ────┘                       │
         Re-run                                      │
     ◄─── cancel ────────────────────────────────────┘
```

**File selection (pre-Settings step):**

1. Click card → `QFileDialog.getOpenFileNames()` rooted at the task's default folder.
2. No file-type filter (all files visible).
3. Multi-select supported within a single folder.
4. Cancel → no task tab is created.
5. OK with ≥1 file → task tab is created; selected files passed into the Settings page.

Default folder resolution (`file_picker.resolve_default_folder(case_root, prefs)`):
- Walks `prefs` (list of `"DISCOVERY/RESPONSES"`-style relative paths) in order, case-insensitive against `os.listdir`.
- Returns the first existing match, or `case_root` if none match.
- Never raises.

Per-task `default_folders` preferences:

| Task                  | Preferences (in order)                            | Fallback   |
|-----------------------|---------------------------------------------------|------------|
| Summarize Documents   | *(none)*                                          | case root  |
| Summarize Discovery   | `DISCOVERY/RESPONSES`, `DISCOVERY`                | case root  |
| Summarize Depositions | `DISCOVERY/TRANSCRIPTS`, `DISCOVERY`              | case root  |
| Medical Records       | `RECORDS`                                         | case root  |

**Multi-instance:** clicking the same card more than once creates additional tabs. Disambiguation suffix: `"Summarize Documents"`, `"Summarize Documents (2)"`, `"Summarize Documents (3)"`, etc., assigned as the lowest unused integer at creation time.

**Closeable:** each task tab shows the standard close `[×]`. Closing a tab on the Settings or Output page is a clean removal. Closing on the Status page cancels the worker first (soft-cancel; see below). Outputs already saved to disk are kept.

#### Settings page (placeholder for now)

- Top section: "Files (N)" with a list of selected paths, [Add files] and [Remove] buttons.
- Body: `QLabel("Settings for <task title> — to be defined")` — placeholder.
- Bottom-right: `[Proceed]` button.
- Per-task `SettingsPage` subclasses will be fleshed out in follow-up work; each provides `to_dict() / from_dict()` for persistence.

#### Status page

- `QProgressBar` (indeterminate by default; tasks that report % update it).
- `QPlainTextEdit` log fed by the worker's `status(str)` signal — mirrors what Advanced Mode's Status tab shows today.
- `[Cancel]` button at the bottom. Cancel → flips the button to "Cancelling…", waits for the worker to acknowledge, then snaps back to Settings.

#### Output page

```
File: Summary 2026-05-15 10:42.docx              [ Open in Word ]
────────────────────────────────────────────────────────────────

  ┌────────────────────────────────────────────────────────┐
  │  <rendered editable .docx content here>                │
  │                                                        │
  │  Patient: John Doe                                     │
  │  ...                                                   │
  └────────────────────────────────────────────────────────┘

[ Copy All ]  [ Re-run ]  [ Edit Settings & Re-run ]  [ Save ]
```

- `QTextEdit` showing the .docx rendered as HTML via **mammoth** (new dependency). Editable by default.
- **Save** writes back to the same `output_path` (overwrite). Implementation: parse the current `QTextDocument` and emit a fresh .docx via `python-docx` (paragraphs, runs, bold/italic, headings, lists). Original template-level styling (theme colors, custom styles) does not survive.
- **Open in Word**: if dirty, prompt to save first, then `os.startfile(output_path)`.
- **Copy All**: copies plain text of editor to clipboard.
- **Re-run**: confirm dialog, then re-runs with the same settings/files; overwrites `output_path` once finished.
- **Edit Settings & Re-run**: flips the tab back to Settings with prior settings pre-filled. Output file is preserved on disk until the next run finishes.

Known editor limitations (acceptable; documented in user-facing tooltip):
- Text boxes, embedded images, complex/nested/merged-cell tables, comments, track changes, and footnotes may render approximately and may be dropped on save.
- For high-fidelity formatting work, use **Open in Word**.

### Worker contract

Each task in the registry has a `runner` callable returning a `QThread`-based worker that exposes:

- `status(str)` — log line.
- `progress(int)` — 0–100.
- `finished(str)` — output path, on success.
- `failed(str)` — error message, on exception.
- `cancelled()` — emitted after a cancel request is honored.

Cancellation is **cooperative / soft**:
- `cancel()` sets `self._cancel_requested = True`.
- Long-running loops in the agent code add a flag check between iterations.
- An in-flight LLM call is allowed to finish (latency up to one LLM round-trip, ~30s); the flag is checked immediately after.
- We do not force-kill threads.
- UI shows "Cancelling…" until the worker confirms.

## Persistence

### Location

`<case_root>/.icharlotte/wizard_state.json`

- Per-case (self-contained: travels with case folder if moved or backed up).
- Hidden-style dot-folder, sorts to the top of directory listings, signals "tool metadata."
- Auto-created on first write.
- Includes a `.icharlotte/README.txt` on creation: *"This folder stores iCharlotte app state for this case. Do not edit manually."*
- Atomic writes (`.tmp` + `os.replace`) so a crash mid-save can't corrupt the file.

### Schema

```json
{
  "version": 1,
  "open_tabs": [
    {
      "task_id": "summarize_discovery",
      "instance_suffix": "",
      "files": ["DISCOVERY/RESPONSES/RFP.pdf", "DISCOVERY/RESPONSES/SROG.pdf"],
      "settings": { "...task-specific dict..." : null },
      "page": "settings",
      "output_path": null
    },
    {
      "task_id": "medical_records",
      "instance_suffix": "(2)",
      "files": ["RECORDS/Hosp_Records.pdf"],
      "settings": { },
      "page": "output",
      "output_path": "NOTES/AI Output/Medical Records 2026-05-15 1042.docx"
    }
  ],
  "recent_tasks": [
    {
      "task_id": "summarize_documents",
      "title": "Summarize Documents",
      "files": ["MEDICAL CHART.pdf"],
      "settings": { },
      "output_path": "NOTES/AI Output/Summary 2026-05-15 1042.docx",
      "completed_at": "2026-05-15T10:42:00"
    }
  ]
}
```

Rules:
- File paths stored **relative to the case root** so the case folder is portable.
- `page` is only ever `"settings"` or `"output"`. A tab that was on `"status"` is reset to `"settings"` on case-switch (running tasks cancel; see below).
- A tab with `page: "output"` requires its `output_path` to exist on disk to restore; if missing, restore as `"settings"`.
- `recent_tasks` is capped at 20; older entries trimmed on every save.

A single `WizardStatePersistence` class (parallel to `ChatPersistence`) handles all reads/writes.

### Mode persistence

Global, in `QSettings`:
- Key: `app/mode`
- Values: `"advanced"` or `"wizard"`
- Default on first run: `"wizard"`

## Lifecycle: case switching and app close

### `MainWindow.load_case_by_number(new_case)` (new lifecycle)

```
1. If a case is currently loaded:
   a. For each open task tab:
      - If page == "status": call worker.cancel(); wait up to 2s for ack.
      - Snapshot the tab: task_id, instance_suffix, files, settings,
        page (settings/output), output_path.
   b. Save snapshot to <old_case>/.icharlotte/wizard_state.json.
   c. Remove all task tabs from the QTabWidget.
2. Load new_case (existing logic — DB, summary, etc.).
3. Read <new_case>/.icharlotte/wizard_state.json if present:
   - Reconstruct each TaskTab in the saved page state.
   - "status" → reset to "settings".
   - "output" with missing output_path → reset to "settings".
4. Activate the appropriate tab based on the current mode:
   - **Wizard Mode** → Wizard tab.
   - **Advanced Mode** → Case View tab (preserves current behavior at `iCharlotte.py:1375`).
```

### Mid-run case switch

- Switch is **not** blocked.
- Cancel signal fires; snapshot stored with `page: "settings"`.
- After the 2-second grace window, switch proceeds even if workers haven't acknowledged. The lingering worker becomes a zombie that finishes on its own; writes are short-circuited by the cancelled flag. Logged at WARN.

### App close

- Same snapshot-then-cancel sequence against the currently-loaded case.
- Force-close on grace-window timeout; Python interpreter shutdown cleans up zombie threads.

### Mode switch (Wizard → Advanced) while a task is running

- Does **not** cancel running tasks or close task tabs.
- Only hides them. Switching back to Wizard restores visibility; workers continue running undisturbed.

## Known limitations (deferred work)

1. **File selection is single-folder.** The pre-Settings `QFileDialog` only lets the user pick files from one folder per click. If the user needs files from, e.g., both `/DISCOVERY/RESPONSES` and `/DISCOVERY/PROPOUNDED` in a single Summarize Discovery run, that's not possible in this design. To be addressed in a follow-up: replace the single dialog with a custom multi-folder picker or with a "+ Add more files" affordance on the Settings page that re-opens the dialog and appends.
2. **Per-task Settings pages are placeholders.** Settings for each of the four initial tasks (Summarize Documents, Summarize Discovery, Summarize Depositions, Medical Records) will be defined in follow-up specs.
3. **Output editor is text-focused.** Complex Word features (text boxes, embedded images, advanced/nested/merged-cell tables, comments, track changes, footnotes) display approximately and may not survive a Save round-trip. **Open in Word** is the supported path for high-fidelity edits.
4. **Cancellation is soft.** A cancel request may take up to ~30s to take effect if the worker is mid-LLM-call. Acceptable for this use case; hardening (abortable HTTP requests through `LLMCaller`) is deferred.
5. **No "Reopen old task tab from another case" UX.** Recent Tasks is per-case. Cross-case task history is out of scope.

## Open questions resolved during brainstorming

- Mode scope → **global** (`QSettings`).
- Task-tab lifecycle → **closeable**, with **Recent Tasks history** per case.
- Mid-run case switch → **cancel + restore to Settings page** next session.
- Multi-instance → **yes**, auto-numbered suffixes.
- Toggle location → **inside Master List content** (not the corner widget).
- Other tabs in Wizard Mode → **hidden**.
- Output Page → **rich in-app editor + action buttons** (with Open in Word fallback).
- Default mode on first run → **wizard**.
- State storage → **`<case>/.icharlotte/wizard_state.json`** (hidden dot-folder).
- mammoth dependency → **OK to add**.
- Save behavior → **overwrite `output_path`** (no versioning).
- Cancellation → **soft** (let in-flight LLM call finish).

## Testing strategy

- **Unit tests:**
  - `mode_controller.py` — get/set/persist; signals fire on change only, not on idempotent set.
  - `persistence.py` — load/save round-trip; missing-file resilience; corrupt-file recovery; cap to 20 recent tasks; atomic-write behavior (write-and-crash simulation).
  - `file_picker.py` — `resolve_default_folder` across exact match, case-mismatch match, missing folder, empty prefs.
  - `registry.py` — all four tasks register; suffix collision math (1st, 2nd, 3rd instance).

- **Integration / UI tests** (`pytest-qt`):
  - Toggle Advanced ↔ Wizard flips tab visibility correctly.
  - Click a card → file dialog appears (mock the dialog) → cancel = no tab; OK = tab created on Settings page.
  - Proceed → Status → Output transitions.
  - Cancel on Status page returns to Settings.
  - Closing a task tab on Status calls `worker.cancel()`.
  - Restart simulation: open tabs → simulate `closeEvent` → reload case → tabs restored to expected pages.

- **Manual UI verification** (per `CLAUDE.md` mandatory testing rule):
  - Run the app in Wizard Mode, run each of the four task cards end-to-end against a real case, verify the .docx in `NOTES/AI Output/`.
  - Round-trip a small edit: open Output Page, change a name, Save, reopen in Word, confirm change persists.
  - Verify Open in Word doesn't disturb the user's other Word windows (per global "never close Word windows" rule).
  - Switch cases mid-run and confirm graceful cancel.

## Implementation order (high-level — actual plan to be produced by `writing-plans`)

1. `ModeController` + tab visibility wiring; toggle UI in Master List; remove Change File button.
2. `TASK_REGISTRY` skeleton + `TaskCard` + `WizardTab` shell (no Recent Tasks yet).
3. `TaskTab` + Settings/Status/Output page scaffolding (placeholder content).
4. `file_picker` + pre-Settings file dialog + default-folder resolution.
5. Runner shims around each of the four existing agents.
6. `WizardStatePersistence` + open-tab snapshot/restore + case-switch wiring.
7. Recent Tasks history + Reopen flow.
8. Output Page editor (mammoth integration + save round-trip) + action buttons.
9. Manual end-to-end verification per task.
