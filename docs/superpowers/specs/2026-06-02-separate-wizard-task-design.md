# Separate → Wizard Mode Task — Design

**Date:** 2026-06-02
**Status:** Approved (brainstorming complete)

## Goal

Port the advanced-mode "Separate" document-splitting workbench to Wizard Mode as a
first-class task, with **full parity** to the advanced-mode experience.

## Decisions (locked during brainstorming)

1. **UI scope:** Full parity — editable table (Sep./Merge Group/Pages/Title),
   embedded PDF preview, Mark Start/End range marking, sensitivity slider +
   Re-analyze, search, add/delete rows, Process.
2. **Multi-PDF:** One PDF per task tab. A second PDF means opening another
   Separate task. (Advanced mode's multi-PDF left list is NOT ported.)
3. **Code reuse:** Extract IndexTab's single-PDF workbench into a reusable
   widget; both advanced mode and the wizard embed it. Single source of truth.
4. Title **"Separate Documents"**, icon 📑.
5. Recent-task reopen re-opens the file picker (analysis is ephemeral; no
   doc-map persistence into wizard state). YAGNI.

## Current state (reference)

- **`Scripts/separate.py`** — CLI agent. `--headless --sensitivity N <pdf>`:
  extracts page headers (PyMuPDF + Tesseract OCR fallback), sends them to the
  LLM (`agent_separate`) in 100-page chunks to identify distinct documents,
  writes a Word index to `NOTES/AI OUTPUT/INDEXES/Index_<base>.docx`
  (`create_index_word`), dumps a JSON doc map to a temp file, and prints
  `JSON_MAP: <path>` to stdout. Also has a legacy console `--interactive` mode
  (not used by the GUI).
- **`iCharlotte.run_separator_path(path, sensitivity)`** (iCharlotte.py:2145) —
  launches `separate.py --headless` via `AgentRunner`, parses `JSON_MAP:` from
  stdout, loads the doc map into `IndexTab.add_pdf(path, docs)`.
- **`IndexTab`** (icharlotte_core/ui/tabs.py:2338) — the advanced-mode workbench:
  left PDF list + middle editable table (Sep. checkbox, Merge Group, ID, Date,
  editable Pages, editable Title) + sensitivity slider/Re-analyze + search +
  Add/Delete rows + right-side `PdfViewerWidget` with Mark Start / Mark End /
  Clear + `process_documents()` which splits checked rows and merges grouped
  rows into `PULLED-<source>/`.
- **Wizard task patterns:**
  - `TaskTab` (Settings→Status→Output) for script-backed tasks.
  - `InProcessTaskTab` for QThread-backed tasks (subpoena, med extractor,
    respond-to-discovery): settings widget emits `run_requested(dict)`.
  - `OpposeMotionTaskTab(WizardTaskContainer)` — fully custom 3-page tab; the
    closest model for Separate's interactivity.
  - Dispatch: `iCharlotte._open_task_tab` (iCharlotte.py:1510+) looks up
    `get_in_process_task_builder_name(task_id)` and calls
    `builder(spec, case_path, file_number, parent)`; a `None` return aborts.

## Architecture

### A. Extract the shared workbench

New module **`icharlotte_core/ui/separator_workbench.py`**:

```
class SeparatorWorkbench(QWidget):
    reanalyze_requested = Signal(int)      # sensitivity 1..3
    processing_complete = Signal(dict)     # {"created": [...], "errors": [...], "output_folder": str}

    def load_docs(self, pdf_path: str, docs: list[dict]) -> None: ...
    def set_busy(self, busy: bool) -> None: ...   # disable Re-analyze/slider while analyzing
```

Owns everything currently right of IndexTab's PDF list:
- `doc_table` + `_add_doc_to_table`, search/filter, sensitivity slider +
  Re-analyze, Add/Delete rows, context-menu batch edit.
- `PdfViewerWidget` + Mark Start / Mark End / Clear.
- `process_documents()` (split + merge into `PULLED-<source>/`), refactored to
  emit `processing_complete` instead of popping a QMessageBox directly (host
  decides how to surface the summary).

**Decoupling:** today `on_reanalyze_clicked` calls
`self.window().run_separator_path(...)`. In the widget it emits
`reanalyze_requested(sensitivity)`; the host wires it.

### B. IndexTab becomes a thin host

`IndexTab` keeps its left PDF list + persistence (`{file_number}_index.json`,
`GEMINI_DATA_DIR`) and embeds one `SeparatorWorkbench`:
- `on_pdf_selected` → `workbench.load_docs(path, docs)`.
- `workbench.reanalyze_requested` → `self.window().run_separator_path(path, sensitivity)`.
- `add_pdf` continues to persist + select; selection drives `load_docs`.

Advanced mode must behave identically — re-verified post-refactor.

### C. Wizard task — `SeparateTaskTab`

New module **`icharlotte_core/ui/wizard/pages/separate_page.py`**:

```
class SeparateSettingsPage(QWidget):    # PDF label + sensitivity slider + "Analyze"
    analyze_requested = Signal(int)     # sensitivity

class SeparateAnalysisWorker(QThread):  # runs separate.py --headless
    progress = Signal(str)
    finished_analysis = Signal(bool, object)   # (success, docs | error_str)

class SeparateTaskTab(WizardTaskContainer):  # 3 pages, modeled on OpposeMotionTaskTab
    task_completed = Signal(dict)
    # PAGE_SETTINGS / PAGE_STATUS / PAGE_WORKBENCH

def build_separate_tab(spec, case_path, file_number, parent) -> SeparateTaskTab | None
```

- **Settings page:** PDF name, sensitivity slider (1=Broad, 2=Default, 3=Fine),
  *Analyze* primary button → `analyze_requested(sensitivity)`.
- **Status page:** reuse `StatusPage`; indeterminate bar; worker `progress` → status.
- **Workbench page:** embeds `SeparatorWorkbench`.
  - `reanalyze_requested` → re-run worker (back to Status, then re-load).
  - `processing_complete` → summary banner + *Open Folder* button.
- **Worker:** subprocess `python separate.py --headless --sensitivity N <pdf>`,
  `encoding="utf-8", errors="replace"` (Windows cp1252 gotcha). Parse
  `JSON_MAP:` from stdout, read the JSON doc map, emit `finished_analysis(True, docs)`.
  Reusing the subprocess path means the Word index is created for free, exactly
  as advanced mode does.

### D. Registry & routing

- **`registry.py`:** add
  ```
  "separate": TaskSpec(
      task_id="separate",
      title="Separate Documents",
      description="Split a combined PDF into individually-named documents using AI.",
      icon_glyph="\U0001F4D1",  # 📑
      script_name="separate.py",
      default_folders=[],
  )
  ```
- **`task_routing.py`:** add `"separate": "build_separate_tab"` to
  `_IN_PROCESS_TASK_BUILDERS`.
- **`in_process_task_tab.py`:** add a `build_separate_tab` shim that imports the
  real builder from `separate_page.py` (mirrors how `build_oppose_motion_tab` is
  re-exported), so `getattr(in_process_task_tab, name)` resolves.
- **`build_separate_tab`:** open a one-PDF `QFileDialog` (start dir via
  `resolve_default_folder(case_path, spec.default_folders)`), return the tab or
  `None` on cancel.

### E. Output locations — unchanged

- Word index → `NOTES/AI OUTPUT/INDEXES/Index_<base>.docx`.
- Split/merged PDFs → `PULLED-<source>/` next to the source PDF.

### F. Word validation

`separate.py:create_index_word` (shared with advanced mode) gains a lightweight
`word_validator` check on the generated `.docx`, per CLAUDE.md's mandatory
validation rule. Small, benefits both modes.

## Testing

- Registry entry + routing resolution for `separate`.
- `build_separate_tab` returns `None` when the picker is cancelled.
- `SeparateAnalysisWorker` parses `JSON_MAP:` and surfaces the doc list (mock the
  subprocess / feed canned stdout).
- `SeparatorWorkbench.load_docs` populates the table; `process_documents` splits
  and merges correctly against a tiny generated multi-page PDF; emits
  `processing_complete` with the created files.
- Signal wiring: `reanalyze_requested`, `analyze_requested`, `processing_complete`.
- **Regression:** advanced-mode `IndexTab` still loads, edits, and processes
  (existing tests + a manual app run).

## Known limitation

Recent-task reopen re-opens the file picker — the analysis (doc map) is not
persisted into wizard state. Doc-map persistence is deferred (YAGNI). The
existing reopen/restore special-casing in iCharlotte.py currently only handles
`build_oppose_motion_tab`; Separate's reopen path follows the generic builder
route (re-pick the PDF).

## Out of scope

- Multi-PDF list inside one wizard tab.
- Persisting the analyzed doc map across sessions.
- Changes to `separate.py`'s LLM prompt / sensitivity logic.
