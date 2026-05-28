# Depo Prep — Wizard Task Design

**Date:** 2026-05-27
**Status:** Approved design; ready for implementation plan
**Audience:** Future implementer (Claude or human) of the Depo Prep wizard task

---

## Goal

Add a new "Depo Prep" task to iCharlotte's wizard mode that generates a deposition-prep outline (topic-organized questions grounded in case sources) for a lawyer-specified deponent. The output is intended to be used directly by the lawyer at the deposition — printed, annotated, and worked from line by line.

## Non-goals

- Real-time use during a live deposition (no voice/transcript integration).
- Cross-case templates or persistent question banks.
- Automatic exhibit preparation (creating exhibit binders, marking copies, etc.).
- Replacing the lawyer's judgment — the agent surfaces material; the lawyer curates.

---

## High-level architecture

Two-phase subprocess wizard task, matching the existing `summarize_depositions` / `med_chron_analysis` pattern (subprocess-driven, `AWAITING_INPUT:` handoff). **No speculative Phase 1** — Phase 1 fires only when the lawyer clicks "Analyze Sources" after filling in all settings (deponent, sources, strategy notes, style). This avoids wasted LLM calls and ensures the topic list reflects the lawyer's strategy.

```
Lawyer fills Settings page (deponent, sources, style, strategy notes, per-topic flags)
        │
        ▼ click "Analyze Sources"
Phase 1 subprocess: depo_prep.py --phase=analyze
  Stage 1.1: Extract text from each source file
  Stage 1.2: Per-source structured digest (parallel LLM, max 4 concurrent, Flash)
  Stage 1.3: Topic clustering (single LLM call, Pro)
  Stage 1.4: Persist session + AWAITING_INPUT
        │
        ▼ wizard reveals editable topic list
Lawyer reviews / edits / reorders / adds topics
        │
        ▼ click "Generate Outline"
Phase 2 subprocess: depo_prep.py --phase=generate
  Stage A: Per-topic question generation (parallel LLM fan-out, max 4 concurrent, Pro)
  Stage B: Cross-topic dedup + coverage check (single LLM call)
  Stage C: Polish — phrasing/transitions only, no substantive changes (single LLM call)
  Stage D: Render outline.docx + outline.md
        │
        ▼
Custom DepoPrepOutputPage with collapsible markdown view + standard wizard actions
```

### Why this shape

- **Two phases with a curation step in between** lets the lawyer steer the outline before the heavy-cost per-topic generation runs. They can drop irrelevant topics, add missed ones, and edit strategic notes — all before the per-topic LLM fan-out fires.
- **Per-topic fan-out in Phase 2** (vs. one big call) keeps each topic's generation focused on its own facts/quotes. With 5+ source documents, a single mega-prompt would hit output-token caps and shortchange later topics.
- **Source digest as the Phase 1 → Phase 2 bridge** means Phase 2 never re-reads raw source files. Each topic's per-topic call gets exactly its `relevant_digest_refs` entries, keeping prompts focused and reproducible.

---

## File layout

```
icharlotte_core/ui/wizard/
  registry.py                                    (modified)
  task_tab.py                                    (modified — _output_page_cls_factory lookup)
  pages/
    depo_prep_settings_page.py                   (new)
    depo_prep_output_page.py                     (new)

Scripts/
  depo_prep.py                                   (new — thin CLI orchestrator)
  depo_prep_lib/                                 (new)
    __init__.py
    source_digest.py                             # Phase 1: per-source extraction
    topics.py                                    # Phase 1: topic clustering
    questions.py                                 # Phase 2 Stage A
    merge.py                                     # Phase 2 Stage B
    polish.py                                    # Phase 2 Stage C
    render_docx.py                               # Phase 2 Stage D — .docx
    render_md.py                                 # Phase 2 Stage D — .md
    prompts.py                                   # All prompt templates (style-aware)
    schemas.py                                   # JSON schemas + dataclass definitions

config/
  llm_preferences.json                           (modified — add DepoPrep agent)

tests/test_wizard/
  test_depo_prep_settings_page.py                (new)
  test_depo_prep_phase1.py                       (new)
  test_depo_prep_phase2.py                       (new)
  test_depo_prep_render.py                       (new)
  test_depo_prep_integration.py                  (new)
```

---

## Settings page

### UI layout

```
┌─ Deponent ────────────────────────────────────────────────────────┐
│ Name:  [ Jane Doe                                    ▼ ]          │
│        ↑ dropdown from MasterCaseDatabase parties for the active  │
│          case, with free-text fallback                            │
│ Role:  [ Plaintiff's treating orthopedist                ]        │
├─ Source files ────────────────────────────────────────────────────┤
│ Deponent's own materials                                          │
│   [+ Add files]   [list with remove buttons]                      │
│ Case context                                                      │
│   [+ Add files]   [list with remove buttons]                      │
├─ Instructions ────────────────────────────────────────────────────┤
│ Style:  ( ) Discovery / Fact-gathering                            │
│         ( ) Lock-down (leading admissions for MSJ/trial)          │
│         ( ) Expert challenge (Daubert-style)                      │
│         ( ) Friendly (own client prep)                            │
│                                                                   │
│ Free-text strategy notes:                                         │
│ ┌──────────────────────────────────────────────────────────────┐ │
│ │ (Multi-line QPlainTextEdit for case theory, topics to        │ │
│ │  emphasize, things to avoid, key admissions to extract…)     │ │
│ └──────────────────────────────────────────────────────────────┘ │
├─ Per-topic content (what should appear under each topic) ─────────┤
│ ☑ Strategic note ("why this topic")            [checked by default]│
│ ☑ Key source facts (with citations)            [checked by default]│
│ ☐ Impeachment hooks / inconsistencies                             │
│ ☐ Anticipated objections + workaround phrasings                   │
├───────────────────────────────────────────────────────────────────┤
│ [ Analyze Sources ]                                               │
│                                                                   │
│ ─── after Phase 1 completes, topic editor appears here ──────     │
└───────────────────────────────────────────────────────────────────┘
```

### Behaviors

- **"Analyze Sources"** disabled unless: deponent name non-empty, at least one source file selected, style selected, at least one per-topic content flag checked.
- Clicking **"Analyze Sources"** writes `config.json` to the session folder, then launches `depo_prep.py --phase=analyze` as a subprocess. Status line / progress bar inline below the button.
- On `AWAITING_INPUT:` from the subprocess, the **topic editor** appears below the button.
- **"Re-analyze"** button (appears after first Phase 1 completes) re-runs Phase 1 with current settings; per-source digests are cached by file hash (only Stage 1.3 — topic clustering — re-runs unless source files changed).

### Deponent dropdown

Pulls party names from `MasterCaseDatabase.get_parties(case_number)` for the active case. The dropdown is editable (QComboBox with `setEditable(True)`); if the lawyer types a name that doesn't match a party, it's used as-is. The "Role" field is always free-text — never auto-populated.

---

## Phase 1 pipeline

CLI:
```
python Scripts/depo_prep.py --phase=analyze \
    --session=<session_path> \
    --config=<config_json_path>
```

### Stage 1.1 — Source ingestion

For each file in `config.json`'s source lists:
- PDFs: `document_processor.extract_text_from_pdf` (handles OCR fallback already).
- DOCX: `document_processor.extract_docx_text` (walks `doc.element.body` — handles tables, headers, nested content; see cross-cutting memory `MEMORY.md` note on python-docx silently skipping tables).
- TXT: read directly.

Each source's extracted text is cached on disk in `digests/raw/<filename>.txt` so re-runs don't re-extract.

### Stage 1.2 — Per-source digest

One LLM call per source, fanned out with `concurrent.futures.ThreadPoolExecutor(max_workers=4)`. Uses the `extraction` task config (Gemini 2.5 Flash).

Output schema (one JSON file per source in `digests/<filename>.json`):

```json
{
  "source_id": "med_records_2024-08-15.pdf",
  "source_kind": "medical_records | deposition_transcript | discovery_response | pleading | other",
  "deponent_statements": [
    {
      "text": "I have never had back pain before this accident.",
      "location": "p.47:18-22",
      "context": "Direct examination by plaintiff's counsel"
    }
  ],
  "factual_anchors": [
    {
      "fact": "MRI on 2024-09-12 showed L4-L5 disc protrusion 4mm",
      "location": "p.12 (Bates DEF-00154)",
      "topic_tags": ["injury", "imaging", "causation"]
    }
  ],
  "inconsistencies": [
    {
      "claim_a": "Discovery response #7: pain began immediately after collision",
      "claim_a_source": "this file, RFA #7",
      "claim_b": "ER triage note: 'no acute pain, will see PCP next week'",
      "claim_b_source": "med_records.pdf p.3",
      "topic_tags": ["injury_onset", "credibility"]
    }
  ],
  "summary": "Brief 2-3 sentence summary of what this source contributes."
}
```

`location` strings are passed through verbatim to Phase 2 so the rendered outline can cite source page-lines / Bates numbers without re-deriving them.

Caching: keyed by SHA-256 of the source file. If the file hash matches an existing `digests/<filename>.json`, the call is skipped.

### Stage 1.3 — Topic clustering

Single LLM call using the `general` task config (Gemini 2.5 Pro). Input: all per-source digests + deponent name/role + style + free-text strategy notes.

Output schema (`topics.json`):

```json
{
  "topics": [
    {
      "id": "t01",
      "title": "Pre-existing back conditions",
      "strategic_note": "Establish that plaintiff had documented chronic low back pain dating to 2019, contradicting their RFA #7 sworn denial.",
      "relevant_digest_refs": [
        "med_records.pdf#factual_anchors[2]",
        "med_records.pdf#factual_anchors[6]",
        "rfa_responses.docx#deponent_statements[0]"
      ],
      "default_checked": true
    }
  ]
}
```

The LLM is instructed to produce 8–15 topics. `relevant_digest_refs` use the format `<source_id>#<schema_field>[<index>]` so Phase 2 can resolve them by JSON-path lookup. Custom topics added by the lawyer have empty `relevant_digest_refs` — Phase 2 sees the full digest list for those, flagged as `"lawyer_added"`.

**Robustness:** Accept whatever count the LLM returns. If <3 topics, surface a warning in the topic editor ("Source material thin — consider adding more sources or strategy detail."). If >20 topics, truncate to 20 and log a warning. Never hard-fail on topic count.

### Stage 1.4 — Persist + AWAITING_INPUT

Subprocess writes:
- `session.json` — settings snapshot + paths to all digests
- `topics.json` — output of Stage 1.3
- `trace.log` — LLM call log (for debugging)

Then emits `AWAITING_INPUT: <session_path>` on stdout and exits 0.

---

## Topic confirmation UI (between Phase 1 and Phase 2)

Once the wizard sees `AWAITING_INPUT:`, the Settings page reveals the topic editor:

```
─────── Topics found ───────
☑ 1. ▣ Pre-existing back conditions                    [✏] [🗑]
      Strategic: Establish documented chronic LBP since 2019,
                 contradicting RFA #7 denial.
☑ 2. ▣ Treatment timeline & gaps                       [✏] [🗑]
      Strategic: Highlight 8-month gap between PT and surgery
                 to support causation challenge.
☐ 3. ▣ Pre-accident athletic activity                  [✏] [🗑]
      Strategic: ...
...
[+ Add custom topic]

[ Generate Outline (8 topics selected) ]
```

Behaviors:
- **Drag-handle (▣)** on each row → drag to reorder. Implemented via `QListWidget.InternalMove`.
- **Edit (✏)** opens an inline editor (or a small modal) for title + strategic note. Both editable.
- **Custom topic** opens the same editor; lawyer types title + strategic note. Marked `lawyer_added: true` in `topics.json`.
- **Generate Outline** enabled when ≥1 topic checked.
- Topic list state persists to `topics.json` on every edit. If the user closes and reopens the case, the wizard restores the same topic editor state.

### Important UI gotcha

Per `MEMORY.md` note `qlistwidget_setitemwidget_drag.md`: do NOT use `setItemWidget` on the topic rows — it breaks drag-reorder. Use native `Qt.ItemIsUserCheckable | Qt.ItemIsEditable | Qt.ItemIsDragEnabled` flags. If the strategic-note text needs to render on two lines, do it via a custom delegate (`QStyledItemDelegate.paint`), not embedded widgets.

---

## Phase 2 pipeline

CLI:
```
python Scripts/depo_prep.py --phase=generate --session=<session_path>
```

Reads `session.json` + `topics.json` + `digests/*.json` from the session folder.

### Stage A — Per-topic question generation

One LLM call per **selected** topic, fanned out with `ThreadPoolExecutor(max_workers=4)`. Uses `general` task config (Gemini 2.5 Pro).

Input to each call:
- The topic (title, strategic_note, `relevant_digest_refs`)
- The resolved digest entries (read from `digests/*.json` by the orchestrator and passed inline)
- Deponent name + role
- Style — swaps in style-specific instructions from `prompts.py`
- Per-topic content flags — which sub-sections to include
- Free-text strategy notes (truncated to 2000 chars if longer)

Output:
```json
{
  "topic_id": "t01",
  "questions": [
    {
      "n": 1,
      "text": "Before the August 15, 2024 collision, had you ever sought medical treatment for back pain?",
      "purpose": "Establish baseline for impeachment by 2019 PT records.",
      "source_facts": [
        "Plaintiff's RFA #7 denied any prior back pain (rfa_responses.docx)",
        "2019-03 PT intake form documents 'chronic low back pain 2 years' (med_records p.12)"
      ],
      "impeachment_hook": "If denies: confront with rfa_responses.docx RFA #7, then med_records p.12",
      "objection_alts": null
    }
  ]
}
```

**Conditional sub-fields:** The prompt template only includes instructions for `purpose`, `source_facts`, `impeachment_hook`, and `objection_alts` when the corresponding per-topic content flag is checked. When a flag is off, the field is omitted entirely from the prompt — the model never sees the instruction, so it can't generate the field. Saves output tokens and prevents unrequested content.

**Failure handling:** On any single-topic LLM failure, the topic is marked `{"topic_id": "...", "error": "<message>", "questions": []}`. The orchestrator logs it and proceeds. Partial output is rendered with an "AI errors encountered" callout at the top.

### Stage B — Cross-topic dedup + coverage

Single LLM call. Input: all topic outputs from Stage A + the full digest summary list + the curated topic list.

Output:
```json
{
  "duplicates": [
    {"keep": "t02.q5", "drop": "t05.q3", "reason": "Same question phrased differently"}
  ],
  "coverage_gaps": [
    "No question addresses the 2019 chiropractor visits in med_records p.8-9 — consider adding to Topic 1."
  ],
  "renumber_after_dedup": true
}
```

The orchestrator applies the drops, renumbers within each topic, and stashes `coverage_gaps` for rendering.

### Stage C — Polish

Single LLM call. Input: full assembled outline post-dedup. Output: same structure, with transitions normalized and redundant phrasing tightened.

**The polish prompt explicitly forbids substantive changes:** no new facts, no new questions, no dropped questions, no changed strategic notes. Only phrasing tweaks and transitions. This protects the lawyer's curated topic list from being silently rewritten in the final stage.

### Stage D — Render

**`outline.docx`** via `python-docx`:

```
Depo Prep Outline — Jane Doe                        ← Heading 1
Plaintiff's treating orthopedist                      ← subtitle (italic)
Case: Smith v. Jones, Case No. 24CV01234              ← case metadata line
Prepared: May 27, 2026                                ← date line

──────────────────────────────────────────────────────  ← horizontal rule

1. PRE-EXISTING BACK CONDITIONS                       ← Heading 2

   Strategic: Establish documented chronic LBP since   ← Italic, indented 0.25"
              2019, contradicting RFA #7 denial.

   1.1  Before the August 15, 2024 collision, had     ← Numbered, body indent
        you ever sought medical treatment for back
        pain?

        Purpose: Establish baseline for impeachment.   ← Italic small, hanging indent
        Source facts:                                  ← Italic small
          • RFA #7 denied prior back pain              ← Bullet, deeper indent
          • 2019-03 PT intake: "chronic LBP 2 years"
        Impeachment: If denies → confront with RFA
                     #7, then med_records p.12

   1.2  ...

────────────────────────────────────────────────────
AI Coverage Notes
- No question addresses the 2019 chiropractor visits in med_records p.8-9.
- Consider adding a topic on social media activity post-accident.
```

**Word formatting rules** (per `MEMORY.md` cross-cutting lessons):
- Spacing between paragraphs uses `space_after` paragraph property — never empty paragraphs.
- Never directly manipulate paragraph marks (`\r`).
- After rendering, run `word_validator.validate_after_edit` and emit any errors to the trace log.

**`outline.md`** — same content as markdown, generated alongside the docx in a parallel pass. Renders identically structured (H1 / H2 / numbered list / italic sub-notes / bullet lists).

### Output paths

Files land in:
```
<case_root>/NOTES/AI Output/Depo Prep — <Deponent Name> — <YYYY-MM-DD HH:MM>/
  session.json
  topics.json
  digests/
    raw/<source>.txt
    <source>.json    (one per source)
  outline.docx
  outline.md
  trace.log
```

The session folder name includes the deponent name so multiple deponents per case produce distinct folders.

---

## Output page (custom)

`DepoPrepOutputPage` extends the base `OutputPage`. The standard header (with `Open in Word` / `Save As…` / `Rerun` / `Edit Settings` buttons) is preserved; the central preview area is replaced with a custom widget.

### Layout

```
┌─ Header bar ──────────────────────────────────────────┐
│ Depo Prep — Jane Doe                                  │
│ [📄 Open in Word] [💾 Save As…] [🔄 Rerun] [✏ Settings]│
├─ Outline (collapsible) ───────────────────────────────┤
│ ▼ 1. Pre-existing back conditions                     │
│     Strategic: …                                      │
│     ▼ 1.1 Before the August 15, 2024 collision…       │
│         Purpose: …                                    │
│         Source facts:                                 │
│           • RFA #7 denied prior back pain   [📋 copy] │
│         Impeachment: …                                │
│         [📋 Copy question] [🔍 Jump to source PDF]    │
│     ▶ 1.2 …                                           │
│ ▶ 2. Treatment timeline & gaps                        │
│ ▶ 3. ...                                              │
├───── Coverage notes from the AI ──────────────────────┤
│ • No question addresses the 2019 chiropractor visits…│
└────────────────────────────────────────────────────────┘
```

### Implementation

- Renders `outline.md` via `QTextBrowser` with markdown → HTML conversion + a small JS layer for collapse/expand toggles, OR uses a `QTreeWidget` populated from a parse of `outline.md`. The latter is more robust for collapse state across re-renders.
- **Copy question** button → puts just the question text on the clipboard via `QApplication.clipboard().setText()`.
- **Jump to source PDF** button → parses the cited location (e.g., `"med_records.pdf p.12"`) and opens the file via the existing `bridge.py` `local-resource://` scheme with a `#page=12` anchor for pdf.js. Falls back to opening the file at page 1 if the page reference can't be parsed.
- Coverage notes section appears at the bottom; collapsible.

### Registry change

`TaskSpec` gains an optional `_output_page_cls_factory` field. If set, `task_tab.py` uses it when constructing the OutputPage; otherwise falls back to the default. One-line registry change, one-line lookup change.

```python
"depo_prep": TaskSpec(
    task_id="depo_prep",
    title="Depo Prep",
    description="Generate a deposition outline with questions grounded in case sources.",
    icon_glyph="❓",
    script_name="depo_prep.py",
    default_folders=["DISCOVERY", "PLEADINGS", "RECORDS"],
    phase1_args=["--phase=analyze"],
    phase2_flag="--phase=generate",
    _settings_page_cls_factory=_depo_prep_settings_page_cls,
    _output_page_cls_factory=_depo_prep_output_page_cls,
)
```

---

## LLM configuration

New agent block in `config/llm_preferences.json`:

```json
{
  "agent_id": "DepoPrep",
  "task_configs": {
    "general": {
      "primary_model": "gemini-2.5-pro",
      "fallback_sequence": ["gemini-2.5-pro", "claude-opus-4-7", "gpt-4o"]
    },
    "extraction": {
      "primary_model": "gemini-2.5-flash",
      "fallback_sequence": ["gemini-2.5-flash", "claude-haiku-4-5"]
    }
  }
}
```

Stage usage:
- **Stage 1.2 (per-source digest):** `extraction` — Flash is fast and accurate for structured extraction.
- **Stage 1.3 (topic clustering):** `general` — quality and reasoning matter.
- **Stage A (per-topic question gen):** `general` — quality is the whole point.
- **Stage B (dedup + coverage):** `general`.
- **Stage C (polish):** `general`.

All calls use `LLMCaller` for automatic model fallback. Trace log captures: stage name, topic ID (if applicable), prompt token count, output token count, model used, latency.

---

## Testing

```
tests/test_wizard/
  test_depo_prep_settings_page.py
    # UI smoke: party dropdown population, file picker add/remove,
    # button enable/disable logic (deponent + sources + style + ≥1 flag)
  test_depo_prep_phase1.py
    # source_digest schema validation (mocked LLM)
    # topics output validation (mocked LLM)
    # digest caching by file hash
    # custom topic handling (empty relevant_digest_refs)
  test_depo_prep_phase2.py
    # per-topic generation prompt — fields conditionally included based on flags
    # dedup application + renumbering
    # polish stage rejects added/dropped questions (defense-in-depth check)
    # partial failure: one topic errors, others render
  test_depo_prep_render.py
    # .docx structure (validate_after_edit passes)
    # .md structure matches .docx structure
    # spacing uses space_after, never empty paragraphs
    # coverage notes section appears when populated, omitted when empty
  test_depo_prep_integration.py
    # end-to-end with stub LLM, fixture sources (1 depo + 1 discovery + 1 med record)
    # session folder structure matches spec
    # config.json round-trip through Phase 1 and Phase 2
```

All LLM calls in unit tests are mocked at the `LLMCaller.call` boundary. The integration test uses a `StubLLMCaller` that returns canned responses keyed by prompt fingerprint.

---

## Open questions / explicit deferrals

- **Exhibit prep:** Not in scope. A future enhancement could let the lawyer mark which source-cited documents should be pulled into an exhibit binder; deferred.
- **Per-topic regenerate:** The output page could grow a "regenerate just this topic" button (re-runs Stage A for one topic with optional refinement notes). Deferred until the base feature is in use; UX should be informed by how lawyers actually iterate.
- **Persistent question banks:** Not in scope. Could be a future global "boilerplate questions per witness type" library.
- **Voice/real-time depo integration:** Out of scope.

---

## Cross-cutting concerns to follow

- **Word output validation:** Every `.docx` produced must be validated via `icharlotte_core/word_validator.py` per project CLAUDE.md. Use `validate_after_edit` since this is a python-docx generation, not a COM redline.
- **Subprocess encoding:** All subprocess calls that handle user/document text must use `encoding="utf-8", errors="replace"` (see `MEMORY.md` `windows_subprocess_text_encoding.md`).
- **Agent script sys.path:** `Scripts/depo_prep.py` must `sys.path.insert(0, project_root)` BEFORE importing from `icharlotte_core` (see `MEMORY.md` `agent_script_syspath.md`).
- **Worktree vs main checkout:** All edits must be applied to BOTH `C:\geminiterminal2\` (where iCharlotte runs) and the worktree (see `MEMORY.md` `worktree_vs_main_checkout.md`).
- **No hardcoded tab indices:** Use `_index_of_tab(name)` if any tab-switching is needed (see `MEMORY.md` `hardcoded_tab_indices.md`).
- **QListWidget topic editor:** Use native item flags, not `setItemWidget`, to preserve drag-reorder (see `MEMORY.md` `qlistwidget_setitemwidget_drag.md`).
- **Wizard restore:** If Phase 1 results exist when the case is reopened, restore the topic editor state from `topics.json`.

---

## Success criteria

A senior litigator preparing for a deposition can:
1. Open the case in iCharlotte, switch to wizard mode, select Depo Prep.
2. Pick the deponent from the case parties dropdown, add 3–6 source files, pick a style, type 1–3 paragraphs of strategy, check 2–3 per-topic content flags.
3. Click "Analyze Sources" and wait ~30–90 seconds for the topic list.
4. Review, edit, reorder, and add topics in the editor.
5. Click "Generate Outline" and wait ~2–4 minutes.
6. Get a `.docx` outline with topic-organized, source-grounded questions — usable directly at the deposition, with citations to source page-lines / Bates numbers for any factual claim made in the questions.
7. Use the interactive markdown view to copy individual questions or jump to source PDFs while reviewing the outline.

The quality bar is: the lawyer should not need to rewrite questions wholesale. They may delete a few and reorder a few, but the bulk of what the AI produces should be usable as-is.
