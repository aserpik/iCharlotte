# Document Text Library — Design Spec

**Date:** 2026-06-01
**Status:** Approved (design) — pending implementation plan
**Branch:** `feature/document-text-library`

## Problem

When iCharlotte runs a task on a case document (summarize discovery, summarize
deposition, med chron, etc.), it extracts the document's text, uses it, and
discards it. Any extraction caching that exists today is transient — e.g.
`icharlotte_core/deposition/session_manager.py` caches transcript text to
`logs/depo_sessions/<hash>.txt` keyed by input path, then deletes it on cleanup.

As a result, when the user wants to ask the Chat tab a question about a document
that was *already* processed (e.g. "what did the plaintiff say in their
deposition?"), they must re-upload the file and have the LLM extract the text all
over again — slow and redundant, especially for large transcripts and OCR'd PDFs.

## Goal

A persistent, per-case **Document Text Library**: tasks deposit extracted source
text into it as a side effect, the user can also add documents to it manually, and
the Chat tab exposes the library as a checkbox list to query against — feeding the
already-extracted raw text to the LLM with zero re-extraction at query time.

## Requirements (decided during brainstorming)

1. **Stored/queried content:** full **raw extracted source text** (not the task's
   summary). The Chat tab feeds raw text to the LLM. The design must handle large
   documents so selecting several big transcripts doesn't silently overflow the
   model's context window.
2. **Granularity:** one entry **per task run**, **expandable** to its individual
   member files (a task may process several files; an entry bundles them).
3. **Labeling:** **auto-generated, user-friendly, list-friendly** labels (e.g.
   "Plaintiff's Deposition Transcript"), **inline-renameable**, with reset-to-auto.
4. **Population:** **both** — tasks auto-populate as a side effect, **and** a manual
   "Add to Library" path (covers past documents whose text wasn't kept, and
   task-less docs like a traffic collision report).

## Chosen approach

**Centralized `DocumentLibrary` service + post-task hook** (Approach 1 of 3).

A single new module owns extraction + storage. The 40+ `Scripts/` agents stay
untouched (low regression risk — the agent/subprocess wiring is brittle, per
prior incident notes). Content-hash dedup makes extraction a once-ever cost per
unique document rather than the per-query cost we have today.

Rejected alternatives:
- **Shared extract-once cache wired through every agent** — most elegant but
  touches all agents and the subprocess boundary; far too much blast radius.
- **Promote existing transient session caches** — only depo/med-chron cache text
  today; inconsistent coverage, bespoke per-task formats; collapses into Approach 1
  anyway.

Optional later optimization: let depo/med-chron hand their already-cached text to
the library to skip the second extraction. Not required for v1.

## Section 1 — Data model & storage layout

New module: `icharlotte_core/doc_library/`. Storage **travels with the case** under
the existing `.icharlotte` state folder, alongside `wizard_state.json` — so the
library moves/shares with the case files and is visible to every user on the
network drive. Resolved via the same convention as `WizardStatePersistence`
(`<case_root>/NOTES/AI OUTPUT/.icharlotte`, see
`icharlotte_core/ui/wizard/persistence.py`).

```
<case_root>/NOTES/AI OUTPUT/.icharlotte/doc_library/
  index.json            # the catalog of entries for this case
  blobs/{sha1}.txt      # extracted text, one file per UNIQUE source document
```

The library is keyed by `case_root` (the case folder path), which `ChatTab`
already has via `main_win.case_path`. Because it lives on the case (Z:) drive,
querying the library requires the drive to be available — acceptable, since the
case's source documents live there anyway.

`index.json`:

```jsonc
{
  "version": 1,
  "entries": [
    {
      "id": "5f3c…",                                  // uuid
      "label": "Plaintiff's Deposition Transcript",   // shown in chat; renameable
      "auto_label": "Plaintiff's Deposition Transcript", // kept so reset works
      "task_type": "summarize_depositions",           // or "manual"
      "created_at": "2026-06-01T14:22:00",
      "members": [
        {
          "source_path": "Z:\\…\\Buchalter Depo Vol 1.pdf",
          "source_name": "Buchalter Depo Vol 1.pdf",
          "blob": "a1b2c3….txt",          // content-hash → file in blobs/
          "char_count": 482133,
          "est_tokens": 120533,
          "extract_method": "pdf_fitz",
          "error": null                    // or a short failure string
        }
      ]
    }
  ]
}
```

- **Blobs content-hashed (SHA-1 of file bytes):** two entries referencing the same
  physical file share one blob — extracted once, ever. Deletion is **ref-counted**:
  a blob is removed only when no entry still references it.
- Each member stores `char_count` / `est_tokens` so the Chat side can budget
  context without re-reading the text.
- The **entry** is the expandable unit: the row shows `label`; expanding lists
  `members`.
- Per-case isolation is inherent — each case folder has its own `.icharlotte/doc_library/`.
- **Concurrency (shared drive):** atomic tmp-+-`os.replace` writes prevent a
  partial/corrupt `index.json`. Two users adding to the *same* case's library at the
  same instant is rare; v1 accepts last-writer-wins on the index (a re-add recovers
  any lost entry). Blobs are content-addressed, so concurrent blob writes are
  idempotent.

## Section 2 — Population flow & dedup

One entry point, two callers:

```python
DocumentLibrary(case_root).add_entry(
    task_type,        # "summarize_depositions" | "manual" | …
    source_paths,     # list of files in this entry
    metadata,         # role/party/date hints for labeling (optional)
)
```

For each source file:
1. Compute SHA-1 of the file bytes.
2. If a blob with that hash exists → reuse it (no extraction).
3. Otherwise extract via a small `extract_any(path)` dispatcher reusing existing
   extractors — `document_processor.extract_text()` for PDFs (incl. OCR),
   `extract_docx_text()` for `.docx`, the Word-COM `.doc` path for legacy `.doc` —
   and write `blobs/{sha1}.txt`.
4. Record the member (`char_count`, `est_tokens`, `extract_method`).

Then write one entry referencing those blobs and return it.

**Caller A — task completion (automatic).** In the wizard task's
completion/success handler (where the input file list, `task_type`, and
settings-page metadata are in hand), call `add_entry(...)`. Fires only on success,
and only for **opted-in task types** — a small allow-list so non-source-doc tasks
don't create noise. *The exact handler per task is confirmed during implementation
planning.*

**Opt-in task IDs** (from `icharlotte_core/ui/wizard/registry.py`):

| Auto-populate | Task ID | Why |
|:---:|---|---|
| ✅ | `summarize_documents` | summarizes arbitrary case documents |
| ✅ | `summarize_discovery` | discovery responses |
| ✅ | `summarize_depositions` | deposition transcripts |
| ✅ | `medical_records` | medical records |
| ✅ | `med_chron_analysis` | medical records / chronology source |
| ✅ | `med_record_extractor` | medical records |
| ❌ | `depo_prep` | prep workproduct, not a captured source doc |
| ❌ | `subpoena_tracker` | tracking, no source-text to query |
| ❌ | `respond_to_discovery` | generates responses, not source capture |
| ❌ | `oppose_motion` | motion workproduct |
| ❌ | `chat` | the chat tab itself |

The non-opted tasks remain available via **manual add** if their inputs are ever
worth querying.

**Caller B — manual add.** A button on the Chat "Saved Documents" panel opens a
file picker → `add_entry("manual", picked_files, …)`.

**Threading.** Extraction (esp. OCR) can be slow → both callers run `add_entry` on
a background `QThread` worker with a small progress indicator. Index writes are
atomic (tmp + `os.replace`, the `session_manager` pattern). UI list refreshes on
completion.

**Idempotency.** Re-running the same task on the same files reuses blobs (dedup)
and **replaces** the prior entry for that `(task_type, file-set)` rather than
piling up duplicates.

## Section 3 — Auto-labeling

| Task | Auto-label pattern | Example |
|------|-------------------|---------|
| `summarize_depositions` | `{Party}'s Deposition Transcript` | "Plaintiff's Deposition Transcript" |
| `summarize_discovery` | `{Party}'s Discovery Responses` | "Defendant's Discovery Responses" |
| `medical_records` / `med_chron_analysis` / `med_record_extractor` | `Medical Records — {name}` | "Medical Records — Brier Buchalter" |
| `summarize_documents` | cleaned filename, title-cased | "Traffic Collision Report" |
| `manual` | cleaned filename, title-cased | "Traffic Collision Report" |

`summarize_documents` processes arbitrary documents with no inherent role/party, so
its auto-label falls back to the cleaned source filename (same as `manual`). When
an entry bundles several files, the label uses the first file's cleaned name with a
"+N more" suffix (e.g. "Traffic Collision Report +2 more") — renameable as always.

- **`{Party}` / `{name}`** come from metadata the task already collects (deponent
  name on the depo settings page; propounding/responding party in discovery). When
  role is unknown, fall back gracefully — drop the possessive and use the bare noun
  phrase ("Deposition Transcript"), or cleaned filename for manual adds. **Never**
  surface a raw cryptic filename as the label.
- **Collisions:** if a new entry's auto-label matches an existing one, append a
  disambiguator (date or volume): "Plaintiff's Deposition Transcript (Vol. 2)" /
  "(2026-04-30)".
- **Renaming:** entries are inline-editable; store user `label` separately from
  `auto_label` so right-click "Reset to auto name" restores the generated one.
  Inline edit uses the native `ItemIsEditable` flag — no `setItemWidget` (drag
  gotcha).

## Section 4 — Chat UI & query flow

**UI.** A new collapsible **"Saved Documents"** group in the Chat settings sidebar,
below the existing attached-files list — a `QTreeWidget` populated from
`DocumentLibrary(case_root).list_entries()`:
- Top level = entries (renameable label) with a checkbox.
- Expand → member files, each independently checkable. Checking an entry checks
  all members (tri-state).
- Toolbar: **Add to Library…**, **All / None**, **Refresh**, and a live
  **"Selected: N docs · ~T tokens"** readout.

**Query flow.** Rides the existing pipeline unchanged. Today `read_files_content()`
builds `"--- FILE: {name} ---\n{content}"` blocks from checked uploads. Extend it:
for each checked library member, **load its cached blob from disk** (no extraction)
and append the same framed block. Library text + ad-hoc uploads merge into the one
`file_content` string already flowing to `LLMHandler.generate()`. A checked
document reaches the model exactly as if uploaded and extracted — instantly.

**Context-budget handling** (raw text → large transcripts):
- "Selected: ~T tokens" uses stored `est_tokens` and feeds the existing
  `ContextIndicator`, so load is visible *before* sending.
- On send, if selected text + message + history would exceed the active model's
  context window, a **pre-send warning** names the large selections and lets the
  user proceed or deselect. No silent truncation; no automatic summary
  substitution (per the raw-text choice).

**Selection memory.** Checked entries persist per case (stored in the chat settings
block, like `attached_files`), so reopening the case keeps the last selection.

## Section 5 — Error handling & testing

**Error handling (library add is best-effort — never blocks a task's own output):**
- **Extraction failure** (corrupt PDF, OCR fails, COM error): member recorded with
  an `error` flag, no blob; entry still saves, file shows a warning marker. The
  triggering task is unaffected.
- **`.doc` via Word COM:** attach read-only, never `Quit`, never set `Visible`;
  wrap COM section in the `gc.disable()` guard (cyclic-GC crash note) with
  per-thread COM init.
- **Blob missing at query time:** that member is skipped with an inline notice;
  other selections still send.
- **Index corruption:** atomic writes prevent partial files; an unreadable index is
  backed up and re-initialized rather than crashing the Chat tab.
- **Offline / unavailable case drive:** the library lives on the case (Z:) drive,
  so if the drive is unreachable the "Saved Documents" panel shows an "unavailable —
  reconnect the case drive" state rather than erroring. (The case's source files are
  on the same drive, so the case isn't workable while it's offline regardless.)
  Add/re-extract of an unreachable file errors gracefully.
- **Snapshot semantics:** a blob is keyed to the file's bytes at capture time. If
  the source later changes, the old capture stays valid; re-adding produces a new
  hash/blob. (Optional: flag "source modified since capture" — not required v1.)

**Testing:**
- **Unit (no Qt, no LLM):** dedup (same file twice → one blob), ref-counted blob
  deletion, atomic index write/recovery, label generation per task type + collision
  disambiguation, rename/reset-to-auto, token estimate via existing
  `token_counter`, budget over/under threshold logic.
- **`extract_any` dispatcher:** small PDF/`.docx`/`.txt` fixtures per branch.
- **UI (`pytest.importorskip("pytestqt")`):** tree populates from a fake library,
  checking an entry checks its members, selection persists across reload, and
  checked library text appears in assembled `read_files_content()` output.

## Out of scope (v1)

- Automatic summary substitution when raw text overflows (user chose raw text).
- Semantic search / embeddings over the library (this is exact-text inclusion, not
  retrieval).
- Backfilling past task runs automatically (manual add covers the need).
- Locking/merge for simultaneous multi-user writes to the same case (last-writer-
  wins on the index is accepted for v1; blobs are content-addressed and safe).
- "Source modified since capture" detection (noted as optional).

## Key integration points (existing code)

- `icharlotte_core/document_processor.py` — `extract_text()` (PDF+OCR, line ~135),
  `extract_docx_text()` (~line 1019).
- `icharlotte_core/ui/tabs.py` — `ChatTab`: `read_files_content()` (~1235),
  settings sidebar file list (~343), `load_case()` (~507), `get_attachment_info()`.
- `icharlotte_core/chat/persistence.py` — per-case JSON pattern; **selection
  memory** (last-checked library entries) is stored in the chat settings block.
- `icharlotte_core/ui/wizard/persistence.py` — `.icharlotte` folder convention
  (`<case_root>/NOTES/AI OUTPUT/.icharlotte`) and atomic-write pattern the library
  reuses; library nests under it as `doc_library/`.
- `icharlotte_core/chat/token_counter.py` — token estimation.
- Wizard task completion handlers in `icharlotte_core/ui/wizard/` (per-task hook
  sites — to be enumerated in the implementation plan).
- Atomic-write pattern: `icharlotte_core/deposition/session_manager.py`.
