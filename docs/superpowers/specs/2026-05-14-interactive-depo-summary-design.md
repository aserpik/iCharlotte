# Interactive Deposition Summary — Design

**Date:** 2026-05-14
**Status:** Approved for implementation planning
**Owner:** iCharlotte deposition workflow

## Goal

Replace the single-pass `summarize_deposition` agent with an interactive two-phase workflow. The agent proposes ranked testimony topics, pauses for user input via a popup, and then generates the summary using only the topics and rules the user selected.

## Motivation

Today the agent picks topics and drafts bullets in one LLM call, with no user input. The output frequently doesn't match what the attorney actually wants to highlight, and there's no way to steer phrasing, depth, or focus without rerunning from scratch. The new flow puts a human in the loop between topic discovery and bullet drafting.

## Out of scope (v1)

- Persisting user preferences across runs. Every popup resets to hardcoded defaults.
- Per-topic bullet counts. One global count applies to all selected topics.
- A "resume leftover sessions" UI on app startup. Orphaned session files sit on disk until manually cleaned.
- Editing the rank order in the popup. Topics are shown in agent-assigned rank order; user can only check/uncheck/rename/add.

## Removed features

- **Exhibits Referenced section** in the output docx.
- **Potential Impeachment Material section** in the output docx.
- The `ExhibitExtractor` class and `ImpeachmentDetector` class in `Scripts/summarize_deposition.py`.
- The parallel structured-extraction pass (`DEPOSITION_EXTRACTION_PROMPT.txt`) — was used to feed the cross-check. Cross-check now runs against the original transcript directly.

## Architecture

Two subprocess invocations of `Scripts/summarize_deposition.py`, coordinated through a sidecar JSON file.

```
Phase 1: python summarize_deposition.py --phase=topics <input>
   ↓ writes session JSON, prints "AWAITING_INPUT:<session_path>", exits 0
   ↓
UI shows READY button on the agent status row
   ↓ user clicks READY → DepoSummaryConfigDialog opens
   ↓ user submits → UI rewrites session JSON with user_config
   ↓
Phase 2: python summarize_deposition.py --phase=summary <session_path>
   ↓ reads session, generates summary, saves docx, cleans up
```

### Why two subprocesses

- Matches the existing `AgentRunner` subprocess pattern. No new IPC plumbing — the sidecar JSON is the contract.
- Watchdog timer doesn't false-trip while the user thinks (no process running between phases).
- Each phase is independently retryable.
- Session JSON on disk survives an app restart, which leaves a clean path to add a resume UI later.

## Session file contract

**Location:** `logs/depo_sessions/<sha1(input_path)[:12]>_<YYYYMMDD_HHMMSS>.json`

**Cached transcript:** sibling file, same basename, `.txt` extension. UTF-8. Written once in phase 1, deleted in phase 2 cleanup on success.

**Schema after phase 1 (`phase: "awaiting_input"`):**

```json
{
  "version": 1,
  "phase": "awaiting_input",
  "input_path": "Z:\\...\\Smith Depo.pdf",
  "cached_text_path": "logs\\depo_sessions\\a1b2c3d4_20260514_153022.txt",
  "deponent_name": "John Smith",
  "deposition_date": "January 15, 2024",
  "deponent_type": "Plaintiff",
  "file_number": "3850.084",
  "topics": [
    {"id": 1, "title": "Pre-accident medical history", "rank": 1, "discussion_density": "high"},
    {"id": 2, "title": "Mechanism of injury", "rank": 2, "discussion_density": "high"}
  ],
  "user_config": null
}
```

**After popup submit (`phase: "ready_for_summary"`):**

```json
"user_config": {
  "selected_topics": ["Pre-accident medical history", "Mechanism of injury"],
  "added_topics": ["Communications with treating providers"],
  "bullets_per_topic": 5,
  "deponent_label": "Plaintiff",
  "custom_rules": "Use past tense. Reference page:line citations where available.",
  "cross_check_enabled": true
}
```

**Atomic writes.** The UI writes user_config by writing to a `.tmp` sibling and calling `os.replace`. Prevents corrupt sessions if the app crashes mid-write.

**Cleanup.** Phase 2 deletes the session JSON and the cached `.txt` only on a successful summary save. On failure both are left on disk for debugging.

## Agent script changes (`Scripts/summarize_deposition.py`)

`main()` becomes a `--phase` dispatcher. Shared helpers (`extract_file_number`, `DeponentExtractor`, `get_output_directory`, `save_to_docx`, `add_markdown_to_doc`) stay as-is. The `process_document` function is split into `process_topics` (phase 1) and `process_summary` (phase 2).

### Phase 1 — `process_topics(input_path, logger)`

1. Extract text with the existing `DocumentProcessor.extract_with_dynamic_ocr` (OCR-as-needed pipeline unchanged).
2. Cache the extracted text to `logs/depo_sessions/<hash>_<ts>.txt`.
3. Use existing `DeponentExtractor` to pull deponent name, date, type.
4. Call `LLMCaller.call(topic_discovery_prompt, text, task_type="summary")`. Same model sequence as the narrative pass (user explicitly does not want the cheaper extraction model for topic discovery).
5. Parse the LLM response as JSON. Fall back to bullet-list parsing if JSON parsing fails. Truncate to the top 25 topics by rank in code (don't rely on the LLM to respect a cap).
6. Write session JSON.
7. Print `AWAITING_INPUT:<session_path>` to stdout (newline-terminated, `flush=True`).
8. Exit 0.

**Progress tokens:** 5% extract → 20% cached → 60% LLM → 95% session written. UI sees the status idle at 95% until READY is clicked.

### Phase 2 — `process_summary(session_path, logger)`

1. Load session JSON. Validate `phase == "ready_for_summary"` and `user_config` is present. Hard-fail with `PASS_FAILED:` if not.
2. Read cached text from `session.cached_text_path`. Hard-fail if missing.
3. Build the topic-locked summary prompt (see prompt changes below).
4. Call `LLMCaller.call(prompt, text, task_type="summary")`. Single call. No parallel extraction pass.
5. If `user_config.cross_check_enabled`, run the cross-check pass against the original transcript (no separate extraction input).
6. Call existing `save_to_docx` with deponent name/date pulled from session JSON.
7. Register in `CaseDataManager` and `DocumentRegistry` exactly as today.
8. On success, delete session JSON and cached text. Emit progress to 100.

**Phase 2 progress tokens:** 5% load session → 10% read cached text → 30% LLM start → 70% LLM done → 85% cross-check (if enabled, else skip to 85) → 95% docx saved → 100%.

### Prompt files

- **`Scripts/DEPOSITION_TOPIC_DISCOVERY_PROMPT.txt`** (new). Asks the LLM to return a JSON array of `{title, rank, discussion_density}` objects sorted from most-important/most-discussed to least. Includes JSON schema example and "JSON only — no prose" instruction.
- **`Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt`** (rewritten). Topic-locked template:
  - Injects `{deponent_label}`, `{bullets_per_topic}`, `{custom_rules}` near the top.
  - Includes the final ordered topic list (selected + added) as explicit section headings.
  - Tells the LLM: "Generate exactly N bullets under each heading. Do not add other sections. Do not skip topics."
- **`Scripts/DEPOSITION_CROSS_CHECK_PROMPT.txt`** (modified). The `{extraction}` placeholder is removed; the prompt now compares `{summary}` against `{original}` only.
- **`Scripts/DEPOSITION_EXTRACTION_PROMPT.txt`** — no longer referenced. Leave the file in place to avoid breaking other agents that may share it; just don't load it from this script.

## UI changes

### `AgentRunner` extension (`icharlotte_core/ui/widgets.py`)

New signal:

```python
awaiting_input = Signal(str)  # session_path
```

New branch in `parse_progress`, placed alongside the existing `OUTPUT_FILE:` and `PROGRESS:` branches:

```python
if line.startswith("AWAITING_INPUT:"):
    session_path = line[len("AWAITING_INPUT:"):].strip()
    self.session_path = session_path
    self.awaiting_input.emit(session_path)
    continue
```

`handle_finished` is taught the paused-exit case: if `self.session_path` is set and `self.success is None`, an exit-0 is treated as "paused for input" — the `finished` signal is **not** emitted, the status widget stays alive, the watchdog timer is stopped (no running process to monitor).

New method `resume_with_config(session_path)`:
- Instantiates a fresh `QProcess` on `self`.
- Re-wires `readyReadStandardOutput` / `readyReadStandardError` / `finished` slots.
- Starts the process with args `["--phase=summary", session_path]`.
- Preserves `log_history`, `pass_info`, `last_progress` so the status row shows continuous history across the phase boundary.
- Starts the watchdog timer again.

### Status widget

New slot `on_awaiting_input(session_path)`:
- Shows a prominent **READY** button next to the progress bar.
- Tooltip: "Click to choose topics and configure summary."
- Stores the `session_path` for later.
- Emits `ready_clicked(session_path)` when the button is pressed.

The existing cancel button cleans up the session JSON + cached text if pressed while paused.

The signal is wired up by the parent tab (`IndexTab` — same place `cancel_requested` and `retry_pass_requested` are connected) which:
1. Loads the session JSON.
2. Opens `DepoSummaryConfigDialog` (modal, non-blocking to other status rows).
3. On Accept, calls `agent_runner.resume_with_config(session_path)`.
4. On Cancel, leaves the READY button clickable — user can re-open the popup.

### Popup — `DepoSummaryConfigDialog`

New file: `icharlotte_core/ui/depo_summary_config_dialog.py`. `QDialog` with `Qt.ApplicationModal`, ~700×600, vertical scroll for the topic list. Only one popup is open at a time across the whole app — if multiple status rows are in READY state, the user clicks each row's READY button in turn after closing the previous popup.

Layout top to bottom:

1. **Header label** — "Configure summary for *{deponent_name}* ({deponent_type}, {deposition_date})".
2. **Topic rows** (scrollable QVBoxLayout): one row per agent-suggested topic, in rank order. Each row is `[QCheckBox (checked)] [QLineEdit (editable, pre-filled with title)]`.
3. **Additional topics** — `QPlainTextEdit`, ~3 lines tall, placeholder "One per line. These get added to the summary in the order entered.".
4. **Settings row** (horizontal):
   - `QSpinBox`: "Bullets per topic" (default 5, range 1–15).
   - `QLineEdit`: "Deponent label" (pre-filled with `deponent_type` from session).
   - `QCheckBox`: "Run cross-check pass" (default checked).
5. **Custom rules** — `QPlainTextEdit`, ~4 lines tall, placeholder "Any extra instructions for the summary (tense, citation style, things to avoid, etc.)".
6. **Buttons** — Cancel | Generate Summary.

On Accept: assemble `user_config`, set `session.phase = "ready_for_summary"`, atomically write to disk (`.tmp` + `os.replace`), close dialog.

No persistence between invocations — every open starts with the defaults above (with topic titles and deponent label pre-filled from the session).

### Batch / folder mode

Folder mode keeps working as today (each file gets its own subprocess agent and its own status row). Each phase 1 surfaces its own READY button on its own row. Multiple sessions can be in `awaiting_input` state simultaneously, but only one popup is open at a time (application-modal); the user clicks each row's READY button in turn.

## Testing

### Unit tests — agent script (`tests/test_deposition/test_summarize_deposition_phases.py`)

LLM calls monkey-patched.

- `test_phase1_writes_session_json_and_caches_text` — synthetic transcript, canned topic JSON. Assert session file has all required keys, cached `.txt` exists, last stdout line is `AWAITING_INPUT:<path>`, exit code 0.
- `test_phase1_handles_malformed_llm_json` — stub returns a bulleted list, not JSON. Assert fallback parser produces a topics array; run still succeeds.
- `test_phase1_caps_topic_count` — stub returns 50 topics. Assert session JSON contains at most 25.
- `test_phase2_reads_session_and_generates_summary` — session + cached text fixture, cross-check disabled. Stub summary LLM to a known string. Assert `save_to_docx` is called with that string, session and cached text are deleted after success.
- `test_phase2_cross_check_runs_only_when_enabled` — two runs of the same fixture, `cross_check_enabled` true and false. Assert two LLM calls vs one.
- `test_phase2_uses_selected_plus_added_topics` — `user_config` with 2 selected + 1 added. Assert the summary prompt contains all 3 headings in order.

### Unit tests — UI (`tests/test_deposition/test_agent_runner_awaiting_input.py`, `test_depo_summary_config_dialog.py`)

Uses `pytest-qt` (verify availability when implementing; install if missing).

- `test_agent_runner_emits_awaiting_input_on_token` — feed `AWAITING_INPUT:C:\some\path.json\n` into `parse_progress`. Assert signal fires with `"C:\some\path.json"`.
- `test_agent_runner_does_not_emit_finished_when_paused` — simulate QProcess `finished(0)` after `awaiting_input` was emitted. Assert `finished` is NOT emitted.
- `test_resume_with_config_starts_phase_two_process` — mock `QProcess.start`. Call `resume_with_config(session_path)`. Assert args `["--phase=summary", session_path]`.
- `test_dialog_loads_session_and_populates_topics` — fixture session JSON. Assert each topic row has a checked QCheckBox and an editable QLineEdit with the title.
- `test_dialog_accept_writes_user_config_back_to_session` — open, uncheck topic 2, add "Custom topic X", bullets=7, deponent="Mr. Smith". Click Generate. Assert session JSON now contains exactly that `user_config`.
- `test_dialog_cancel_does_not_modify_session` — open, change values, click Cancel. Assert file unchanged.
- `test_dialog_atomic_write` — patch `os.replace` to raise. Assert original session JSON is not corrupted.

### Integration smoke test (`tests/test_deposition/test_full_flow_smoke.py`)

End-to-end with mocked LLM but real subprocess. Spawn phase 1 against a 200-line fake transcript fixture, capture `AWAITING_INPUT:` from stdout, write user_config into the session, spawn phase 2, assert the output `.docx` exists and contains the expected topic headings.

### Manual test plan

Run after implementation per CLAUDE.md ("always test after developing or changing a feature"):

1. Single deposition PDF in a test case folder. READY appears, popup populates, output docx looks right.
2. Folder of 2–3 depositions. Each gets its own status row and its own READY button. Popups can be opened in any order.
3. Cancel popup, click READY again — popup re-opens.
4. Click status-row cancel while paused — session JSON and cached text get cleaned up.
5. Cross-check unchecked — only one LLM summary call in the agent log.

## Open questions

None. All clarifications resolved during brainstorming.

## Files touched

**New:**
- `Scripts/DEPOSITION_TOPIC_DISCOVERY_PROMPT.txt`
- `icharlotte_core/ui/depo_summary_config_dialog.py`
- `tests/test_deposition/__init__.py`
- `tests/test_deposition/test_summarize_deposition_phases.py`
- `tests/test_deposition/test_agent_runner_awaiting_input.py`
- `tests/test_deposition/test_depo_summary_config_dialog.py`
- `tests/test_deposition/test_full_flow_smoke.py`

**Modified:**
- `Scripts/summarize_deposition.py` (phase dispatcher; remove ExhibitExtractor, ImpeachmentDetector, parallel extraction pass; rewrite `process_document`)
- `Scripts/SUMMARIZE_DEPOSITION_PROMPT.txt` (topic-locked template)
- `Scripts/DEPOSITION_CROSS_CHECK_PROMPT.txt` (drop `{extraction}` placeholder)
- `icharlotte_core/ui/widgets.py` (AgentRunner: `awaiting_input` signal, `AWAITING_INPUT:` parse branch, paused-exit handling, `resume_with_config`; StatusWidget: `on_awaiting_input` slot, READY button, `ready_clicked` signal)
- `icharlotte_core/ui/tabs.py` (wire `ready_clicked` → open dialog → `resume_with_config`)

**Unchanged but referenced:**
- `icharlotte_core/agent_logger.py`
- `icharlotte_core/llm_config.py`
- `icharlotte_core/document_processor.py`
- `Scripts/case_data_manager.py`
- `Scripts/document_registry.py`
