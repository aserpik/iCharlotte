# Task Debug Console Design

## Goal

Add an always-capturing, external debug console for iCharlotte task execution so wizard-mode and chat legal-research runs expose detailed step-by-step background activity for troubleshooting.

## Approved Approach

Use a structured task-debug event bus plus a floating debug console and per-run JSONL traces. The console is opened from the existing `View` menu as `Debug Console`. Closing the window only hides the viewer; task events continue to be captured in memory and written to disk.

## User-Facing Behavior

- `View > Debug Console` opens a separate floating window.
- The debug console shows all task events emitted since app start, including events created while the console was closed.
- The window provides task/source/level filtering, text search, pause autoscroll, clear view, copy selected/all visible lines, and open trace folder.
- Each run has a run id and a trace file under `logs/task_debug/`.
- Normal task status pages keep their current concise messages. The debug console carries the granular detail.

## Event Model

Each debug event has:

- `timestamp`: ISO timestamp with local time.
- `run_id`: stable id for one task run.
- `task_id`: wizard task id or `chat_legal_research`.
- `task_title`: display name, such as `Oppose Motion` or `Chat Legal Research`.
- `phase`: high-level stage, such as `start`, `extract`, `research`, `search`, `select`, `verify`, `save`, `finish`, or `error`.
- `level`: `debug`, `info`, `warning`, or `error`.
- `message`: human-readable detail.
- `elapsed_ms`: milliseconds since the run started when available.
- `source`: component that emitted the event, such as `TaskTab`, `ChatLegalResearchService`, `CourtListener API`, or `Local California corpus`.
- `details`: small JSON-safe metadata such as counts, selected sources, file basenames, warning counts, freshness reason, or candidate totals.

The `details` field must not include API keys, raw full document text, full prompts, or full LLM responses.

## Architecture

### Core Debug Module

Add `icharlotte_core/task_debug.py`.

Responsibilities:

- Define `TaskDebugEvent`.
- Own a process-wide recorder with a bounded in-memory event buffer.
- Create and track run contexts.
- Write JSONL traces to `logs/task_debug/`.
- Provide `emit_event(...)`, `start_run(...)`, and `finish_run(...)` helpers.
- Provide a Qt signal bridge so UI widgets can subscribe without polling.

The module stays small and dependency-light. Pure service code can receive a debug callback from UI code instead of importing Qt-specific UI widgets.

### Debug Console Window

Add `icharlotte_core/ui/task_debug_window.py`.

Responsibilities:

- Render a table or log-style list of events.
- Subscribe to the task-debug signal bridge.
- Load the current in-memory buffer when opened.
- Filter by run/task/level/source and search message/details text.
- Pause and resume autoscroll.
- Copy visible entries.
- Open `logs/task_debug/` in Explorer.

The window should be owned by `MainWindow` and reused, not recreated for every click.

### Main Window Integration

Modify `iCharlotte.py` only where needed to:

- Add `Debug Console` to the `View` menu.
- Lazily create and show the console window.
- Keep the window available in both Wizard and Advanced modes.

### Wizard Task Integration

Instrument the central task containers first:

- `TaskTab`
- `InProcessTaskTab`
- custom task tabs for Oppose Motion, Generate Motion, Mediation Brief, Separate PDFs, and Case Intake/Docket

The initial instrumentation wraps existing status/progress/failure/finish wiring so current behavior does not change. It emits:

- run start with selected task, file count, and case number.
- every worker status/progress line.
- warnings and failures.
- output path and completion duration.
- cancellation attempts.

### Chat Legal Research Integration

Extend `ChatLegalResearchService.research(...)` to accept an optional debug callback while preserving the current `status_callback` API.

Initial granular events:

- selected research settings.
- proposition extraction start/finish and proposition count.
- local corpus availability and freshness warning.
- firm authority search start/finish and candidate count.
- local corpus search start/finish, semantic/text passes, hit count, and candidate count.
- CourtListener fallback decision and reason.
- CourtListener search start/finish and candidate count.
- warnings generated.
- authority selection start/finish per proposition.
- final selected authority count.

Chat UI continues to append the existing concise italic status messages. The debug callback emits the detailed events.

## Data Flow

1. A task starts and creates a run context.
2. Existing worker status/progress signals continue to update the current status page.
3. The same signals are mirrored into the task-debug recorder as structured events.
4. The recorder appends JSONL lines to the run trace file.
5. If the debug console is open, it receives the event via the Qt bridge and updates immediately.
6. If the console is closed, the recorder still buffers and writes events. Opening the console later loads the buffered history.

## Error Handling

- Debug logging must never break a task. All recorder writes and UI subscriber notifications fail closed with local `app.log` warnings only.
- If trace directory creation fails, keep in-memory events and continue the task.
- If an event contains a non-JSON value in `details`, coerce it to a safe string.
- If a worker emits a malformed status/progress value, log the raw safe string and continue.

## Privacy And Safety

Included:

- file basenames and selected file counts.
- query/proposition text.
- source names.
- counts, timings, warning messages, freshness reasons, and output paths.

Excluded by default:

- API keys or credentials.
- full document text.
- full prompts.
- full LLM responses.
- raw retrieved opinion text.

Raw prompt debugging is out of scope for this pass.

## Testing Plan

- Unit-test event serialization, detail sanitization, bounded buffer behavior, and JSONL write behavior.
- Unit-test run start/finish helpers and elapsed time handling.
- Qt-test the debug console with synthetic events for filtering, search, pause autoscroll, and copy/open-folder button wiring where practical.
- Wizard tests should assert that representative task status/failure/finish events are mirrored to the recorder without changing existing status page behavior.
- Chat legal research tests should assert granular debug callbacks for proposition extraction, source search counts, fallback decisions, selection, warnings, and completion.

## Out Of Scope

- Capturing full raw prompts or LLM responses.
- Adding remote log upload.
- Replacing existing crash logs.
- Rewriting every task internals for deep domain-specific diagnostics in the first pass.
- Changing legal-research authority selection behavior.
