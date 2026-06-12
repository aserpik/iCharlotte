# Wizard Web Companion — Design

**Date:** 2026-06-12
**Status:** Approved design, pending implementation plan

## Purpose

Let the user run iCharlotte wizard-mode tasks from an iPhone. A standalone web
server on the desktop serves a mobile-friendly UI over Tailscale; the phone
picks a case, picks files, configures settings, launches a task, monitors
progress, answers mid-run prompts, and views the finished `.docx`.

This complements (does not replace) remote-desktop access: the web companion
runs tasks; testing new PyQt UI features still happens over Parsec/CRD.

## Scope

**In scope — the seven script-based wizard tasks** (all run as
`python -u Scripts/<script>.py` subprocesses and speak the stdout protocol):

| Task | Script | Phases |
|---|---|---|
| Summarize Documents | `summarize.py` | single |
| Summarize Discovery | `summarize_discovery.py` | single |
| Summarize Depositions | `summarize_deposition.py` | two-phase (topic picking) |
| Depo Prep | `depo_prep.py` | two-phase (`--phase=analyze` / `--phase=generate`) |
| Medical Records Review | `med_record.py` | single |
| Med Chron Analysis | `med_chron.py` | two-phase (`--phase=prep` / `--phase=run`) |
| Separate Documents | `separate.py` | single |

**Out of scope:** in-process Qt tasks (Respond to Discovery, Oppose Motion,
Generate Motion, Mediation Brief, Med Record Extractor, Subpoena Tracker,
Case Intake & Docket, Chat). These are tangled with Qt worker classes and
would each require extraction work. Possible later phases.

## Decisions made during brainstorming

1. **Task scope:** script-based tasks only (above).
2. **File picking:** simple folder browser rooted at the case folder, starting
   in the task's `default_folders[0]`, with multi-select.
3. **Settings:** per-task settings forms mirroring what the desktop settings
   pages expose (not defaults-only).
4. **Two-phase tasks:** fully supported. The phone shows an
   "awaiting your input" state and a simplified picker (checkboxes +
   free-text), then resumes phase 2.
5. **Outputs:** finished `.docx` served directly; iOS Safari previews Word
   documents natively (and can hand off to Word for iOS).
6. **Deployment:** standalone process (independent of the desktop app), bound
   to the Tailscale interface. Tailscale is the entire auth/security
   perimeter; no login page. `--lan` flag for LAN-only dev binding.

## Approach

**Chosen: standalone job server that speaks the existing stdout protocol.**

The contract between the wizard and the scripts is the stdout protocol emitted
by `Scripts/*.py` and parsed today by
`icharlotte_core/ui/wizard/runners/subprocess_worker.py`:

- `PROGRESS:<int>` or `PROGRESS:<int>:<message>` → progress updates
- `AWAITING_INPUT:<session_path>` → script wrote a session JSON
  (via `session_manager.write_session`) and exited 0; caller edits the session
  and relaunches with the phase-2 flag (`resume_with_config` semantics)
- `OUTPUT:<path>` → authoritative declaration of the produced file
- all other lines → status log

The web companion re-implements the thin subprocess-driving logic with plain
`subprocess.Popen` + a reader thread (no Qt). This duplicates ~150 lines of
`SubprocessWorker` semantics deliberately: the desktop app is untouched (zero
risk to working code), and the protocol — not the Qt class — is the stable
interface.

**Rejected alternatives:**
- *Shared Qt-free runner core extracted from SubprocessWorker:* churns
  working, signal-heavy desktop code to avoid a shallow duplication.
- *HTTP bridge inside the desktop app driving the real wizard:* requires the
  app to be running (conflicts with the standalone decision) and poking Qt
  widgets from an HTTP thread is a threading hazard.

## Architecture

```
iPhone Safari ──Tailscale──> FastAPI server (webcompanion/, port 8765, separate process)
                                │
                                ├─ master_db.py        → case list (read-only)
                                ├─ filesystem          → case folder browser
                                ├─ JobManager          → python -u Scripts/<script>.py <file>
                                │     (Popen + reader thread; parses PROGRESS /
                                │      AWAITING_INPUT / OUTPUT protocol)
                                ├─ jobs.json           → persistent job state
                                └─ file serving        → finished .docx
```

New top-level package `webcompanion/`:

- `server.py` — FastAPI app, routes, startup (binds Tailscale IP by default,
  `--lan` for LAN). Entry point: `python -m webcompanion.server`.
- `job_manager.py` — Qt-free job lifecycle: queue, launch, stdout parsing,
  two-phase pause/resume, cancellation, persistence to `jobs.json`.
  Job states: `queued → running → awaiting_input → running → done`
  plus `failed`, `cancelled`, `interrupted`.
- `task_schemas.py` — per-task settings schema (field name, type, label,
  default, choices) mirroring what each desktop settings page emits, plus the
  mapping of settings into CLI args / session-file edits per task.
- `templates/` — server-rendered mobile pages (Jinja2, plain HTML forms, a
  small inline script for run-page polling; no SPA build step, no JS
  framework).
- `run_webcompanion.bat` — launcher for a Windows startup task.

Mirrors `SubprocessWorker` behaviors that matter: multi-file sequential runs
with scaled progress, `NOTES/AI Output` mtime-snapshot fallback when no
`OUTPUT:` line is emitted, terminate-then-kill cancellation.

## Mobile UI flow

1. **Home** — active/recent jobs (tap to open run page) + case search backed
   by `MasterCaseDatabase.get_all_cases()`.
2. **Case page** — the seven task cards.
3. **File picker** — folder listing rooted at the case path, starting in the
   task's `default_folders[0]` (fallback: case root), breadcrumb navigation,
   checkbox multi-select, PDFs/docx filtered appropriately per task.
4. **Settings form** — rendered from `task_schemas.py` for the chosen task.
5. **Run page** — progress bar + recent log lines, polling every few seconds
   (battery-friendly; SSE deferred unless polling proves inadequate).
6. **Awaiting-input page** — for two-phase tasks: reads the session JSON,
   renders a simplified picker (checkboxes + free-text additions — not the
   desktop's rich topic editor), writes choices back into the session file
   exactly as the desktop form does, then JobManager relaunches with the
   task's `phase2_flag`.
7. **Done page** — tap-to-view link streaming the `.docx` (Safari native
   preview; share sheet to Word for iOS).

## Error handling

- **Script crash / nonzero exit** → job `failed`; last 50 log lines viewable
  on the phone.
- **Server restart** → `jobs.json` reloaded; jobs that were `running` are
  marked `interrupted` (honest state — the child process died with the parent
  or is orphaned; we do not pretend to reattach).
- **Concurrency** → one job at a time per case (FIFO queue per case): the
  scripts share `NOTES/AI Output` and LLM rate limits. Jobs across different
  cases may run concurrently, capped at 2 globally.
- **Desktop app running simultaneously** → safe; scripts are independent
  processes and per-job output detection uses `OUTPUT:` lines or per-run
  mtime snapshots.
- **Path safety** → all file-browser and file-serving paths are validated to
  stay inside the case root (no traversal outside case folders).

## Testing

- **Unit:** `JobManager` against a fake protocol-emitting script — progress,
  status lines, `AWAITING_INPUT` pause/resume, `OUTPUT:` detection,
  mtime-snapshot fallback, crash → `failed`, cancel → terminate/kill,
  persistence round-trip including `interrupted` marking.
- **Endpoint:** FastAPI `TestClient` covering the full flow for one
  single-phase and one two-phase task against a temp case folder.
- **Manual:** one end-to-end run per task family from the actual iPhone over
  Tailscale on a real case before completion (per the mandatory
  test-after-develop rule).

## Open items deferred to the implementation plan

- Exact per-task settings field lists (read from each desktop settings page:
  `deposition_settings_page.py`, `med_chron_settings_page.py`,
  `depo_prep_settings_page.py`, base `settings_page.py`) and how each setting
  reaches its script (CLI arg vs. session-file edit) — to be catalogued
  task-by-task during planning.
- Session JSON schemas per two-phase task (`session_manager` formats for
  deposition topics, depo-prep topics, med-chron analyses).
- Tailscale IP detection on startup (query the Tailscale interface; fail
  loudly with guidance if Tailscale is not running).
