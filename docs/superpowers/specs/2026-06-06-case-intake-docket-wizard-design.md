# Case Intake & Docket - Wizard Task Design

**Date:** 2026-06-06
**Status:** Approved design, pending implementation plan
**Scope:** Add a combined Wizard Mode task that runs the existing Complaint Agent first, lets the user review key case metadata, then runs the Docket Agent from those reviewed values.

## Goal

Move the advanced-mode complaint and docket workflow into Wizard Mode as a guided case-level task.

The new task should reduce bad docket runs by putting a mandatory human review gate between complaint extraction and docket download. The complaint step remains the source of first-pass metadata. The docket step should consume the reviewed metadata, then update the same case artifacts that the current advanced-mode agents already update.

## Current State

- `iCharlotte.py` creates advanced-mode buttons for `docket.py` and `complaint.py` in the Case View agent panel.
- `run_agent(...)` runs those scripts with the loaded `file_number` and adds `--headless`.
- `Scripts/complaint.py` is a case-level agent. It finds the complaint and insurance/caption documents, extracts metadata, saves variables, updates `variables.docx`, and drafts factual background.
- `Scripts/docket.py` is also a case-level agent. It reads `case_number` and `venue_county`, runs the complaint agent automatically if either value is missing, downloads or skips the docket depending on venue support, extracts hearings, updates master database fields, updates master status, saves procedural history, and updates `variables.docx`.
- The generic Wizard `TaskTab` expects selected document paths. That is not a good fit because complaint and docket are file-number/case-level jobs.
- Existing in-process Wizard tasks use `registry.py`, `task_routing.py`, and custom task builders when generic file selection is not enough.

## Approved Workflow

Wizard task title: **Case Intake & Docket**

Category: **General**

Default behavior:

1. User opens a case and selects the Wizard task.
2. Settings page shows the loaded file number and explains that the flow will run complaint extraction first.
3. User clicks **Run Complaint Intake**.
4. Wizard launches `Scripts/complaint.py <file_number> --headless`.
5. When complaint finishes, the wizard loads current case metadata and displays an editable review page.
6. User reviews and may edit:
   - case number
   - venue county
   - case name
   - filing date
   - plaintiffs
   - defendants
   - client name
   - client email
   - plaintiff counsel
   - causes of action
   - complaint file found, if detectable
7. User clicks **Run Docket**.
8. Wizard persists the reviewed metadata through the same variable storage used by the agents, then launches `Scripts/docket.py <file_number> --headless`.
9. Output page summarizes the results and gives access to:
   - latest `NOTES/AI OUTPUT/Docket_*.pdf`
   - `variables.docx`
   - trial date
   - other hearings
   - procedural history
   - any clear failure or partial-success notes from the subprocess logs

The review page is mandatory. Docket does not run automatically after complaint extraction.

## Architecture

### Task Registration

Add a `TaskSpec` in `icharlotte_core/ui/wizard/registry.py`:

- `task_id="case_intake_docket"`
- title `Case Intake & Docket`
- description `Extract complaint metadata, review case details, then download and process the court docket.`
- `script_name=""` because the task uses a custom page and case-level subprocess orchestration.
- category `General`
- keywords: `complaint`, `docket`, `case number`, `venue`, `intake`, `hearing`, `trial`

Route it through `icharlotte_core/ui/wizard/task_routing.py` to a custom builder in `in_process_task_tab.py` or directly to a page factory, following the existing custom-task pattern.

### New Page Module

Create `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py`.

The page owns a `WizardTaskContainer` with four logical pages:

1. Intake settings
2. Complaint status
3. Metadata review
4. Docket status/output

The task can use a stacked layout similar to `OpposeMotionTaskTab`, because the review page is richer than the generic `InProcessTaskTab` settings/status/output cycle.

### Case-Level Subprocess Worker

Add a small worker dedicated to file-number agents. It should be separate from the existing document-path `SubprocessWorker` because its command contract differs.

Suggested module:

`icharlotte_core/ui/wizard/runners/case_agent_worker.py`

Worker contract:

- input: `script_name`, `file_number`, `case_path`, optional flags
- command: `python -u Scripts/<script_name> <file_number> --headless`
- emit status lines from stdout
- emit finished success/failure
- support cancellation through `QProcess.terminate()` then kill after timeout
- capture recent log lines for the output summary

This worker should not scan `NOTES/AI OUTPUT` for `.docx` the way the generic worker does. Complaint and docket update several case artifacts, and docket may succeed without a new docket PDF if the venue is unsupported or the scraper is skipped.

### Metadata Loading And Saving

The review page should read and write the same variable store used by the scripts.

Implementation should prefer existing project APIs rather than duplicating document parsing:

- load current values after complaint finishes
- write edited values before docket starts
- update `variables.docx` through the existing variable/document update path if that API is already exposed

If implementation touches code that creates, modifies, or saves `.docx`, it must validate through `icharlotte_core/word_validator.py` per project rules. The first implementation should avoid new Word-writing code and instead use existing script/core paths where possible.

### Output Summary

The output page should derive results from current case state after docket finishes:

- newest `Docket_*.pdf` in `NOTES/AI OUTPUT`
- `trial_date` and `other_hearings` from master database or variables
- `procedural_history` from variables
- `variables.docx` path if present
- subprocess success/failure and important final status lines

Partial success must be explicit. For example, if docket exits successfully but no docket PDF exists because the venue is unsupported, the output should say that the docket download was skipped or unavailable rather than implying a PDF was produced.

### Persistence And Reopen

Recent-task entry should store:

- `task_id`
- reviewed settings/metadata
- `completed_at`
- output summary fields
- output paths that exist, especially latest docket PDF and `variables.docx`

Open-tab snapshot should restore:

- settings/review page if unfinished
- output page if completed and saved artifacts still exist

If reopen cannot find a saved artifact, show the task settings/review state and allow rerun.

## Non-Goals

- Do not replace advanced-mode Case View buttons in v1.
- Do not refactor the full complaint or docket scripts into core services in v1.
- Do not add support for unsupported docket venues.
- Do not auto-run docket immediately after complaint extraction.
- Do not build a multi-case batch docket flow in Wizard Mode.

## Testing

Focused tests should cover:

- registry entry exists and appears in the General category
- task routing uses the custom builder and skips the generic file picker
- case-agent worker builds commands as `script file_number --headless`
- complaint completion loads metadata into the review model/page
- review edits are persisted before docket starts
- docket output summary finds latest docket PDF and key case values
- partial success messaging when docket succeeds with no new PDF
- recent-task/open-tab persistence for completed and unfinished task states

Suggested focused run after implementation:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_page.py tests/test_wizard/test_task_routing.py tests/test_wizard/test_wizard_tab.py -q
python -m py_compile icharlotte_core/ui/wizard/pages/case_intake_docket_page.py icharlotte_core/ui/wizard/runners/case_agent_worker.py
```

## Risks

- The existing scripts use broad side effects: variable storage, Word updates, master database updates, scraper subprocesses, and logs. The Wizard integration should wrap those effects rather than duplicate them.
- Docket already auto-runs complaint when required metadata is missing. The new Wizard task intentionally runs complaint first and then writes reviewed values, so docket should normally skip its internal complaint fallback.
- A successful docket exit does not always mean a PDF was downloaded. Output wording must distinguish download success, skipped venue, scraper failure with procedural-history generation, and total process failure.
- Current advanced-mode button state is independent from Wizard task state. V1 can leave those separate, but docket last-download metadata should still update through the script/master database path.
