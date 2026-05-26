# Respond to Discovery Wizard — Parallel Generation + Streaming Review

**Date:** 2026-05-26
**Status:** Design — pending user approval
**Scope:** `icharlotte_core/ui/wizard/pages/respond_discovery_page.py`, `icharlotte_core/discovery/response_generation_engine.py`, `icharlotte_core/discovery/response_review_state.py`, new tests under `tests/test_wizard/` and `tests/test_discovery/`.

## Problem

The current Respond to Discovery wizard generates proposed objections + substantive responses **serially**, one LLM call per request, inside a single QThread. On a typical large set (35–100+ requests for an RPD or special interrogatory propounding) this takes several minutes, during which the user sees only a static status label.

Two pain points emerged from review:

1. **Speed.** Serial calls dominate wall-clock. Wait-for-everything-before-review forces a long idle stretch.
2. **Quality.** The prompt asks the model to populate `proposed_objections`, which the application then *discards* in favor of canned rule text — wasted tokens and a confusing instruction.

This design addresses both with a single coordinated change.

## Goals

- Generate proposals concurrently (cap 8 in-flight) instead of serially.
- Show the review screen immediately after parsing — populate proposals live as they arrive.
- Let the user edit, approve, and even start regenerating individual requests while later ones are still cooking.
- Drop the dead `proposed_objections` field from the prompt.
- Separate parse work from draft work so re-drafting (e.g., after changing context files) doesn't re-parse.
- Add a per-request **Regenerate** button.

## Non-goals

- Smarter context selection (embeddings, OCR fallback). Tracked as a separate spec.
- Self-critique pass. Separate spec.
- Case-level memory of past responses. Separate spec.
- Review screen redesign beyond what streaming requires (list+detail, keyboard nav). Separate spec.

## Architecture

```
discovery file picked
        │
        ▼
DiscoveryParseWorker (QThread, one-shot)
        │
        │ emits parse_finished(ParsedDiscovery)
        ▼
RespondDiscoverySettingsPage caches parsed_discovery
        │
        ▼  user picks context files, clicks Next
        │
        ▼
ProposalCoordinator.start(parsed, rules, chunks, ...)
        │
        ├─ enqueues N ProposalTask(QRunnable) into QThreadPool (cap 8)
        │
        ▼
Review screen opens immediately with placeholder rows
        │
        ▼
ProposalTask N completes
        │ emits proposal_ready(req_number, proposal) — or fallback w/ needs_review
        ▼
Page updates RequestReview in self.review_state and refreshes UI
        │
        ▼ all done
ProposalCoordinator.all_done → Finalize button enabled
```

### New components

#### `DiscoveryParseWorker(QThread)`
- **Inputs:** `discovery_file: str`, `detected_type: str`.
- **Behavior:** reads the PDF (existing `read_document_text`), calls the parse LLM (`build_parse_prompt` → `call_llm` with `agent_id="agent_sum_disc"`), runs `parse_llm_response`, normalizes + filters via `_normalize_and_filter_parsed_discovery`.
- **Emits:** `parse_finished(success: bool, parsed_or_error: ParsedDiscovery | str)`.
- **Lifecycle:** runs once when the discovery file is set; result cached on the page; never re-runs unless the file changes.

#### `ProposalTask(QRunnable)`
- **Inputs:** `request: ParsedRequest`, `parsed: ParsedDiscovery`, `context_packet: str`, `selected_rules: list[ResponseRule]`, `response_rules: ResponseRules`, `fi_mode: str`, `override_instruction: str = ""`.
- **Behavior:**
  1. Build the structured-proposal prompt via `build_structured_proposal_prompt` (now with the dead field removed). If `override_instruction` is non-empty, append it as an "ADDITIONAL INSTRUCTIONS:" tail.
  2. `call_llm(prompt, "", task_type="general", agent_id="agent_sum_disc")`.
  3. `parse_structured_proposal_response(raw)`. On parse failure, build the existing repair prompt, retry once.
  4. `_ensure_context_warning(proposal, context_packet)`.
  5. On any exception in steps 1–4: return `build_fallback_structured_proposal(...)` with `needs_review=true` and the truncated exception text in `review_reason`.
- **Signals (via `WorkerSignals(QObject)` companion):**
  - `proposal_ready(req_number: str, proposal: StructuredProposal)`.
  - `proposal_failed(req_number: str, reason: str)` — emitted only when the fallback itself errors, which would be a bug.

#### `ProposalCoordinator(QObject)`
- **Owns:** `QThreadPool` with `setMaxThreadCount(max_concurrent)`. Default `max_concurrent = 8`, read from `config/llm_preferences.json` key `discovery_response.max_concurrent_proposals`.
- **State:** `_in_flight: set[str]` of request numbers, `_results: dict[str, StructuredProposal]`, `_total: int`, `_cancelled: bool`.
- **Public API:**
  - `start(parsed, selected_rules, context_chunks, response_rules, fi_mode)` — for each request: skip per `_should_skip_structured_proposal` (FI fixed mode), otherwise build the per-request context packet via `select_context_packet` + `format_context_packet` and enqueue a `ProposalTask`. Coordinator owns packet construction so tasks stay deterministic and cheap to construct. Returns immediately.
  - `regenerate(request_number: str, override_instruction: str = "")` — re-enqueues one task. No-op if `request_number` already in `_in_flight`.
  - `cancel()` — sets `_cancelled = True`. Late-arriving results are discarded by checking the flag before emitting.
  - `is_done() -> bool` — `len(_results) == _total and not _in_flight`.
- **Signals:**
  - `proposal_ready(req_number: str, proposal: StructuredProposal)`.
  - `progress(completed: int, total: int)`.
  - `all_done()` — emitted once when the last task finishes.
- **Test seam:** accepts an optional `task_factory` callable so tests can substitute a synchronous fake.

### Changes to existing components

#### `response_review_state.py`
- Add `is_pending: bool = False` to `RequestReview`. Serialized to/from dict.
- Add `pending_replacement: StructuredProposal | None = None` to `RequestReview` — set when a fresh draft arrives for a request the user has edited but not yet viewed. Not serialized (in-memory session state).
- `all_approved()` returns False if any request is still pending (paranoia — UI already blocks finalize while pending).

#### `response_generation_engine.py`
- `build_structured_proposal_prompt`:
  - Remove the `"proposed_objections": "ignored by application"` line from the JSON schema block.
  - Add above the schema: `"Do not draft objection text. The application formats objections from the rule IDs you select."`
- `StructuredProposal.proposed_objections` — keep the field (default `""`) so old persisted review states deserialize. Already ignored downstream.

#### `respond_discovery_page.py`
- Remove `RespondDiscoveryProposalWorker` and its callers.
- Remove `_build_structured_proposal_map` (its work moves into `ProposalCoordinator.start` + `ProposalTask.run`).
- Keep `_apply_fixed_fi_proposal_warnings`, `_callbacks_from_proposal_map` — still used when proposals merge into the review state.
- Replace `_generate_review_state_from_proposals` (which assumed an all-at-once map) with two functions:
  - `_build_pending_review_state(parsed, response_rules, fi_mode) -> ReviewState` — produces one `RequestReview` per request. Requests skipped per `_should_skip_structured_proposal` get their fixed FI response immediately and `is_pending=False`. All other requests get empty draft text and `is_pending=True`.
  - `_apply_proposal_to_review_state(review_state, req_number, proposal, parsed, selected_rules, response_rules, fi_mode)` — invoked on each `proposal_ready`. Calls `apply_structured_proposal(...)` to merge the proposal into the matching `RequestReview`, runs `_apply_fixed_fi_proposal_warnings` for that one row, sets `is_pending=False`. Returns the updated `RequestReview` for UI refresh.
- New methods on `RespondDiscoverySettingsPage`:
  - `_on_parse_finished(success, parsed_or_error)` — wires `DiscoveryParseWorker`.
  - `_open_review_screen_with_pending(parsed)` — creates a `ReviewState` where every non-skipped request is pending. Skipped FI requests get their fixed responses immediately.
  - `_on_proposal_ready(req_number, proposal)` — updates the RequestReview, refreshes the list badge, optionally refreshes the detail panes if current.
  - `_on_proposal_conflict(req_number, new_proposal)` — when the user has edited the pending request, shows a "New draft available" banner with View/Apply/Discard.
  - `_on_regenerate_clicked(instruction)` — calls `coordinator.regenerate(...)`.

### Review screen UI changes

Minimal — just enough to support streaming:

- A small **status bar** at the bottom of the review widget: `"Generated 17 / 50 • 2 needs review • 0 approved"`. Updates on `progress` and on user approval.
- Each `RequestReview` gets a one-character status indicator next to the request number in the header label:
  - `⏳` is_pending
  - `⚠` needs_review
  - `✓` approved
  - `✏` user-edited but not approved
- When `is_pending`, the objection and response edits show `"Generating..."` placeholder text and are read-only (`setReadOnly(True)`). Re-enabled on `proposal_ready`.
- **Regenerate row** below the response edit:
  - `QLineEdit` (placeholder: "Optional instructions for regeneration") + `QPushButton("Regenerate")`.
  - Click sends `_on_regenerate_clicked(line_edit.text())`. Pane goes back to pending state until result arrives.
- **Conflict handling on regenerate:** if the user has typed in the editor, the new draft is *not* auto-applied. Instead, a yellow banner appears: `"New draft available — View / Apply / Discard"`. View opens a small diff dialog (existing pattern from word_assistant).
- **Finalize button** disabled until `coordinator.is_done()` returns True AND `review_state.all_approved()` returns True.

The prev/next pager and quick-objection/quick-response panels remain unchanged.

## Concurrency, rate limits, failure

- 8 concurrent tasks by default, configurable via `llm_preferences.json`.
- `LLMCaller`'s existing model-fallback chain handles transient provider errors; no second retry layer added here.
- One repair-prompt retry on JSON parse failure inside `ProposalTask`.
- Any unhandled exception in `ProposalTask` → `build_fallback_structured_proposal` with `needs_review=True` and reason from `str(exc)[:200]`. Coordinator treats it the same as a success — UI shows ⚠.
- Cancellation: coordinator's `_cancelled` flag is checked before each task emits its signal. In-flight LLM calls run to completion but their results are dropped. The thread pool drains naturally.

## Edit-conflict policy

When `proposal_ready` arrives for a request the user has edited:

- If the request is currently displayed: show the "New draft available" banner. Banner buttons: **View** (diff dialog), **Apply** (overwrites editor with new draft), **Discard** (dismisses banner, keeps user edits).
- If the request is *not* currently displayed: silently store the new proposal in `RequestReview.pending_replacement: StructuredProposal | None`, and add a small `(new draft)` suffix on its list badge. Banner re-appears when the user navigates back.

User edits are sacred — never auto-overwritten.

## Testing

### Unit tests

**New file: `tests/test_wizard/test_proposal_coordinator.py`**
- Test fixture injects a synchronous fake `task_factory` (just runs the task body in-thread).
- `test_start_enqueues_one_task_per_non_skipped_request`
- `test_proposal_ready_emitted_per_task`
- `test_all_done_emitted_when_last_task_completes`
- `test_progress_signal_increments_on_each_completion`
- `test_regenerate_resubmits_only_named_request`
- `test_regenerate_no_op_when_already_in_flight`
- `test_cancel_drops_late_results`
- `test_max_concurrent_honored` (mocks `QThreadPool.maxThreadCount`)
- `test_skipped_requests_use_fixed_response` (FI fixed mode)
- `test_task_exception_uses_fallback_with_needs_review`

**Extend `tests/test_wizard/test_oppose_motion_page.py` patterns into `tests/test_wizard/test_respond_discovery_page.py`** (file already exists per repo layout):
- `test_review_screen_opens_with_all_requests_pending`
- `test_status_badge_transitions_on_proposal_ready`
- `test_user_edit_survives_late_proposal_ready` (conflict banner appears)
- `test_regenerate_button_triggers_one_coordinator_call`
- `test_finalize_disabled_while_any_pending`

### Integration tests

**Extend `tests/test_discovery/test_response_generation_engine.py`:**
- `test_prompt_no_longer_includes_proposed_objections_field`
- `test_apply_structured_proposal_handles_missing_field`
- `test_old_persisted_review_state_with_proposed_objections_still_loads`

### Manual verification (per CLAUDE.md)

- Run on a real 35+ request FROG set with the actual user flow.
- Wall-clock baseline → expect ~5–8× improvement on Gemini Flash (50-request set: minutes → 30–60s).
- Review screen interactive within ~5s of completing context-file selection.
- One deliberately-broken request (e.g., point at an unreadable context file or kill the API key briefly) shows the ⚠ badge with a useful reason.
- Conflict path: start typing in a pending request, wait for its draft to arrive, confirm banner appears and user text is preserved.

## Migration / compatibility

- `RespondDiscoveryProposalWorker` is removed in this PR.
- Persisted `ReviewState` from prior versions:
  - Loads cleanly because new field `is_pending` defaults to `False`.
  - Loads cleanly because `StructuredProposal.proposed_objections` is kept as a deserializable field, just unused.
- `config/llm_preferences.json` gains `discovery_response.max_concurrent_proposals: 8` on first run via a default-merge in `LLMConfig._load`.

## Out of scope (future work)

- Embeddings-based context selection.
- OCR fallback for scanned context PDFs.
- Self-critique pass to auto-flag suspect drafts.
- Per-firm memory of past edits.
- Full review screen redesign (list+detail, keyboard nav, bulk operations).

These are real wins but each deserves its own spec.
