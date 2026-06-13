# Web Companion Chat Task — Design

**Date:** 2026-06-13
**Status:** Approved design, pending implementation plan
**Builds on:** `docs/superpowers/specs/2026-06-12-wizard-web-companion-design.md`

## Purpose

Add the **Chat** task to the wizard web companion so the user can hold a
per-case AI chat conversation from an iPhone over Tailscale — including
attaching case files as context and running legal research — with threads
shared with the desktop iCharlotte chat.

Chat was explicitly out of scope in the original web companion design because
it is interactive and LLM-backed rather than a subprocess script. This spec
adds it as a separate, parallel feature in the same `webcompanion/` package.

## Key architectural decision

**Chat does NOT use `JobManager`/`ScriptRunner`.** Those model a subprocess
`argv` + stdout protocol + `.docx` output. A chat turn is an in-process
synchronous LLM call. Chat gets its own small in-memory background-turn
tracker and reuses the existing desktop building blocks directly:

- `icharlotte_core.chat.persistence.ChatPersistence` — per-case
  `{file_number}_chat.json` (the SAME file the desktop chat uses, so threads
  stay in sync). Methods: `get_conversations()`, `get_conversation(id)`,
  `create_conversation(name, provider, model, system_prompt)`,
  `update_conversation(id, **kwargs)`, `add_message(id, Message)`.
- `icharlotte_core.chat.models.Message` / `Conversation`.
- `icharlotte_core.llm.LLMHandler.generate(provider, model, system_prompt,
  user_prompt, file_contents, settings, history=None)` — with
  `settings={'stream': False}` it returns the full reply string. `history` is
  `[{'role': 'user'|'assistant', 'content': str}, ...]`.
- `icharlotte_core.chat.legal_research.ChatLegalResearchService.from_environment(
  llm_callback=...)` → `.research(user_text, context_text, settings,
  status_callback=...)` → `ChatResearchPacket.build_augmented_system_prompt(
  base_system_prompt)`. Default settings via `ChatResearchSettings.default()`.
- `icharlotte_core.document_processor.DocumentProcessor.extract_text(file_path)`
  for attached-file text.

## Decisions made during brainstorming

1. **Scope:** full conversation management (browse/open/read existing
   conversations, send messages, create new conversations).
2. **Reply delivery:** background turn + page polling (not streaming, not a
   single blocking request) — robust against mobile-Safari/Tailscale timeouts
   and accommodates multi-minute legal research.
3. **Model:** defaults to the conversation's saved provider/model, with a
   phone-side picker to change it per message; the chosen model is saved back
   onto the conversation.
4. **Extras included:** attach case files as context; legal research (on/off
   toggle).
5. **Extras excluded from v1 (YAGNI):** quick prompts, message
   pin/edit/delete, token-by-token streaming, image/audio attachments,
   research sub-settings UI.

## Components

### `webcompanion/chat.py`

Qt-free chat service layer:

- **Persistence wrappers** over `ChatPersistence`:
  `list_conversations(file_number)`, `get_conversation(file_number, conv_id)`,
  `create_conversation(file_number, ...)`, `append_message(file_number,
  conv_id, Message)`.
- **`ChatTurnManager`** — in-memory tracker of in-flight turns (NOT persisted
  to disk; the conversation itself is already persisted via ChatPersistence).
  - `start_turn(file_number, conv_id, user_text, provider, model,
    attach_rel_files, research_on)` → returns a `turn_id`; appends the user
    `Message` immediately, then runs a background thread.
  - Turn statuses: `extracting` → `researching` → `generating` →
    `done` | `failed`. Each turn holds a short status log (for the poll).
  - `get_turn(turn_id)` → status snapshot.
  - Concurrency: at most one in-flight turn per `conv_id`; small global cap
    (2, matching the task side). A second send for a busy conversation is
    rejected (409/avoid in UI).
  - Background pipeline:
    1. **extract** (if `attach_rel_files`): `DocumentProcessor.extract_text`
       per file, concatenated, capped at 100 000 chars (matching desktop
       research-context cap). A file that fails to extract is skipped with a
       logged note.
    2. **research** (if `research_on`):
       `ChatLegalResearchService.from_environment(llm_callback)` then
       `.research(user_text=..., context_text=<extracted text>,
       settings=ChatResearchSettings.default(), status_callback=<append to
       turn log>)`. The returned packet wraps the system prompt via
       `build_augmented_system_prompt`. On any research error: log it and
       **fall back to the un-augmented system prompt** (still answer).
    3. **generate:** `LLMHandler.generate(provider, model, system_prompt,
       user_text, file_contents=<extracted text>, settings={'stream': False},
       history=<prior messages as role/content dicts>)`.
    4. **persist:** append the assistant `Message` (with `model_used`) via
       ChatPersistence; mark turn `done`. On generate error: mark turn
       `failed` with the error text (user message stays persisted).

### `webcompanion/chat_models.py`

`available_models()` → curated list of `(provider, model, label)` pulled from
the existing `LLMConfig` / `config/llm_preferences.json`, for the compose
picker. Defaults resolve to the conversation's saved provider/model.

### Routes (new `_register_chat_routes` in `server.py`)

- `GET  /case/{fn}/chat` — conversation list for the case.
- `POST /case/{fn}/chat/new` — create a conversation → 303 to it.
- `GET  /case/{fn}/chat/{conv_id}` — conversation view: message history +
  compose form (textarea, model `<select>`, research toggle, "attach files"
  link). If a turn is in flight, shows the thinking bubble + poll script.
- `GET  /case/{fn}/chat/{conv_id}/attach` — reuses `cases.browse` /
  `safe_resolve` to pick case files; selections ride back to the compose form.
- `POST /case/{fn}/chat/{conv_id}/send` — validates, calls
  `ChatTurnManager.start_turn`, 303 back to the conversation view (which now
  renders the thinking bubble for `turn_id`).
- `GET  /api/chat/{conv_id}/turn/{turn_id}` — JSON `{status, log, done}`.
  When `done`, the page reloads the conversation to show the persisted reply.

### Templates

- `chat_conversations.html` — conversation cards + "New conversation".
- `chat_conversation.html` — message bubbles (user/assistant), attachment
  chips, compose form, and the inline poll script (same pattern as
  `job.html`).
- The attach picker reuses a small variant of the existing picker markup.

### Case-page wiring

The web companion's `TASKS` does not include `chat`. Add a **Chat card** to
`case.html` that links to `/case/{fn}/chat` (NOT the `/task/{id}` file-picker
route the script tasks use). It renders alongside the task cards.

## Data flow (one message)

```
compose ──POST /send──> append user Message (persists; shows immediately)
                        └─ ChatTurnManager.start_turn → background thread
                        └─ 303 redirect to conversation (thinking bubble for turn_id)
background: [extract attached file text]
            → [legal research → augment system prompt]   (on failure: plain prompt)
            → LLMHandler.generate(..., stream=False, history=prior msgs)
            → append assistant Message (persists) → turn=done
page poll /api/chat/{conv}/turn/{turn} → done → reload conversation → reply shown
```

The user message persists up front so nothing is lost if the reply fails.

## Error handling

- **Research failure** → log + fall back to plain generate (still answers).
- **Generate failure** → turn `failed` with error text in the bubble; user
  message persists; a "Retry" link re-sends the same text.
- **File extraction failure** (per file) → skip that file, logged note,
  proceed with the rest.
- **Second send to a busy conversation** → rejected; UI disables send while a
  turn is in flight.
- **Unknown case / conversation / turn id** → 404.
- **Missing `COURTLISTENER_API_TOKEN`** → research still runs with firm +
  local authority (CourtListener portion degrades gracefully, same as
  desktop).

## Testing

- **Unit (`tests/test_webcompanion/test_chat.py`):** temp `GEMINI_DATA_DIR`,
  monkeypatched `LLMHandler.generate`. Covers: create/list/get conversation;
  send appends the user message immediately; turn lifecycle reaches `done` and
  persists the assistant message with `model_used`; `history` passed in
  role/content shape; research path augments the system prompt (monkeypatched
  `ChatLegalResearchService`); research failure falls back to plain generate;
  generate failure → turn `failed`; extracted file text capped at 100k;
  one-turn-per-conversation enforced.
- **Endpoint (append to `test_server.py`):** TestClient with the chat layer
  mocked — conversation list/view render; `/new` creates + 303; `/send`
  appends + starts a turn + 303; `/api/.../turn` reports status; case page
  shows the Chat card linking to `/case/{fn}/chat`.
- **Manual (phone over Tailscale):** open case → Chat → existing thread loads
  → plain message → reply via poll → message with a file attached → message
  with legal research on (watch status log) → confirm the same thread shows
  the new messages in the desktop app.

## Implementation phasing (for the plan)

1. Chat service + persistence wrappers + turn manager (core send/generate,
   no extras) + conversation list/view/send/poll routes + templates + Chat
   card. End-to-end plain chat working.
2. Model picker (`chat_models.py` + compose `<select>` + save-back).
3. Attach case files as context (attach picker + extraction step).
4. Legal research toggle (research step + status log + fallback).

Each phase is independently testable and leaves a working feature.

## Open items deferred to the implementation plan

- Exact `available_models()` shape from `LLMConfig` (read
  `icharlotte_core/llm_config.py` during planning).
- Whether `DocumentProcessor.extract_text` returns an object with a `.text`
  attribute (`ExtractResult`) vs a string — confirm and adapt the extraction
  helper accordingly.
- The default system prompt string (desktop uses: "You are a helpful legal
  assistant. Do not provide any disclaimers about being an AI or not being an
  attorney. Provide direct analysis only.") — reuse verbatim for new
  conversations.
