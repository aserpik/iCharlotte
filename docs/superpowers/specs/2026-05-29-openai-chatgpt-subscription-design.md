# OpenAI via ChatGPT Subscription (Codex CLI routing) — Design

**Date:** 2026-05-29
**Status:** Approved (design); pending implementation plan
**Author:** iCharlotte / Claude

## Problem

The chat tab — and all other OpenAI usage in iCharlotte — authenticates against
the OpenAI **API platform** using `OPENAI_API_KEY`, billed per token
([`llm.py`](../../../icharlotte_core/llm.py) `provider == "OpenAI"` branch,
~line 332). The user pays for a **ChatGPT subscription** (~$100/month) and wants
OpenAI model usage to draw from that subscription instead of incurring separate
per-token API charges.

OpenAI does not expose the ChatGPT subscription through `api.openai.com`. The only
supported path to subscription-billed model usage is OpenAI's own **Codex** login
("Sign in with ChatGPT"), which third-party tools (Cline, Roo Code, opencode,
OpenClaw) consume today and which OpenAI has — unlike Anthropic/Google — left open.

## Goal

Route OpenAI model calls through the user's ChatGPT subscription via the official
**Codex CLI**, applied **everywhere** OpenAI is used (chat tab + 28 agents + Word
assistant + discovery/respond/liability tabs), with **automatic fallback to the
existing API-key path**.

## Non-Goals

- Native in-app OAuth/PKCE implementation (explicitly rejected — too fragile/
  high-maintenance; see "Alternatives").
- Replacing or removing the existing API-key OpenAI path. It remains as the
  fallback.
- Supporting non-Codex models (gpt-4o, o1, etc.) on the subscription — those are
  not available through the Codex backend and will fall back to the API key.

## Chosen Approach: Codex CLI routing

This mirrors the **existing Claude branch** ([`llm.py`](../../../icharlotte_core/llm.py)
~line 438), which already routes Claude calls through the `claude` CLI to use the
user's Max subscription instead of `ANTHROPIC_API_KEY`. We add the OpenAI analog
using OpenAI's `codex` CLI.

### Prerequisites (environment)
- Node v22.14.0 and npm 10.9.2 are present (verified 2026-05-29).
- One-time setup: `npm install -g @openai/codex`, then `codex login` (browser,
  sign in with ChatGPT). Codex stores auth under the user home (`~/.codex`), which
  is reachable by agent subprocesses.

### Architecture

New helper `_generate_openai_codex_cli(...)` in `llm.py`, called from the top of
the existing `provider == "OpenAI"` branch:

```
if openai_subscription_enabled() and codex_available():
    try:
        return _generate_openai_codex_cli(model, system_prompt, user_prompt,
                                          file_contents, history, settings)
    except Exception as e:
        log_event(f"Codex subscription path failed, falling back to API key: {e}", "warning")
        # fall through
# existing api.openai.com code unchanged — the fallback
```

The existing API-key code is left exactly as-is. Because this lives inside
`generate()`, every caller picks it up with no per-call-site changes.

### Codex invocation

Build a `codex exec` command (parallel to how the Claude branch builds `claude -p`),
constrained to behave as a read-only pure text generator:

- Base args (exact flags confirmed against installed `codex --help` during
  implementation):
  `codex exec --sandbox read-only --ask-for-approval never --skip-git-repo-check -m <model>`
  - `--sandbox read-only`: no file writes.
  - `--ask-for-approval never`: no interactive command-approval prompts.
  - `--skip-git-repo-check`: don't require a git repo.
- **Working directory:** a neutral temp dir (e.g. `TEMP_DIR`), NEVER inside
  `Z:\Shared\Current Clients`, so Codex does not index client files.
- **Prompt:** combined string (system + history + user + attached files) passed via
  stdin, exactly like `combined_prompt` in the Claude branch. Because Codex has no
  clean custom-system-prompt switch, the system prompt is prepended as a strong
  instruction block (see "Persona neutralization").
- **Environment:** `CREATE_NO_WINDOW` on Windows; inherit env so `~/.codex` auth is
  found.

### Streaming

- Use Codex's JSONL event output mode when available; parse incremental text
  deltas the same way `claude_cli_stream()` parses `content_block_delta`.
- If JSONL streaming is unavailable on the installed version, buffer the final
  output and emit once. The chat tab + `generate()` already handle the
  non-streaming return shape (`iter([text])`).

### Persona neutralization

Codex's `exec` mode carries a coding-agent system persona. For general legal chat
this is mitigated by prepending an explicit instruction block to the combined
prompt, e.g.: "You are a general-purpose assistant. Do not write code, run
commands, or modify files unless explicitly asked. Answer the user's request
directly." This is a known quality risk (see "Risks").

### Enablement & model set

- Flag `openai_use_subscription` added to `config/llm_preferences.json` (a config
  file, so agent subprocesses read the same value — not QSettings). Default: enabled
  once Codex is detected + logged in; otherwise inert.
- A helper `openai_subscription_enabled()` reads the flag; `codex_available()`
  checks `codex` on PATH and login state.
- In subscription mode the model picker shows only **Codex-supported (gpt-5.x)**
  models. A mapping table translates existing app IDs to Codex model names, e.g.:
  - `gpt-5.2-thinking` / reasoning IDs → `gpt-5.2-codex` (high reasoning effort)
  - `gpt-5.2-instant` → `gpt-5.2-codex` (low/medium effort)
  - `gpt-4o`, `gpt-4o-mini`, `o1`, `o1-mini` → **not supported** → API-key fallback
- Reasoning effort maps from the app's existing `thinking_level` setting.

### Fallback & error handling

Fall back to the `OPENAI_API_KEY` path when ANY of:
- `codex` not installed / not on PATH
- not logged in (no valid `~/.codex` auth)
- the requested model is not Codex-supported
- the `codex exec` call errors or times out (300s, matching the Claude branch)

Every fallback is logged via `log_event(..., "warning")` with the reason.

## Components / Units

| Unit | Responsibility | Depends on |
|------|----------------|------------|
| `codex_available()` | Detect `codex` on PATH + login state | subprocess, `shutil.which` |
| `openai_subscription_enabled()` | Read `openai_use_subscription` from prefs | `llm_config` / json |
| `_map_openai_model_to_codex(model)` | App model ID → Codex model + effort, or `None` if unsupported | — |
| `_build_codex_command(model, effort)` | Assemble `codex exec` arg list | — |
| `_generate_openai_codex_cli(...)` | Run Codex, stream/return text | subprocess, above units |
| OpenAI branch hook in `generate()` | Try subscription, else API key | all above |

Each unit is independently unit-testable with a mocked subprocess.

## Testing

- **Unit (mock subprocess / filesystem):**
  - command construction (flags, model, neutral cwd)
  - model mapping (supported → codex id+effort; unsupported → None)
  - `codex_available()` true/false branches
  - `openai_subscription_enabled()` reads flag correctly
  - fallback triggers: not installed, not logged in, unsupported model, exec error
  - streaming parser handles JSONL deltas; non-streaming buffers correctly
- **Manual smoke (after `codex login`):**
  - one chat-tab message returns text and is billed to subscription
  - one agent run (e.g. Summarize) routes through Codex
  - force-fallback (e.g. select gpt-4o) confirms API-key path still works

Per project rules, run the test suite and verify manually before claiming done.

## Risks / Trade-offs (accepted)

1. **Prerequisite install + login.** Requires `npm i -g @openai/codex` and a
   one-time `codex login`. npm/node verified present.
2. **Codex coding persona.** `exec` runs with a coding-agent system prompt;
   neutralized via instruction injection, but general legal-chat responses may
   differ in tone/framing from raw-API GPT. Main quality risk.
3. **Limited model set.** Only gpt-5.x family on the subscription; gpt-4o/o1 fall
   back to the API key.
4. **Latency.** The agent harness is slower than a raw API call.
5. **Subscription rate limits.** Heavy batch agent runs may throttle; the API-key
   fallback covers this, but throttled calls cost API money once they fall back.
6. **Unofficial-for-third-parties.** OpenAI supports Codex login for its own tools;
   our use via the official CLI is ToS-cleaner than reverse-engineering the
   backend, but OpenAI could change behavior. Using the official CLI (vs. raw
   backend calls) minimizes breakage.

## Alternatives considered

- **B — Native OAuth (PKCE) + direct Codex backend calls.** Rejected: much more
  code, fragile to backend changes (the failure mode that broke Anthropic/Google
  third-party integrations), reuses OpenAI's client_id (grayer ToS).
- **Keep API key only.** Rejected: does not meet the goal (use the subscription).
```
