# Web Companion Chat Task Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a per-case AI Chat to the web companion — browse/create conversations (shared with the desktop chat file), send messages answered in the background with polling, pick the model, attach case files as context, and run legal research.

**Architecture:** Chat is an in-process LLM call, NOT a subprocess job, so it bypasses `JobManager`/`ScriptRunner`. A new `webcompanion/chat.py` wraps the existing `icharlotte_core.chat.persistence.ChatPersistence` and runs each turn on a background thread tracked in memory (`ChatTurnManager`), polled by the page. It reuses `LLMHandler.generate`, `ChatLegalResearchService`, and `DocumentProcessor.extract_text`.

**Tech Stack:** FastAPI + Jinja2 (existing webcompanion stack), threading, pytest. All deps already installed.

**Spec:** `docs/superpowers/specs/2026-06-13-webcompanion-chat-design.md`

---

## Investigation findings (verified against the codebase — trust these)

**ChatPersistence** (`icharlotte_core/chat/persistence.py`), constructed as
`ChatPersistence(file_number)`. Stores `{file_number}_chat.json` under the
module global `GEMINI_DATA_DIR` (imported into `persistence` as a module name —
tests redirect it with `monkeypatch.setattr(persistence, "GEMINI_DATA_DIR", str(tmp))`).
Methods used:
- `get_conversations() -> List[Conversation]`
- `get_conversation(conv_id) -> Optional[Conversation]`
- `create_conversation(name=None, provider='Gemini', model='gemini-3.5-flash', system_prompt='') -> str` (returns new id; inserts at front)
- `update_conversation(conv_id, **kwargs)`
- `add_message(conv_id, Message)`

**Message** (`icharlotte_core/chat/models.py`): dataclass
`Message(role='user', content='', ..., model_used=None, attachments=[...])`.
`Conversation.messages` is a `List[Message]`; each has `.role` ('user'|'assistant')
and `.content`. `Conversation` has `.id`, `.name`, `.provider`, `.model`,
`.system_prompt`, `.messages`.

**LLMHandler.generate** (`icharlotte_core/llm.py:302`):
`generate(provider, model, system_prompt, user_prompt, file_contents, settings, history=None, media_files=None)`.
With `settings={'stream': False}` returns the full reply **string**. `history` is
`[{'role': 'user'|'assistant', 'content': str}, ...]`. Import as
`from icharlotte_core.llm import LLMHandler`; tests monkeypatch
`chat.LLMHandler.generate`.

**Legal research** (`icharlotte_core/chat/legal_research.py`):
- `ChatLegalResearchService.from_environment(*, llm_callback, courtlistener_token=None) -> service`
- `service.research(*, user_text, context_text, settings, status_callback=None, debug_callback=None) -> ChatResearchPacket`
- `ChatResearchPacket.build_augmented_system_prompt(base_system_prompt) -> str`
- `ChatResearchSettings.default()` → sensible defaults.
- `ChatResearchError` is the failure type.
- `llm_callback(system_prompt, user_prompt) -> str` — wrap `LLMHandler.generate(..., stream=False, temperature=0.2)`.

**Document extraction** (`icharlotte_core/document_processor.py`):
`DocumentProcessor().extract_text(file_path, ocr_enabled=True) -> ExtractResult`;
`ExtractResult.text` is the string, `ExtractResult.error` is set on failure
(method never raises — returns a FAILED result). Use `ocr_enabled=False` for fast
attachment ingestion.

**Model list** (`icharlotte_core/llm_config.py`): module constants
`DEFAULT_MODEL_SEQUENCE` and `FAST_MODEL_SEQUENCE`, each a list of `ModelSpec`
with `.provider` and `.model`. Build the picker list from these.

**Default chat system prompt** (desktop `ChatTab`, verbatim):
`"You are a helpful legal assistant. Do not provide any disclaimers about being an AI or not being an attorney. Provide direct analysis only."`

**Existing webcompanion server** wires routes via `_register_*` functions called
at the end of `create_app` (server.py). Templates use `_render(name, request, **ctx)`
helpers calling `templates.TemplateResponse(request, name, ctx)`. `cases.get_case`,
`cases.browse`, `cases.safe_resolve` exist. The case page (`case.html`) renders
`T.TASKS` values as cards.

**Commit hygiene:** the working tree has unrelated uncommitted changes. Every
commit lists exact paths — `git add` ONLY those, never `git add -A`.

---

## File structure

```
webcompanion/
├── chat.py            # persistence wrappers + ChatTurnManager (background turns)
├── chat_models.py     # available_models() for the picker
├── server.py          # +_register_chat_routes (modify)
└── templates/
    ├── chat_conversations.html   # conversation list + New
    ├── chat_conversation.html    # messages + compose + poll script
    ├── chat_attach.html          # case-file picker for attachments
    └── case.html                 # + Chat card (modify)
tests/test_webcompanion/
├── test_chat.py        # chat service + turn manager
├── test_chat_models.py # available_models
└── test_server.py      # + chat endpoint tests (modify)
```

---

### Task 1: Chat persistence wrappers + default prompt

**Files:**
- Create: `webcompanion/chat.py`
- Test: `tests/test_webcompanion/test_chat.py`

- [ ] **Step 1: Write the failing tests**

```python
"""Tests for webcompanion.chat — persistence wrappers + turn manager."""
import pytest

from webcompanion import chat
from icharlotte_core.chat import persistence as _persistence


@pytest.fixture
def data_dir(tmp_path, monkeypatch):
    monkeypatch.setattr(_persistence, "GEMINI_DATA_DIR", str(tmp_path))
    return tmp_path


def test_default_system_prompt_constant():
    assert "helpful legal assistant" in chat.DEFAULT_SYSTEM_PROMPT
    assert "disclaimers" in chat.DEFAULT_SYSTEM_PROMPT


def test_create_and_list_conversation(data_dir):
    conv_id = chat.create_conversation("9999", name="Test thread")
    convs = chat.list_conversations("9999")
    assert len(convs) == 1
    assert convs[0].id == conv_id and convs[0].name == "Test thread"
    # new conversations get the default legal-assistant system prompt
    assert "helpful legal assistant" in convs[0].system_prompt


def test_create_defaults_provider_model(data_dir):
    conv_id = chat.create_conversation("9999")
    conv = chat.get_conversation("9999", conv_id)
    assert conv.provider == "Gemini" and conv.model == "gemini-3.5-flash"


def test_get_missing_conversation_returns_none(data_dir):
    assert chat.get_conversation("9999", "nope") is None


def test_append_message_persists(data_dir):
    conv_id = chat.create_conversation("9999")
    chat.append_message("9999", conv_id, role="user", content="hello")
    conv = chat.get_conversation("9999", conv_id)
    assert len(conv.messages) == 1
    assert conv.messages[0].role == "user" and conv.messages[0].content == "hello"


def test_append_assistant_message_records_model(data_dir):
    conv_id = chat.create_conversation("9999")
    chat.append_message("9999", conv_id, role="assistant", content="hi",
                        model_used="gemini-3.5-flash")
    conv = chat.get_conversation("9999", conv_id)
    assert conv.messages[0].model_used == "gemini-3.5-flash"
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'webcompanion.chat'`

- [ ] **Step 3: Implement (persistence wrappers only — turn manager comes in Task 2)**

`webcompanion/chat.py`:

```python
"""Chat service for the web companion.

Wraps the desktop ChatPersistence (per-case {file_number}_chat.json, shared
with the iCharlotte desktop chat) and — in Task 2 — runs chat turns on a
background thread. Chat is an in-process LLM call, NOT a subprocess job, so it
does not use JobManager/ScriptRunner.
"""
from typing import List, Optional

from icharlotte_core.chat.persistence import ChatPersistence
from icharlotte_core.chat.models import Conversation, Message

DEFAULT_SYSTEM_PROMPT = (
    "You are a helpful legal assistant. Do not provide any disclaimers about "
    "being an AI or not being an attorney. Provide direct analysis only."
)
DEFAULT_PROVIDER = "Gemini"
DEFAULT_MODEL = "gemini-3.5-flash"


def list_conversations(file_number: str) -> List[Conversation]:
    return ChatPersistence(file_number).get_conversations()


def get_conversation(file_number: str, conv_id: str) -> Optional[Conversation]:
    return ChatPersistence(file_number).get_conversation(conv_id)


def create_conversation(file_number: str, name: str = None,
                        provider: str = DEFAULT_PROVIDER,
                        model: str = DEFAULT_MODEL) -> str:
    return ChatPersistence(file_number).create_conversation(
        name=name, provider=provider, model=model,
        system_prompt=DEFAULT_SYSTEM_PROMPT)


def append_message(file_number: str, conv_id: str, *, role: str, content: str,
                   model_used: str = None) -> None:
    msg = Message(role=role, content=content, model_used=model_used)
    ChatPersistence(file_number).add_message(conv_id, msg)
```

- [ ] **Step 4: Run to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```bash
git add webcompanion/chat.py tests/test_webcompanion/test_chat.py
git commit -m "feat(webcompanion): chat persistence wrappers"
```

---

### Task 2: ChatTurnManager (background reply, core generate path)

**Files:**
- Modify: `webcompanion/chat.py`
- Test: `tests/test_webcompanion/test_chat.py` (append)

The turn manager runs the full pipeline structure now, but the **extract** and
**research** steps are stubs (`_extract_context` returns `""`,
`_augmented_system_prompt` returns the base prompt). Tasks 5 and 6 implement
those stubs. This keeps `start_turn`'s signature stable across phases.

- [ ] **Step 1: Write the failing tests** (append to test_chat.py)

```python
import time

from webcompanion import chat as chatmod


def _wait_turn(mgr, turn_id, status, timeout=5.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        t = mgr.get_turn(turn_id)
        if t and t["status"] == status:
            return t
        time.sleep(0.02)
    raise AssertionError(f"turn did not reach {status}: {mgr.get_turn(turn_id)}")


def test_turn_generates_and_persists(data_dir, monkeypatch):
    monkeypatch.setattr(chatmod.LLMHandler, "generate",
                        staticmethod(lambda **kw: "the answer"))
    conv_id = chat.create_conversation("9999")
    mgr = chatmod.ChatTurnManager()
    turn_id = mgr.start_turn("9999", conv_id, user_text="what is the law?",
                             provider="Gemini", model="gemini-3.5-flash",
                             attach_rel_files=[], research_on=False)
    # user message persisted immediately
    conv = chat.get_conversation("9999", conv_id)
    assert conv.messages[0].role == "user"
    _wait_turn(mgr, turn_id, "done")
    conv = chat.get_conversation("9999", conv_id)
    assert [m.role for m in conv.messages] == ["user", "assistant"]
    assert conv.messages[1].content == "the answer"
    assert conv.messages[1].model_used == "gemini-3.5-flash"


def test_turn_passes_history_excluding_current(data_dir, monkeypatch):
    captured = {}
    def fake_generate(**kw):
        captured.update(kw)
        return "reply2"
    monkeypatch.setattr(chatmod.LLMHandler, "generate", staticmethod(fake_generate))
    conv_id = chat.create_conversation("9999")
    chat.append_message("9999", conv_id, role="user", content="first")
    chat.append_message("9999", conv_id, role="assistant", content="answer1")
    mgr = chatmod.ChatTurnManager()
    turn_id = mgr.start_turn("9999", conv_id, user_text="second",
                             provider="Gemini", model="gemini-3.5-flash",
                             attach_rel_files=[], research_on=False)
    _wait_turn(mgr, turn_id, "done")
    # history is the prior 2 messages; current user_text passed separately
    assert captured["history"] == [
        {"role": "user", "content": "first"},
        {"role": "assistant", "content": "answer1"},
    ]
    assert captured["user_prompt"] == "second"


def test_turn_generate_failure_marks_failed(data_dir, monkeypatch):
    def boom(**kw):
        raise RuntimeError("llm down")
    monkeypatch.setattr(chatmod.LLMHandler, "generate", staticmethod(boom))
    conv_id = chat.create_conversation("9999")
    mgr = chatmod.ChatTurnManager()
    turn_id = mgr.start_turn("9999", conv_id, user_text="x",
                             provider="Gemini", model="gemini-3.5-flash",
                             attach_rel_files=[], research_on=False)
    t = _wait_turn(mgr, turn_id, "failed")
    assert "llm down" in t["error"]
    # user message stays; no assistant message appended
    conv = chat.get_conversation("9999", conv_id)
    assert [m.role for m in conv.messages] == ["user"]


def test_one_turn_per_conversation(data_dir, monkeypatch):
    monkeypatch.setattr(chatmod.LLMHandler, "generate",
                        staticmethod(lambda **kw: (time.sleep(0.3) or "slow")))
    conv_id = chat.create_conversation("9999")
    mgr = chatmod.ChatTurnManager()
    mgr.start_turn("9999", conv_id, user_text="a", provider="Gemini",
                   model="gemini-3.5-flash", attach_rel_files=[], research_on=False)
    with pytest.raises(ValueError):
        mgr.start_turn("9999", conv_id, user_text="b", provider="Gemini",
                       model="gemini-3.5-flash", attach_rel_files=[],
                       research_on=False)
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v -k turn`
Expected: FAIL — `AttributeError`/`ModuleNotFoundError` for `ChatTurnManager` / `LLMHandler`

- [ ] **Step 3: Implement — add to `webcompanion/chat.py`**

Add the import near the top (below the existing imports):

```python
import threading
import uuid
from typing import Dict

from icharlotte_core.llm import LLMHandler
```

Append the turn manager:

```python
# --- Background turn manager -------------------------------------------------

_CONTEXT_CHAR_CAP = 100_000  # matches desktop research-context cap


def _history_for(conv) -> list:
    """Prior messages as [{'role','content'}], excluding the just-added user msg."""
    msgs = list(conv.messages)
    if msgs and msgs[-1].role == "user":
        msgs = msgs[:-1]
    return [{"role": m.role, "content": m.content} for m in msgs]


class ChatTurnManager:
    """Runs chat turns on background threads, tracked in memory.

    A 'turn' = one user message -> assistant reply. The conversation itself is
    persisted via ChatPersistence; only the in-flight turn state lives here.
    Statuses: 'extracting' -> 'researching' -> 'generating' -> 'done'|'failed'.
    """

    def __init__(self, max_concurrent: int = 2):
        self._max = max_concurrent
        self._lock = threading.RLock()
        self._turns: Dict[str, dict] = {}
        self._busy_convs: set = set()

    def get_turn(self, turn_id: str) -> dict | None:
        with self._lock:
            t = self._turns.get(turn_id)
            return dict(t) if t else None

    def start_turn(self, file_number: str, conv_id: str, *, user_text: str,
                   provider: str, model: str, attach_rel_files: list,
                   research_on: bool) -> str:
        conv = get_conversation(file_number, conv_id)
        if conv is None:
            raise ValueError("Unknown conversation.")
        with self._lock:
            if conv_id in self._busy_convs:
                raise ValueError("A reply is already in progress for this thread.")
            running = sum(1 for t in self._turns.values()
                          if t["status"] not in ("done", "failed"))
            if running >= self._max:
                raise ValueError("Too many chats in progress; try again shortly.")
            self._busy_convs.add(conv_id)
            turn_id = uuid.uuid4().hex[:12]
            self._turns[turn_id] = {
                "id": turn_id, "conv_id": conv_id, "status": "generating",
                "log": [], "error": "", "done": False,
            }
        # Persist the user message immediately so it shows while we work.
        append_message(file_number, conv_id, role="user", content=user_text)
        threading.Thread(
            target=self._run, daemon=True,
            args=(turn_id, file_number, conv_id, user_text, provider, model,
                  attach_rel_files, research_on),
        ).start()
        return turn_id

    # ---- internals ----

    def _set(self, turn_id: str, **kw) -> None:
        with self._lock:
            t = self._turns.get(turn_id)
            if t:
                t.update(kw)

    def _log(self, turn_id: str, line: str) -> None:
        with self._lock:
            t = self._turns.get(turn_id)
            if t:
                t["log"].append(line)

    def _run(self, turn_id, file_number, conv_id, user_text, provider, model,
             attach_rel_files, research_on) -> None:
        try:
            # Only touch the master DB to resolve the case folder when there
            # are attachments to extract (keeps the no-attachment path DB-free).
            case_root = _case_root(file_number) if attach_rel_files else ""
            self._set(turn_id, status="extracting")
            context_text = self._extract_context(turn_id, case_root, attach_rel_files)

            self._set(turn_id, status="researching")
            system_prompt = self._augmented_system_prompt(
                turn_id, provider, model, user_text, context_text, research_on)

            self._set(turn_id, status="generating")
            conv = get_conversation(file_number, conv_id)
            reply = LLMHandler.generate(
                provider=provider, model=model, system_prompt=system_prompt,
                user_prompt=user_text, file_contents=context_text,
                settings={"stream": False}, history=_history_for(conv))
            append_message(file_number, conv_id, role="assistant",
                           content=reply or "", model_used=model)
            self._set(turn_id, status="done", done=True)
        except Exception as exc:  # noqa: BLE001 — surface any failure to the UI
            self._log(turn_id, f"Error: {exc}")
            self._set(turn_id, status="failed", error=str(exc), done=True)
        finally:
            with self._lock:
                self._busy_convs.discard(conv_id)

    # Stubs — implemented in Tasks 5 (extract) and 6 (research).
    def _extract_context(self, turn_id, case_root, attach_rel_files) -> str:
        return ""

    def _augmented_system_prompt(self, turn_id, provider, model, user_text,
                                 context_text, research_on) -> str:
        return DEFAULT_SYSTEM_PROMPT
```

Add a `_case_root` helper near the persistence wrappers (used by `_run`):

```python
def _case_root(file_number: str) -> str:
    """Case folder path from the master DB (for resolving attachments)."""
    from .cases import get_case
    case = get_case(file_number)
    return case["case_path"] if case else ""
```

NOTE: `_augmented_system_prompt` returns `DEFAULT_SYSTEM_PROMPT` rather than the
conversation's stored `system_prompt`; new conversations use the default and
Task 6 builds on this. If you prefer to honor a per-conversation custom prompt,
that is a later enhancement — keep the stub as written for now.

- [ ] **Step 4: Run to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v`
Expected: 10 passed (6 from Task 1 + 4 new)

- [ ] **Step 5: Commit**

```bash
git add webcompanion/chat.py tests/test_webcompanion/test_chat.py
git commit -m "feat(webcompanion): background chat turn manager"
```

---

### Task 3: Chat routes + templates + Chat card (core, plain chat end to end)

**Files:**
- Modify: `webcompanion/server.py`
- Create: `webcompanion/templates/chat_conversations.html`, `webcompanion/templates/chat_conversation.html`
- Modify: `webcompanion/templates/case.html`
- Test: `tests/test_webcompanion/test_server.py` (append)

- [ ] **Step 1: Write the failing tests** (append to test_server.py)

```python
from webcompanion import chat as chat_mod


def test_case_page_shows_chat_card(client):
    r = client.get("/case/9999")
    assert r.status_code == 200
    assert "/case/9999/chat" in r.text and "Chat" in r.text


def test_chat_conversations_list(client, monkeypatch):
    class _Conv:
        id = "c1"; name = "Thread 1"
    monkeypatch.setattr(chat_mod, "list_conversations", lambda fn: [_Conv()])
    r = client.get("/case/9999/chat")
    assert r.status_code == 200 and "Thread 1" in r.text


def test_chat_new_creates_and_redirects(client, monkeypatch):
    monkeypatch.setattr(chat_mod, "create_conversation", lambda fn, **kw: "newid")
    r = client.post("/case/9999/chat/new", follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"] == "/case/9999/chat/newid"


def test_chat_conversation_view(client, monkeypatch):
    class _Msg:
        def __init__(self, role, content):
            self.role, self.content, self.attachments, self.model_used = role, content, [], None
    class _Conv:
        id = "c1"; name = "T"; provider = "Gemini"; model = "gemini-3.5-flash"
        messages = [_Msg("user", "hi"), _Msg("assistant", "hello there")]
    monkeypatch.setattr(chat_mod, "get_conversation", lambda fn, cid: _Conv())
    r = client.get("/case/9999/chat/c1")
    assert r.status_code == 200
    assert "hi" in r.text and "hello there" in r.text


def test_chat_conversation_404(client, monkeypatch):
    monkeypatch.setattr(chat_mod, "get_conversation", lambda fn, cid: None)
    assert client.get("/case/9999/chat/nope").status_code == 404


def test_chat_send_starts_turn_and_redirects(client, monkeypatch):
    class _Conv:
        id = "c1"; name = "T"; provider = "Gemini"; model = "gemini-3.5-flash"
        messages = []
    monkeypatch.setattr(chat_mod, "get_conversation", lambda fn, cid: _Conv())
    started = {}
    def fake_start(self, fn, cid, **kw):
        started.update(kw); started["cid"] = cid; return "turn1"
    monkeypatch.setattr(chat_mod.ChatTurnManager, "start_turn", fake_start)
    r = client.post("/case/9999/chat/c1/send",
                    data={"message": "what is the law?"}, follow_redirects=False)
    assert r.status_code == 303
    assert "/case/9999/chat/c1?turn=turn1" in r.headers["location"]
    assert started["user_text"] == "what is the law?"
    assert started["research_on"] is False


def test_chat_send_empty_message_400(client, monkeypatch):
    class _Conv:
        id = "c1"; messages = []; provider = "Gemini"; model = "gemini-3.5-flash"; name = "T"
    monkeypatch.setattr(chat_mod, "get_conversation", lambda fn, cid: _Conv())
    r = client.post("/case/9999/chat/c1/send", data={"message": "  "})
    assert r.status_code == 400


def test_chat_turn_status_api(client, monkeypatch):
    monkeypatch.setattr(chat_mod.ChatTurnManager, "get_turn",
                        lambda self, tid: {"status": "generating", "log": ["x"],
                                           "done": False, "error": ""})
    r = client.get("/api/chat/c1/turn/turn1")
    body = r.json()
    assert body["status"] == "generating" and body["done"] is False
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v -k chat`
Expected: FAIL (404s / missing card)

- [ ] **Step 3: Add the Chat card to `case.html`**

In `webcompanion/templates/case.html`, add this block immediately AFTER the
`{% for task in tasks %}...{% endfor %}` loop (before `{% endblock %}`):

```html
<div class="card"><a href="/case/{{ case.file_number }}/chat">
 &#128172; <b>Chat</b><br>
 <small>Ask an AI about this case; attach files or run legal research.</small>
</a></div>
```

- [ ] **Step 4: Create `webcompanion/templates/chat_conversations.html`**

```html
{% extends "base.html" %}
{% block content %}
<h1>&#128172; Chat &mdash; {{ case.file_number }}</h1>
<form method="post" action="/case/{{ case.file_number }}/chat/new">
 <button class="btn" type="submit">+ New conversation</button>
</form>
{% for conv in conversations %}
<div class="card"><a href="/case/{{ case.file_number }}/chat/{{ conv.id }}">
 <b>{{ conv.name }}</b>
</a></div>
{% else %}
<p>No conversations yet.</p>
{% endfor %}
{% endblock %}
```

- [ ] **Step 5: Create `webcompanion/templates/chat_conversation.html`**

```html
{% extends "base.html" %}
{% block content %}
<div class="crumb"><a href="/case/{{ case.file_number }}/chat">&#8592; Conversations</a></div>
<h1>{{ conv.name }}</h1>
{% for m in conv.messages %}
<div class="card">
 <b>{{ 'You' if m.role == 'user' else 'AI' }}:</b>
 <div style="white-space:pre-wrap">{{ m.content }}</div>
 {% for a in m.attachments %}<small>&#128206; {{ a.name }}</small>{% endfor %}
</div>
{% endfor %}
{% if turn_id %}
<div class="card" id="pending">
 <b>AI:</b> <span id="turnstatus">thinking&hellip;</span>
 <pre id="turnlog"></pre>
</div>
<script>
const timer = setInterval(async () => {
  let r;
  try { r = await fetch('/api/chat/{{ conv.id }}/turn/{{ turn_id }}'); } catch(e){ return; }
  if(!r.ok) return;
  const s = await r.json();
  document.getElementById('turnstatus').textContent = s.status;
  document.getElementById('turnlog').textContent = (s.log || []).join('\n');
  if(s.done){ clearInterval(timer); location.href = '/case/{{ case.file_number }}/chat/{{ conv.id }}'; }
}, 2500);
</script>
{% endif %}
<form method="post" action="/case/{{ case.file_number }}/chat/{{ conv.id }}/send" class="card">
 <textarea name="message" rows="4" placeholder="Ask about this case..." {% if turn_id %}disabled{% endif %}></textarea>
 <button class="btn" type="submit" {% if turn_id %}disabled{% endif %}>Send</button>
</form>
{% endblock %}
```

- [ ] **Step 6: Implement `_register_chat_routes` and wire it in `server.py`**

Add a module-level chat-manager singleton and import near the top of server.py
(after the existing `from . import ...` lines):

```python
from . import chat
_CHAT_MANAGER = chat.ChatTurnManager()
```

Inside `create_app`, add a call alongside the other registrations (before
`return app`):

```python
    _register_chat_routes(app, templates)
```

Add the function (place it near the other `_register_*` functions):

```python
def _register_chat_routes(app, templates):
    def _render(name, request, **ctx):
        return templates.TemplateResponse(request, name, ctx)

    @app.get("/case/{file_number}/chat", response_class=HTMLResponse)
    def chat_list(request: Request, file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        return _render("chat_conversations.html", request, case=case,
                       conversations=chat.list_conversations(file_number))

    @app.post("/case/{file_number}/chat/new")
    def chat_new(file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        conv_id = chat.create_conversation(file_number)
        return RedirectResponse(f"/case/{file_number}/chat/{conv_id}",
                                status_code=303)

    @app.get("/case/{file_number}/chat/{conv_id}", response_class=HTMLResponse)
    def chat_view(request: Request, file_number: str, conv_id: str,
                  turn: str = None):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        conv = chat.get_conversation(file_number, conv_id)
        if conv is None:
            return HTMLResponse("Conversation not found", status_code=404)
        return _render("chat_conversation.html", request, case=case, conv=conv,
                       turn_id=turn)

    @app.post("/case/{file_number}/chat/{conv_id}/send")
    async def chat_send(request: Request, file_number: str, conv_id: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        conv = chat.get_conversation(file_number, conv_id)
        if conv is None:
            return HTMLResponse("Conversation not found", status_code=404)
        form = await request.form()
        message = (form.get("message") or "").strip()
        if not message:
            return HTMLResponse("Type a message.", status_code=400)
        provider = form.get("provider") or conv.provider
        model = form.get("model") or conv.model
        rel_files = [f for f in form.getlist("attach") if f]
        research_on = form.get("research") == "on"
        try:
            turn_id = _CHAT_MANAGER.start_turn(
                file_number, conv_id, user_text=message, provider=provider,
                model=model, attach_rel_files=rel_files, research_on=research_on)
        except ValueError as exc:
            return HTMLResponse(str(exc), status_code=409)
        return RedirectResponse(
            f"/case/{file_number}/chat/{conv_id}?turn={turn_id}",
            status_code=303)

    @app.get("/api/chat/{conv_id}/turn/{turn_id}")
    def chat_turn_status(conv_id: str, turn_id: str):
        t = _CHAT_MANAGER.get_turn(turn_id)
        if t is None:
            return JSONResponse({"error": "not found"}, status_code=404)
        return JSONResponse({"status": t["status"], "log": t["log"],
                             "done": t["done"], "error": t["error"]})
```

- [ ] **Step 7: Run chat endpoint tests, then full suite**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v -k chat`
Expected: 8 passed
Run: `python -m pytest tests/test_webcompanion/ -q`
Expected: all pass

- [ ] **Step 8: Commit**

```bash
git add webcompanion/server.py webcompanion/templates/chat_conversations.html webcompanion/templates/chat_conversation.html webcompanion/templates/case.html tests/test_webcompanion/test_server.py
git commit -m "feat(webcompanion): chat conversation routes, templates, and Chat card"
```

---

### Task 4: Model picker

**Files:**
- Create: `webcompanion/chat_models.py`
- Test: `tests/test_webcompanion/test_chat_models.py`
- Modify: `webcompanion/server.py` (pass models to chat_view), `webcompanion/templates/chat_conversation.html` (add `<select>`)

- [ ] **Step 1: Write the failing test**

```python
"""Tests for webcompanion.chat_models."""
from webcompanion import chat_models


def test_available_models_nonempty_and_shaped():
    models = chat_models.available_models()
    assert len(models) >= 2
    for m in models:
        assert set(m) >= {"provider", "model", "label"}
    # includes the chat default
    assert any(m["model"] == "gemini-3.5-flash" for m in models)


def test_available_models_deduped():
    models = chat_models.available_models()
    pairs = [(m["provider"], m["model"]) for m in models]
    assert len(pairs) == len(set(pairs))
```

Also append this save-back test to `tests/test_webcompanion/test_chat.py`
(it reuses the `data_dir` fixture already defined there):

```python
def test_persist_model_choice_updates_conversation(data_dir):
    conv_id = chat.create_conversation("9999")  # defaults Gemini/flash
    chat._persist_model_choice("9999", conv_id, "Claude",
                               "claude-sonnet-4-20250514")
    conv = chat.get_conversation("9999", conv_id)
    assert conv.provider == "Claude" and conv.model == "claude-sonnet-4-20250514"
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_webcompanion/test_chat_models.py tests/test_webcompanion/test_chat.py -v -k "models or persist_model"`
Expected: FAIL — `ModuleNotFoundError` (chat_models) and `AttributeError` (`_persist_model_choice`)

- [ ] **Step 3: Implement `webcompanion/chat_models.py`**

```python
"""Curated provider/model list for the chat compose picker.

Derived from the real iCharlotte model sequences so the phone offers the same
models the desktop is configured for.
"""
from icharlotte_core.llm_config import DEFAULT_MODEL_SEQUENCE, FAST_MODEL_SEQUENCE


def available_models() -> list:
    """Return [{'provider','model','label'}], de-duplicated, order preserved."""
    out, seen = [], set()
    for spec in list(DEFAULT_MODEL_SEQUENCE) + list(FAST_MODEL_SEQUENCE):
        key = (spec.provider, spec.model)
        if key in seen:
            continue
        seen.add(key)
        out.append({"provider": spec.provider, "model": spec.model,
                    "label": f"{spec.provider} {spec.model}"})
    return out
```

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_webcompanion/test_chat_models.py -v`
Expected: 2 passed

- [ ] **Step 5: Pass models into the conversation view**

In `server.py` `chat_view`, change the render call to include models and the
current selection:

```python
        from . import chat_models
        return _render("chat_conversation.html", request, case=case, conv=conv,
                       turn_id=turn, models=chat_models.available_models())
```

- [ ] **Step 6: Add the picker to `chat_conversation.html`**

Inside the compose `<form>`, BEFORE the textarea, add:

```html
 <label>Model
  <select name="model" {% if turn_id %}disabled{% endif %}>
  {% for m in models %}
   <option value="{{ m.model }}" data-provider="{{ m.provider }}"
     {% if m.model == conv.model %}selected{% endif %}>{{ m.label }}</option>
  {% endfor %}
  </select></label>
 <input type="hidden" name="provider" value="{{ conv.provider }}">
```

NOTE: the `model` field is sent; `provider` defaults to the conversation's
provider via the hidden field. The send route already saves neither back —
add save-back now: in `server.py` `chat_send`, immediately AFTER computing
`provider`/`model` and BEFORE `start_turn`, persist the choice so the thread
remembers it:

```python
        from .chat import _persist_model_choice
        _persist_model_choice(file_number, conv_id, provider, model)
```

And add that helper to `webcompanion/chat.py`:

```python
def _persist_model_choice(file_number: str, conv_id: str, provider: str,
                          model: str) -> None:
    ChatPersistence(file_number).update_conversation(
        conv_id, provider=provider, model=model)
```

- [ ] **Step 7: Run full suite**

Run: `python -m pytest tests/test_webcompanion/ -q`
Expected: all pass

- [ ] **Step 8: Commit**

```bash
git add webcompanion/chat_models.py webcompanion/chat.py webcompanion/server.py webcompanion/templates/chat_conversation.html tests/test_webcompanion/test_chat_models.py
git commit -m "feat(webcompanion): chat model picker with save-back"
```

---

### Task 5: Attach case files as context

**Files:**
- Modify: `webcompanion/chat.py` (implement `_extract_context`)
- Modify: `webcompanion/server.py` (attach picker route)
- Create: `webcompanion/templates/chat_attach.html`
- Modify: `webcompanion/templates/chat_conversation.html` (attach link + carry selections)
- Test: `tests/test_webcompanion/test_chat.py` (append)

- [ ] **Step 1: Write the failing tests** (append to test_chat.py)

```python
def test_extract_context_concatenates_and_caps(data_dir, monkeypatch, tmp_path):
    f1 = tmp_path / "a.txt"; f1.write_text("AAA", encoding="utf-8")
    f2 = tmp_path / "b.txt"; f2.write_text("BBB", encoding="utf-8")

    class _Res:
        def __init__(self, text): self.text = text; self.error = ""
    monkeypatch.setattr(chatmod.DocumentProcessor, "extract_text",
                        lambda self, p, ocr_enabled=True: _Res(open(p, encoding="utf-8").read()))
    mgr = chatmod.ChatTurnManager()
    text = mgr._extract_context("t", str(tmp_path), ["a.txt", "b.txt"])
    assert "AAA" in text and "BBB" in text
    assert len(text) <= chatmod._CONTEXT_CHAR_CAP


def test_extract_context_skips_failed_file(data_dir, monkeypatch, tmp_path):
    (tmp_path / "ok.txt").write_text("OK", encoding="utf-8")
    class _Res:
        def __init__(self, text, error=""): self.text = text; self.error = error
    def fake_extract(self, p, ocr_enabled=True):
        if p.endswith("bad.txt"):
            return _Res("", "boom")
        return _Res("OK")
    monkeypatch.setattr(chatmod.DocumentProcessor, "extract_text", fake_extract)
    mgr = chatmod.ChatTurnManager()
    text = mgr._extract_context("t", str(tmp_path), ["ok.txt", "bad.txt"])
    assert "OK" in text  # bad file skipped, no raise


def test_extract_context_empty_when_none(data_dir):
    mgr = chatmod.ChatTurnManager()
    assert mgr._extract_context("t", "/case", []) == ""
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v -k extract`
Expected: FAIL (`_extract_context` returns "" stub / `DocumentProcessor` undefined)

- [ ] **Step 3: Implement in `webcompanion/chat.py`**

Add the import near the other icharlotte imports:

```python
from icharlotte_core.document_processor import DocumentProcessor
```

Add a safe-resolve import at the top of the file (with the other `from .`):

```python
from .cases import safe_resolve
```

Replace the `_extract_context` stub with:

```python
    def _extract_context(self, turn_id, case_root, attach_rel_files) -> str:
        if not attach_rel_files:
            return ""
        processor = DocumentProcessor()
        parts: list = []
        total = 0
        for rel in attach_rel_files:
            try:
                abs_path = str(safe_resolve(case_root, rel))
            except ValueError:
                self._log(turn_id, f"Skipped (invalid path): {rel}")
                continue
            result = processor.extract_text(abs_path, ocr_enabled=False)
            if getattr(result, "error", "") or not result.text:
                self._log(turn_id, f"Skipped (no text): {rel}")
                continue
            chunk = f"--- {rel} ---\n{result.text}"
            parts.append(chunk)
            total += len(chunk)
            self._log(turn_id, f"Attached: {rel}")
            if total >= _CONTEXT_CHAR_CAP:
                break
        return "\n\n".join(parts)[:_CONTEXT_CHAR_CAP]
```

- [ ] **Step 4: Run to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v -k extract`
Expected: 3 passed

- [ ] **Step 5: Add the attach picker route in `server.py`**

Inside `_register_chat_routes`, add:

```python
    @app.get("/case/{file_number}/chat/{conv_id}/attach", response_class=HTMLResponse)
    def chat_attach(request: Request, file_number: str, conv_id: str,
                    path: str = None):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        rel = path or ""
        try:
            dirs, files = cases.browse(case["case_path"], rel,
                                       (".pdf", ".docx", ".doc", ".txt"))
        except ValueError:
            return HTMLResponse("Invalid path", status_code=400)
        return _render("chat_attach.html", request, case=case, conv_id=conv_id,
                       path=rel, dirs=dirs, files=files)
```

- [ ] **Step 6: Create `webcompanion/templates/chat_attach.html`**

```html
{% extends "base.html" %}
{% block content %}
<h1>Attach files</h1>
<div class="crumb"><a href="/case/{{ case.file_number }}/chat/{{ conv_id }}/attach?path=">root</a></div>
{% for d in dirs %}
<div class="card"><a href="/case/{{ case.file_number }}/chat/{{ conv_id }}/attach?path={{ ((path ~ '/' ~ d) if path else d) | urlencode }}">&#128193; {{ d }}</a></div>
{% endfor %}
<form method="post" action="/case/{{ case.file_number }}/chat/{{ conv_id }}/send">
{% for f in files %}
<div class="card"><label class="row">
 <input type="checkbox" name="attach" value="{{ (path ~ '/' ~ f) if path else f }}"> {{ f }}
</label></div>
{% endfor %}
<label>Message <textarea name="message" rows="3" placeholder="Ask about these files..."></textarea></label>
<button class="btn" type="submit">Send with attachments</button>
</form>
{% endblock %}
```

- [ ] **Step 7: Add the attach link to `chat_conversation.html`**

Inside the compose `<form>`, AFTER the Send button, add:

```html
 <a href="/case/{{ case.file_number }}/chat/{{ conv.id }}/attach">&#128206; Attach files instead</a>
```

- [ ] **Step 8: Run full suite**

Run: `python -m pytest tests/test_webcompanion/ -q`
Expected: all pass

- [ ] **Step 9: Commit**

```bash
git add webcompanion/chat.py webcompanion/server.py webcompanion/templates/chat_attach.html webcompanion/templates/chat_conversation.html tests/test_webcompanion/test_chat.py
git commit -m "feat(webcompanion): attach case files as chat context"
```

---

### Task 6: Legal research toggle

**Files:**
- Modify: `webcompanion/chat.py` (implement `_augmented_system_prompt`)
- Modify: `webcompanion/templates/chat_conversation.html` (research checkbox)
- Test: `tests/test_webcompanion/test_chat.py` (append)

- [ ] **Step 1: Write the failing tests** (append to test_chat.py)

```python
def test_research_off_returns_base_prompt(data_dir):
    mgr = chatmod.ChatTurnManager()
    out = mgr._augmented_system_prompt("t", "Gemini", "m", "q", "", False)
    assert out == chatmod.DEFAULT_SYSTEM_PROMPT


def test_research_on_augments_prompt(data_dir, monkeypatch):
    class _Packet:
        def build_augmented_system_prompt(self, base):
            return base + "\n\n[AUTHORITY]"
    class _Service:
        @classmethod
        def from_environment(cls, *, llm_callback, courtlistener_token=None):
            return cls()
        def research(self, *, user_text, context_text, settings,
                     status_callback=None, debug_callback=None):
            if status_callback:
                status_callback("searching authority")
            return _Packet()
    monkeypatch.setattr(chatmod, "ChatLegalResearchService", _Service)
    mgr = chatmod.ChatTurnManager()
    out = mgr._augmented_system_prompt("t", "Gemini", "m", "q", "ctx", True)
    assert out.endswith("[AUTHORITY]")


def test_research_failure_falls_back_to_base(data_dir, monkeypatch):
    class _Service:
        @classmethod
        def from_environment(cls, *, llm_callback, courtlistener_token=None):
            return cls()
        def research(self, **kw):
            raise RuntimeError("research broke")
    monkeypatch.setattr(chatmod, "ChatLegalResearchService", _Service)
    mgr = chatmod.ChatTurnManager()
    out = mgr._augmented_system_prompt("t", "Gemini", "m", "q", "ctx", True)
    assert out == chatmod.DEFAULT_SYSTEM_PROMPT
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v -k research`
Expected: FAIL (`_augmented_system_prompt` ignores research / `ChatLegalResearchService` undefined)

- [ ] **Step 3: Implement in `webcompanion/chat.py`**

Add the imports near the other icharlotte imports:

```python
from icharlotte_core.chat.legal_research import (
    ChatLegalResearchService, ChatResearchSettings)
```

Replace the `_augmented_system_prompt` stub with:

```python
    def _augmented_system_prompt(self, turn_id, provider, model, user_text,
                                 context_text, research_on) -> str:
        if not research_on:
            return DEFAULT_SYSTEM_PROMPT

        def llm_callback(system_prompt, user_prompt):
            return LLMHandler.generate(
                provider=provider, model=model, system_prompt=system_prompt,
                user_prompt=user_prompt, file_contents="",
                settings={"stream": False, "temperature": 0.2})

        try:
            service = ChatLegalResearchService.from_environment(
                llm_callback=llm_callback)
            packet = service.research(
                user_text=user_text,
                context_text=context_text[:_CONTEXT_CHAR_CAP] if context_text else "",
                settings=ChatResearchSettings.default(),
                status_callback=lambda msg: self._log(turn_id, str(msg)))
            return packet.build_augmented_system_prompt(DEFAULT_SYSTEM_PROMPT)
        except Exception as exc:  # noqa: BLE001 — research is best-effort
            self._log(turn_id, f"Legal research failed ({exc}); answering without it.")
            return DEFAULT_SYSTEM_PROMPT
```

- [ ] **Step 4: Run to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_chat.py -v -k research`
Expected: 3 passed

- [ ] **Step 5: Add the research toggle to `chat_conversation.html`**

Inside the compose `<form>`, AFTER the textarea and BEFORE the Send button, add:

```html
 <label class="row"><input type="checkbox" name="research" {% if turn_id %}disabled{% endif %}> Run legal research</label>
```

- [ ] **Step 6: Run full suite**

Run: `python -m pytest tests/test_webcompanion/ -q`
Expected: all pass

- [ ] **Step 7: Commit**

```bash
git add webcompanion/chat.py webcompanion/templates/chat_conversation.html tests/test_webcompanion/test_chat.py
git commit -m "feat(webcompanion): legal research toggle in chat"
```

---

### Task 7: Docs + manual verification

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: Document in CLAUDE.md** — append under the "Wizard Web Companion" entry in "Recent Features":

```markdown
### Web Companion Chat (2026-06-13)
- Per-case AI chat from the phone: conversation management (shared with the
  desktop chat file), background replies with polling, model picker, file
  attachments as context, and a legal-research toggle
- In-process LLM call (not a job); `webcompanion/chat.py` + `_register_chat_routes`
- Spec: `docs/superpowers/specs/2026-06-13-webcompanion-chat-design.md`
```

- [ ] **Step 2: Commit**

```bash
git add CLAUDE.md
git commit -m "docs(webcompanion): note chat task"
```

- [ ] **Step 3: Manual verification (MANDATORY — per the global test-after-develop rule)**

Start the server (`run_webcompanion.bat`), then from the iPhone over Tailscale:
- Open a case → **Chat** → conversation list loads
- **New conversation** → send a plain message → "thinking…" → reply appears via poll
- Confirm the same conversation + messages appear in the desktop iCharlotte chat for that case
- Change the **model** in the picker → send → reply uses it (and the thread remembers the choice on reload)
- **Attach files** → pick a small case PDF → ask about it → reply reflects the content
- Toggle **Run legal research** on → send → watch the status log advance → reply includes authority
- Fix anything found, re-run `python -m pytest tests/test_webcompanion/ -q`, commit fixes as `fix(webcompanion): …`

---

## Notes for the implementer

- **Monkeypatch targets:** tests patch names as imported into `webcompanion.chat`
  (`chat.LLMHandler`, `chat.DocumentProcessor`, `chat.ChatLegalResearchService`).
  So import those as module-level names in `chat.py` exactly as the tasks show —
  do not call them via fully-qualified paths.
- **Thread state in tests:** the `_wait_turn` helper polls `get_turn`; turns run
  on daemon threads. Generous timeouts are already set.
- **No `git add -A`** anywhere — the working tree has unrelated changes; stage
  only the listed paths.
- Run the FULL `tests/test_webcompanion/` suite at the end of each task to catch
  cross-file regressions (server tests import the chat module).
