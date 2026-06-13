"""Chat service for the web companion.

Wraps the desktop ChatPersistence (per-case {file_number}_chat.json, shared
with the iCharlotte desktop chat) and — in Task 2 — runs chat turns on a
background thread. Chat is an in-process LLM call, NOT a subprocess job, so it
does not use JobManager/ScriptRunner.
"""
import threading
import uuid
from typing import Dict, List, Optional

from icharlotte_core.chat.persistence import ChatPersistence
from icharlotte_core.chat.models import Conversation, Message
from icharlotte_core.llm import LLMHandler

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


def _persist_model_choice(file_number: str, conv_id: str, provider: str,
                          model: str) -> None:
    ChatPersistence(file_number).update_conversation(
        conv_id, provider=provider, model=model)


def _case_root(file_number: str) -> str:
    """Case folder path from the master DB (for resolving attachments)."""
    from .cases import get_case
    case = get_case(file_number)
    return case["case_path"] if case else ""


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
                "id": turn_id, "conv_id": conv_id, "status": "extracting",
                "log": [], "error": "", "done": False,
            }
        # Persist the user message immediately so it shows while we work.
        # If startup fails before the worker thread runs, free the conversation
        # here (the thread's finally-block would otherwise never run).
        try:
            append_message(file_number, conv_id, role="user", content=user_text)
            threading.Thread(
                target=self._run, daemon=True,
                args=(turn_id, file_number, conv_id, user_text, provider, model,
                      attach_rel_files, research_on),
            ).start()
        except Exception:
            with self._lock:
                self._busy_convs.discard(conv_id)
                self._turns.pop(turn_id, None)
            raise
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
