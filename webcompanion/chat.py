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
