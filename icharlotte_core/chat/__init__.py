# Chat module for iCharlotte
# Provides conversation persistence, token counting, and context management

from .persistence import ChatPersistence
from .token_counter import TokenCounter
from .models import Conversation, Message, QuickPrompt, BUILTIN_PROMPTS, TRANSCRIBE_PROMPT
from .legal_research import (
    ChatResearchError,
    ChatResearchOutputMode,
    ChatResearchPacket,
    ChatResearchSettings,
    CourtListenerMode,
)

__all__ = [
    'ChatPersistence',
    'TokenCounter',
    'Conversation',
    'Message',
    'QuickPrompt',
    'BUILTIN_PROMPTS',
    'TRANSCRIBE_PROMPT',
    'ChatResearchError',
    'ChatResearchOutputMode',
    'ChatResearchPacket',
    'ChatResearchSettings',
    'CourtListenerMode',
]
