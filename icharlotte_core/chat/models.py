"""
Data models for chat conversations and messages.
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from datetime import datetime
import uuid


@dataclass
class Attachment:
    """Represents a file attachment in a message."""
    name: str
    path: str
    type: str  # 'file', 'image'
    token_estimate: int = 0

    def to_dict(self) -> dict:
        return {
            'name': self.name,
            'path': self.path,
            'type': self.type,
            'token_estimate': self.token_estimate
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'Attachment':
        return cls(
            name=data.get('name', ''),
            path=data.get('path', ''),
            type=data.get('type', 'file'),
            token_estimate=data.get('token_estimate', 0)
        )


@dataclass
class Message:
    """Represents a single message in a conversation."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    role: str = 'user'  # 'user' or 'assistant'
    content: str = ''
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
    attachments: List[Attachment] = field(default_factory=list)
    pinned: bool = False
    edited: bool = False
    original_content: Optional[str] = None
    token_count: int = 0
    model_used: Optional[str] = None
    response_time_ms: Optional[int] = None

    def __post_init__(self):
        """Normalize attachments to Attachment objects."""
        self.attachments = [
            Attachment.from_dict(a) if isinstance(a, dict) else a
            for a in self.attachments
        ]

    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'role': self.role,
            'content': self.content,
            'timestamp': self.timestamp,
            'attachments': [a.to_dict() if hasattr(a, 'to_dict') else a for a in self.attachments],
            'pinned': self.pinned,
            'edited': self.edited,
            'original_content': self.original_content,
            'token_count': self.token_count,
            'model_used': self.model_used,
            'response_time_ms': self.response_time_ms
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'Message':
        attachments = [Attachment.from_dict(a) for a in data.get('attachments', [])]
        return cls(
            id=data.get('id', str(uuid.uuid4())),
            role=data.get('role', 'user'),
            content=data.get('content', ''),
            timestamp=data.get('timestamp', datetime.now().isoformat()),
            attachments=attachments,
            pinned=data.get('pinned', False),
            edited=data.get('edited', False),
            original_content=data.get('original_content'),
            token_count=data.get('token_count', 0),
            model_used=data.get('model_used'),
            response_time_ms=data.get('response_time_ms')
        )


@dataclass
class Conversation:
    """Represents a chat conversation with messages and metadata."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = 'New Chat'
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())
    provider: str = 'Gemini'
    model: str = 'gemini-3-flash-preview'
    system_prompt: str = ''
    settings: Dict[str, Any] = field(default_factory=dict)
    messages: List[Message] = field(default_factory=list)
    pinned_message_ids: List[str] = field(default_factory=list)
    tags: List[str] = field(default_factory=list)
    context_summary: Optional[str] = None
    total_tokens_used: int = 0

    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'name': self.name,
            'created_at': self.created_at,
            'updated_at': self.updated_at,
            'provider': self.provider,
            'model': self.model,
            'system_prompt': self.system_prompt,
            'settings': self.settings,
            'messages': [m.to_dict() for m in self.messages],
            'pinned_message_ids': self.pinned_message_ids,
            'tags': self.tags,
            'context_summary': self.context_summary,
            'total_tokens_used': self.total_tokens_used
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'Conversation':
        messages = [Message.from_dict(m) for m in data.get('messages', [])]
        return cls(
            id=data.get('id', str(uuid.uuid4())),
            name=data.get('name', 'New Chat'),
            created_at=data.get('created_at', datetime.now().isoformat()),
            updated_at=data.get('updated_at', datetime.now().isoformat()),
            provider=data.get('provider', 'Gemini'),
            model=data.get('model', 'gemini-3-flash-preview'),
            system_prompt=data.get('system_prompt', ''),
            settings=data.get('settings', {}),
            messages=messages,
            pinned_message_ids=data.get('pinned_message_ids', []),
            tags=data.get('tags', []),
            context_summary=data.get('context_summary'),
            total_tokens_used=data.get('total_tokens_used', 0)
        )

    def add_message(self, message: Message):
        """Add a message and update timestamps."""
        self.messages.append(message)
        self.updated_at = datetime.now().isoformat()
        self.total_tokens_used += message.token_count

    def get_history_for_llm(self) -> List[Dict[str, str]]:
        """Get conversation history in LLM-compatible format."""
        return [{'role': m.role, 'content': m.content} for m in self.messages]


@dataclass
class QuickPrompt:
    """Represents a quick prompt template."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = ''
    prompt: str = ''
    category: str = 'General'
    is_builtin: bool = False

    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'name': self.name,
            'prompt': self.prompt,
            'category': self.category,
            'is_builtin': self.is_builtin
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'QuickPrompt':
        return cls(
            id=data.get('id', str(uuid.uuid4())),
            name=data.get('name', ''),
            prompt=data.get('prompt', ''),
            category=data.get('category', 'General'),
            is_builtin=data.get('is_builtin', False)
        )


# Built-in legal prompts
BUILTIN_PROMPTS = [
    QuickPrompt(
        id='builtin_summary',
        name='Summarize Document',
        prompt='Provide a comprehensive summary of the attached document, highlighting key facts, dates, and parties involved.',
        category='Summary',
        is_builtin=True
    ),
    QuickPrompt(
        id='builtin_liability',
        name='Liability Analysis',
        prompt='Analyze the attached documents for potential liability issues. Identify strengths and weaknesses of each party\'s position.',
        category='Analysis',
        is_builtin=True
    ),
    QuickPrompt(
        id='builtin_timeline',
        name='Timeline Extraction',
        prompt='Extract all dates and events from the attached documents and present them in chronological order.',
        category='Extraction',
        is_builtin=True
    ),
    QuickPrompt(
        id='builtin_deposition',
        name='Deposition Summary',
        prompt='Summarize the key testimony from this deposition, including admissions, contradictions, and notable statements.',
        category='Summary',
        is_builtin=True
    ),
    QuickPrompt(
        id='builtin_interrogatory',
        name='Interrogatory Responses',
        prompt='Draft responses to the interrogatories in the attached document based on the case facts provided.',
        category='Drafting',
        is_builtin=True
    ),
    QuickPrompt(
        id='builtin_discovery',
        name='Discovery Analysis',
        prompt='Review the attached discovery requests and identify any objectionable requests, overly broad requests, or requests that may be problematic to respond to.',
        category='Analysis',
        is_builtin=True
    ),
    QuickPrompt(
        id='builtin_damages',
        name='Damages Assessment',
        prompt='Based on the attached documents, provide an assessment of potential damages including economic and non-economic damages.',
        category='Analysis',
        is_builtin=True
    ),
    QuickPrompt(
        id='builtin_key_facts',
        name='Key Facts Extraction',
        prompt='Extract and list all key facts from the attached documents that are relevant to establishing liability or damages.',
        category='Extraction',
        is_builtin=True
    ),
]
