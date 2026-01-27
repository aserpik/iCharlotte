"""
Chat persistence module for saving and loading conversations.
"""
import os
import json
from datetime import datetime
from typing import List, Optional, Dict, Any
import uuid

from ..config import GEMINI_DATA_DIR
from ..utils import log_event
from .models import Conversation, Message, QuickPrompt, BUILTIN_PROMPTS


class ChatPersistence:
    """Handles saving and loading chat conversations to JSON files."""

    VERSION = "1.0"

    def __init__(self, file_number: str):
        """
        Initialize persistence for a specific case.

        Args:
            file_number: The case file number (e.g., "5800.014")
        """
        self.file_number = file_number
        self._data = None
        self._ensure_data_dir()

    def _ensure_data_dir(self):
        """Ensure the data directory exists."""
        if not os.path.exists(GEMINI_DATA_DIR):
            os.makedirs(GEMINI_DATA_DIR)

    @property
    def file_path(self) -> str:
        """Get the path to the chat data file."""
        return os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_chat.json")

    def _get_default_data(self) -> dict:
        """Get default data structure."""
        return {
            'version': self.VERSION,
            'file_number': self.file_number,
            'conversations': [],
            'quick_prompts': [p.to_dict() for p in BUILTIN_PROMPTS],
            'settings': {
                'theme': 'light',
                'default_provider': 'Gemini',
                'default_model': 'gemini-3-flash-preview'
            }
        }

    def load(self) -> dict:
        """
        Load all chat data from file.

        Returns:
            Dictionary containing conversations, prompts, and settings.
        """
        if self._data is not None:
            return self._data

        if not os.path.exists(self.file_path):
            self._data = self._get_default_data()
            return self._data

        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                self._data = json.load(f)

            # Migrate if needed
            self._data = self._migrate_data(self._data)

            # Ensure builtin prompts are present
            self._ensure_builtin_prompts()

            log_event(f"Loaded chat data for case {self.file_number}")
            return self._data

        except Exception as e:
            log_event(f"Error loading chat data: {e}", "error")
            self._data = self._get_default_data()
            return self._data

    def _migrate_data(self, data: dict) -> dict:
        """Migrate data from older versions if needed."""
        version = data.get('version', '0.0')

        # Future migrations can be added here
        if version != self.VERSION:
            data['version'] = self.VERSION

        return data

    def _ensure_builtin_prompts(self):
        """Ensure all builtin prompts are present."""
        existing_ids = {p['id'] for p in self._data.get('quick_prompts', [])}
        for builtin in BUILTIN_PROMPTS:
            if builtin.id not in existing_ids:
                self._data['quick_prompts'].append(builtin.to_dict())

    def save(self, data: Optional[dict] = None):
        """
        Save chat data to file.

        Args:
            data: Optional data to save. If None, saves current cached data.
        """
        if data is not None:
            self._data = data

        if self._data is None:
            return

        self._ensure_data_dir()

        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(self._data, f, indent=2, ensure_ascii=False)
            log_event(f"Saved chat data for case {self.file_number}")
        except Exception as e:
            log_event(f"Error saving chat data: {e}", "error")

    # --- Conversation Management ---

    def get_conversations(self) -> List[Conversation]:
        """Get all conversations for this case."""
        data = self.load()
        return [Conversation.from_dict(c) for c in data.get('conversations', [])]

    def get_conversation(self, conv_id: str) -> Optional[Conversation]:
        """Get a specific conversation by ID."""
        data = self.load()
        for conv_data in data.get('conversations', []):
            if conv_data.get('id') == conv_id:
                return Conversation.from_dict(conv_data)
        return None

    def create_conversation(self, name: str = None, provider: str = 'Gemini',
                          model: str = 'gemini-3-flash-preview',
                          system_prompt: str = '') -> str:
        """
        Create a new conversation.

        Args:
            name: Conversation name. Defaults to "Chat HH:MM AM/PM"
            provider: LLM provider
            model: Model name
            system_prompt: System prompt for the conversation

        Returns:
            The new conversation's ID.
        """
        if name is None:
            name = f"Chat {datetime.now().strftime('%I:%M %p')}"

        conv = Conversation(
            name=name,
            provider=provider,
            model=model,
            system_prompt=system_prompt
        )

        data = self.load()
        data['conversations'].insert(0, conv.to_dict())  # Insert at beginning (most recent)
        self.save()

        return conv.id

    def update_conversation(self, conv_id: str, **kwargs):
        """
        Update conversation properties.

        Args:
            conv_id: Conversation ID
            **kwargs: Properties to update (name, provider, model, system_prompt, etc.)
        """
        data = self.load()
        for conv_data in data.get('conversations', []):
            if conv_data.get('id') == conv_id:
                for key, value in kwargs.items():
                    if key in conv_data:
                        conv_data[key] = value
                conv_data['updated_at'] = datetime.now().isoformat()
                break
        self.save()

    def delete_conversation(self, conv_id: str):
        """Delete a conversation."""
        data = self.load()
        data['conversations'] = [c for c in data.get('conversations', [])
                                  if c.get('id') != conv_id]
        self.save()

    def rename_conversation(self, conv_id: str, name: str):
        """Rename a conversation."""
        self.update_conversation(conv_id, name=name)

    # --- Message Management ---

    def add_message(self, conv_id: str, message: Message):
        """
        Add a message to a conversation.

        Args:
            conv_id: Conversation ID
            message: Message to add
        """
        data = self.load()
        for conv_data in data.get('conversations', []):
            if conv_data.get('id') == conv_id:
                conv_data['messages'].append(message.to_dict())
                conv_data['updated_at'] = datetime.now().isoformat()
                conv_data['total_tokens_used'] = conv_data.get('total_tokens_used', 0) + message.token_count
                break
        self.save()

    def update_message(self, conv_id: str, msg_id: str, content: str):
        """
        Update a message's content.

        Args:
            conv_id: Conversation ID
            msg_id: Message ID
            content: New content
        """
        data = self.load()
        for conv_data in data.get('conversations', []):
            if conv_data.get('id') == conv_id:
                for msg_data in conv_data.get('messages', []):
                    if msg_data.get('id') == msg_id:
                        if not msg_data.get('edited'):
                            msg_data['original_content'] = msg_data.get('content')
                        msg_data['content'] = content
                        msg_data['edited'] = True
                        break
                conv_data['updated_at'] = datetime.now().isoformat()
                break
        self.save()

    def delete_message(self, conv_id: str, msg_id: str):
        """Delete a message from a conversation."""
        data = self.load()
        for conv_data in data.get('conversations', []):
            if conv_data.get('id') == conv_id:
                conv_data['messages'] = [m for m in conv_data.get('messages', [])
                                          if m.get('id') != msg_id]
                conv_data['updated_at'] = datetime.now().isoformat()
                break
        self.save()

    def delete_messages_after(self, conv_id: str, msg_id: str):
        """Delete all messages after (and including) the specified message."""
        data = self.load()
        for conv_data in data.get('conversations', []):
            if conv_data.get('id') == conv_id:
                messages = conv_data.get('messages', [])
                # Find the index of the message
                idx = None
                for i, m in enumerate(messages):
                    if m.get('id') == msg_id:
                        idx = i
                        break
                if idx is not None:
                    conv_data['messages'] = messages[:idx]
                conv_data['updated_at'] = datetime.now().isoformat()
                break
        self.save()

    def toggle_pin_message(self, conv_id: str, msg_id: str) -> bool:
        """
        Toggle pin status of a message.

        Returns:
            New pin status (True if pinned, False if unpinned)
        """
        data = self.load()
        new_status = False
        for conv_data in data.get('conversations', []):
            if conv_data.get('id') == conv_id:
                for msg_data in conv_data.get('messages', []):
                    if msg_data.get('id') == msg_id:
                        msg_data['pinned'] = not msg_data.get('pinned', False)
                        new_status = msg_data['pinned']

                        # Update pinned_message_ids list
                        pinned_ids = conv_data.get('pinned_message_ids', [])
                        if new_status and msg_id not in pinned_ids:
                            pinned_ids.append(msg_id)
                        elif not new_status and msg_id in pinned_ids:
                            pinned_ids.remove(msg_id)
                        conv_data['pinned_message_ids'] = pinned_ids
                        break
                break
        self.save()
        return new_status

    # --- Search ---

    def search_conversations(self, query: str) -> List[Dict[str, Any]]:
        """
        Search across all conversations.

        Args:
            query: Search query (case-insensitive)

        Returns:
            List of matches with conversation and message info
        """
        query_lower = query.lower()
        results = []

        data = self.load()
        for conv_data in data.get('conversations', []):
            conv_name = conv_data.get('name', '')

            # Search in conversation name
            if query_lower in conv_name.lower():
                results.append({
                    'type': 'conversation',
                    'conversation_id': conv_data.get('id'),
                    'conversation_name': conv_name,
                    'match_text': conv_name,
                    'timestamp': conv_data.get('updated_at')
                })

            # Search in messages
            for msg_data in conv_data.get('messages', []):
                content = msg_data.get('content', '')
                if query_lower in content.lower():
                    # Get snippet around match
                    idx = content.lower().find(query_lower)
                    start = max(0, idx - 50)
                    end = min(len(content), idx + len(query) + 50)
                    snippet = content[start:end]
                    if start > 0:
                        snippet = '...' + snippet
                    if end < len(content):
                        snippet = snippet + '...'

                    results.append({
                        'type': 'message',
                        'conversation_id': conv_data.get('id'),
                        'conversation_name': conv_name,
                        'message_id': msg_data.get('id'),
                        'role': msg_data.get('role'),
                        'match_text': snippet,
                        'timestamp': msg_data.get('timestamp')
                    })

        # Sort by timestamp (most recent first)
        results.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
        return results

    # --- Quick Prompts ---

    def get_quick_prompts(self) -> List[QuickPrompt]:
        """Get all quick prompts (builtin + custom)."""
        data = self.load()
        return [QuickPrompt.from_dict(p) for p in data.get('quick_prompts', [])]

    def add_quick_prompt(self, name: str, prompt: str, category: str = 'Custom') -> str:
        """
        Add a custom quick prompt.

        Returns:
            The new prompt's ID.
        """
        qp = QuickPrompt(name=name, prompt=prompt, category=category, is_builtin=False)
        data = self.load()
        data['quick_prompts'].append(qp.to_dict())
        self.save()
        return qp.id

    def update_quick_prompt(self, prompt_id: str, name: str = None,
                           prompt: str = None, category: str = None):
        """Update a custom quick prompt (cannot update builtin prompts)."""
        data = self.load()
        for p_data in data.get('quick_prompts', []):
            if p_data.get('id') == prompt_id and not p_data.get('is_builtin', False):
                if name is not None:
                    p_data['name'] = name
                if prompt is not None:
                    p_data['prompt'] = prompt
                if category is not None:
                    p_data['category'] = category
                break
        self.save()

    def delete_quick_prompt(self, prompt_id: str):
        """Delete a custom quick prompt (cannot delete builtin prompts)."""
        data = self.load()
        data['quick_prompts'] = [
            p for p in data.get('quick_prompts', [])
            if p.get('id') != prompt_id or p.get('is_builtin', False)
        ]
        self.save()

    # --- Settings ---

    def get_settings(self) -> dict:
        """Get chat settings."""
        data = self.load()
        return data.get('settings', {})

    def update_settings(self, **kwargs):
        """Update chat settings."""
        data = self.load()
        settings = data.get('settings', {})
        settings.update(kwargs)
        data['settings'] = settings
        self.save()

    # --- Utility ---

    def get_conversation_count(self) -> int:
        """Get the number of conversations."""
        data = self.load()
        return len(data.get('conversations', []))

    def get_most_recent_conversation_id(self) -> Optional[str]:
        """Get the ID of the most recently updated conversation."""
        data = self.load()
        conversations = data.get('conversations', [])
        if not conversations:
            return None
        # Conversations are sorted by updated_at, first one is most recent
        return conversations[0].get('id')

    def export_conversation(self, conv_id: str, format: str = 'markdown') -> str:
        """
        Export a conversation to a string.

        Args:
            conv_id: Conversation ID
            format: 'markdown' or 'text'

        Returns:
            Formatted conversation string
        """
        conv = self.get_conversation(conv_id)
        if not conv:
            return ''

        lines = []
        lines.append(f"# {conv.name}")
        lines.append(f"Created: {conv.created_at}")
        lines.append(f"Provider: {conv.provider} / {conv.model}")
        lines.append("")

        for msg in conv.messages:
            role_label = "**You:**" if msg.role == 'user' else "**Assistant:**"
            if format == 'text':
                role_label = "You:" if msg.role == 'user' else "Assistant:"

            lines.append(role_label)
            lines.append(msg.content)
            lines.append("")

        return '\n'.join(lines)
