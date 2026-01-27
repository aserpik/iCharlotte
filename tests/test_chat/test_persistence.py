"""
Tests for chat persistence module.
"""
import os
import sys
import json
import tempfile
import shutil
import unittest
from datetime import datetime

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from icharlotte_core.chat.persistence import ChatPersistence
from icharlotte_core.chat.models import Message, Conversation, QuickPrompt, BUILTIN_PROMPTS


class TestChatPersistence(unittest.TestCase):
    """Test cases for ChatPersistence class."""

    def setUp(self):
        """Set up test fixtures."""
        # Create a temporary directory for test data
        self.temp_dir = tempfile.mkdtemp()
        self.original_data_dir = None

        # Patch GEMINI_DATA_DIR to use temp directory
        import icharlotte_core.chat.persistence as persistence_module
        self.original_data_dir = persistence_module.GEMINI_DATA_DIR
        persistence_module.GEMINI_DATA_DIR = self.temp_dir

        # Also patch in config module
        import icharlotte_core.config as config_module
        config_module.GEMINI_DATA_DIR = self.temp_dir

        self.file_number = "TEST.001"
        self.persistence = ChatPersistence(self.file_number)

    def tearDown(self):
        """Clean up test fixtures."""
        # Restore original data directory
        import icharlotte_core.chat.persistence as persistence_module
        persistence_module.GEMINI_DATA_DIR = self.original_data_dir

        import icharlotte_core.config as config_module
        config_module.GEMINI_DATA_DIR = self.original_data_dir

        # Remove temp directory
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_creates_default_data(self):
        """Test that loading non-existent file creates default data."""
        data = self.persistence.load()

        self.assertIn('version', data)
        self.assertIn('conversations', data)
        self.assertIn('quick_prompts', data)
        self.assertIn('settings', data)
        self.assertEqual(data['file_number'], self.file_number)

    def test_save_and_load(self):
        """Test saving and loading data."""
        # Modify data
        data = self.persistence.load()
        data['settings']['theme'] = 'dark'
        self.persistence.save()

        # Create new persistence instance and load
        new_persistence = ChatPersistence(self.file_number)
        loaded_data = new_persistence.load()

        self.assertEqual(loaded_data['settings']['theme'], 'dark')

    def test_create_conversation(self):
        """Test creating a new conversation."""
        conv_id = self.persistence.create_conversation(
            name="Test Conversation",
            provider="Gemini",
            model="gemini-3-flash-preview",
            system_prompt="Test prompt"
        )

        self.assertIsNotNone(conv_id)

        # Verify conversation was created
        conv = self.persistence.get_conversation(conv_id)
        self.assertIsNotNone(conv)
        self.assertEqual(conv.name, "Test Conversation")
        self.assertEqual(conv.provider, "Gemini")

    def test_delete_conversation(self):
        """Test deleting a conversation."""
        conv_id = self.persistence.create_conversation(name="To Delete")
        self.assertIsNotNone(self.persistence.get_conversation(conv_id))

        self.persistence.delete_conversation(conv_id)
        self.assertIsNone(self.persistence.get_conversation(conv_id))

    def test_rename_conversation(self):
        """Test renaming a conversation."""
        conv_id = self.persistence.create_conversation(name="Original Name")
        self.persistence.rename_conversation(conv_id, "New Name")

        conv = self.persistence.get_conversation(conv_id)
        self.assertEqual(conv.name, "New Name")

    def test_add_message(self):
        """Test adding a message to a conversation."""
        conv_id = self.persistence.create_conversation(name="Test")

        message = Message(
            role='user',
            content='Hello, world!',
            token_count=5
        )
        self.persistence.add_message(conv_id, message)

        conv = self.persistence.get_conversation(conv_id)
        self.assertEqual(len(conv.messages), 1)
        self.assertEqual(conv.messages[0].content, 'Hello, world!')

    def test_update_message(self):
        """Test updating a message."""
        conv_id = self.persistence.create_conversation(name="Test")
        message = Message(role='user', content='Original')
        self.persistence.add_message(conv_id, message)

        self.persistence.update_message(conv_id, message.id, 'Updated content')

        conv = self.persistence.get_conversation(conv_id)
        self.assertEqual(conv.messages[0].content, 'Updated content')
        self.assertTrue(conv.messages[0].edited)
        self.assertEqual(conv.messages[0].original_content, 'Original')

    def test_delete_message(self):
        """Test deleting a message."""
        conv_id = self.persistence.create_conversation(name="Test")
        message = Message(role='user', content='To delete')
        self.persistence.add_message(conv_id, message)

        self.persistence.delete_message(conv_id, message.id)

        conv = self.persistence.get_conversation(conv_id)
        self.assertEqual(len(conv.messages), 0)

    def test_delete_messages_after(self):
        """Test deleting messages after a specific message."""
        conv_id = self.persistence.create_conversation(name="Test")

        msg1 = Message(role='user', content='First')
        msg2 = Message(role='assistant', content='Second')
        msg3 = Message(role='user', content='Third')

        self.persistence.add_message(conv_id, msg1)
        self.persistence.add_message(conv_id, msg2)
        self.persistence.add_message(conv_id, msg3)

        # Delete from msg2 onwards
        self.persistence.delete_messages_after(conv_id, msg2.id)

        conv = self.persistence.get_conversation(conv_id)
        self.assertEqual(len(conv.messages), 1)
        self.assertEqual(conv.messages[0].content, 'First')

    def test_toggle_pin_message(self):
        """Test toggling pin status of a message."""
        conv_id = self.persistence.create_conversation(name="Test")
        message = Message(role='user', content='Important')
        self.persistence.add_message(conv_id, message)

        # Pin the message
        is_pinned = self.persistence.toggle_pin_message(conv_id, message.id)
        self.assertTrue(is_pinned)

        conv = self.persistence.get_conversation(conv_id)
        self.assertTrue(conv.messages[0].pinned)
        self.assertIn(message.id, conv.pinned_message_ids)

        # Unpin the message
        is_pinned = self.persistence.toggle_pin_message(conv_id, message.id)
        self.assertFalse(is_pinned)

    def test_search_conversations(self):
        """Test searching across conversations."""
        conv_id = self.persistence.create_conversation(name="Legal Discussion")
        msg = Message(role='user', content='Discuss liability issues')
        self.persistence.add_message(conv_id, msg)

        # Search in conversation name
        results = self.persistence.search_conversations("Legal")
        self.assertTrue(len(results) > 0)

        # Search in message content
        results = self.persistence.search_conversations("liability")
        self.assertTrue(len(results) > 0)

        # Search for non-existent term
        results = self.persistence.search_conversations("nonexistent123456")
        self.assertEqual(len(results), 0)

    def test_quick_prompts_builtin(self):
        """Test that builtin prompts are loaded."""
        prompts = self.persistence.get_quick_prompts()
        builtin_prompts = [p for p in prompts if p.is_builtin]

        self.assertTrue(len(builtin_prompts) > 0)

    def test_add_custom_quick_prompt(self):
        """Test adding a custom quick prompt."""
        prompt_id = self.persistence.add_quick_prompt(
            name="My Custom Prompt",
            prompt="Do something custom",
            category="Custom"
        )

        prompts = self.persistence.get_quick_prompts()
        custom = [p for p in prompts if p.id == prompt_id]

        self.assertEqual(len(custom), 1)
        self.assertEqual(custom[0].name, "My Custom Prompt")
        self.assertFalse(custom[0].is_builtin)

    def test_delete_custom_quick_prompt(self):
        """Test deleting a custom prompt."""
        prompt_id = self.persistence.add_quick_prompt(name="To Delete", prompt="Text")
        self.persistence.delete_quick_prompt(prompt_id)

        prompts = self.persistence.get_quick_prompts()
        deleted = [p for p in prompts if p.id == prompt_id]
        self.assertEqual(len(deleted), 0)

    def test_cannot_delete_builtin_prompt(self):
        """Test that builtin prompts cannot be deleted."""
        builtin_id = BUILTIN_PROMPTS[0].id
        self.persistence.delete_quick_prompt(builtin_id)

        prompts = self.persistence.get_quick_prompts()
        builtin = [p for p in prompts if p.id == builtin_id]
        self.assertEqual(len(builtin), 1)  # Still exists

    def test_get_most_recent_conversation(self):
        """Test getting the most recent conversation."""
        conv1_id = self.persistence.create_conversation(name="First")
        conv2_id = self.persistence.create_conversation(name="Second")

        # Most recent should be conv2 (added last)
        recent_id = self.persistence.get_most_recent_conversation_id()
        self.assertEqual(recent_id, conv2_id)

    def test_export_conversation(self):
        """Test exporting a conversation to markdown."""
        conv_id = self.persistence.create_conversation(name="Export Test")
        self.persistence.add_message(conv_id, Message(role='user', content='Hello'))
        self.persistence.add_message(conv_id, Message(role='assistant', content='Hi there!'))

        exported = self.persistence.export_conversation(conv_id, format='markdown')

        self.assertIn("Export Test", exported)
        self.assertIn("**You:**", exported)
        self.assertIn("Hello", exported)
        self.assertIn("**Assistant:**", exported)
        self.assertIn("Hi there!", exported)

    def test_multiple_conversations_isolation(self):
        """Test that multiple conversations are properly isolated."""
        conv1_id = self.persistence.create_conversation(name="Conv 1")
        conv2_id = self.persistence.create_conversation(name="Conv 2")

        self.persistence.add_message(conv1_id, Message(role='user', content='Message in conv 1'))
        self.persistence.add_message(conv2_id, Message(role='user', content='Message in conv 2'))

        conv1 = self.persistence.get_conversation(conv1_id)
        conv2 = self.persistence.get_conversation(conv2_id)

        self.assertEqual(len(conv1.messages), 1)
        self.assertEqual(len(conv2.messages), 1)
        self.assertEqual(conv1.messages[0].content, 'Message in conv 1')
        self.assertEqual(conv2.messages[0].content, 'Message in conv 2')

    def test_conversation_count(self):
        """Test getting conversation count."""
        self.assertEqual(self.persistence.get_conversation_count(), 0)

        self.persistence.create_conversation(name="Conv 1")
        self.assertEqual(self.persistence.get_conversation_count(), 1)

        self.persistence.create_conversation(name="Conv 2")
        self.assertEqual(self.persistence.get_conversation_count(), 2)

    def test_settings_update(self):
        """Test updating settings."""
        self.persistence.update_settings(theme='dark', custom_setting='value')

        settings = self.persistence.get_settings()
        self.assertEqual(settings['theme'], 'dark')
        self.assertEqual(settings['custom_setting'], 'value')


class TestMessageModel(unittest.TestCase):
    """Test cases for Message model."""

    def test_message_creation(self):
        """Test creating a message."""
        msg = Message(role='user', content='Test content')

        self.assertIsNotNone(msg.id)
        self.assertEqual(msg.role, 'user')
        self.assertEqual(msg.content, 'Test content')
        self.assertFalse(msg.pinned)
        self.assertFalse(msg.edited)

    def test_message_to_dict(self):
        """Test converting message to dictionary."""
        msg = Message(role='user', content='Test', pinned=True)
        data = msg.to_dict()

        self.assertEqual(data['role'], 'user')
        self.assertEqual(data['content'], 'Test')
        self.assertTrue(data['pinned'])

    def test_message_from_dict(self):
        """Test creating message from dictionary."""
        data = {
            'id': 'test-id',
            'role': 'assistant',
            'content': 'Response',
            'timestamp': '2025-01-01T12:00:00',
            'pinned': True,
            'edited': True,
            'original_content': 'Original'
        }
        msg = Message.from_dict(data)

        self.assertEqual(msg.id, 'test-id')
        self.assertEqual(msg.role, 'assistant')
        self.assertTrue(msg.pinned)
        self.assertTrue(msg.edited)


class TestConversationModel(unittest.TestCase):
    """Test cases for Conversation model."""

    def test_conversation_creation(self):
        """Test creating a conversation."""
        conv = Conversation(name='Test Conv', provider='Gemini')

        self.assertIsNotNone(conv.id)
        self.assertEqual(conv.name, 'Test Conv')
        self.assertEqual(conv.provider, 'Gemini')
        self.assertEqual(len(conv.messages), 0)

    def test_add_message_updates_timestamp(self):
        """Test that adding message updates the conversation."""
        conv = Conversation(name='Test')
        original_updated = conv.updated_at

        msg = Message(role='user', content='Hello', token_count=2)
        conv.add_message(msg)

        self.assertEqual(len(conv.messages), 1)
        self.assertNotEqual(conv.updated_at, original_updated)
        self.assertEqual(conv.total_tokens_used, 2)

    def test_get_history_for_llm(self):
        """Test getting conversation history in LLM format."""
        conv = Conversation(name='Test')
        conv.add_message(Message(role='user', content='Hello'))
        conv.add_message(Message(role='assistant', content='Hi'))

        history = conv.get_history_for_llm()

        self.assertEqual(len(history), 2)
        self.assertEqual(history[0]['role'], 'user')
        self.assertEqual(history[0]['content'], 'Hello')
        self.assertEqual(history[1]['role'], 'assistant')


if __name__ == '__main__':
    unittest.main()
