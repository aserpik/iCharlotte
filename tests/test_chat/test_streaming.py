"""
Tests for streaming functionality in the LLM module.
"""
import os
import sys
import unittest
from unittest.mock import Mock, patch, MagicMock

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


class TestLLMStreaming(unittest.TestCase):
    """Test cases for LLM streaming functionality."""

    def setUp(self):
        """Set up test fixtures."""
        # Mock API keys
        self.api_keys_patch = patch.dict(
            'icharlotte_core.llm.API_KEYS',
            {'Gemini': 'test-key', 'OpenAI': 'test-key', 'Claude': 'test-key'}
        )
        self.api_keys_patch.start()

    def tearDown(self):
        """Clean up test fixtures."""
        self.api_keys_patch.stop()

    @patch('icharlotte_core.llm.requests.post')
    def test_openai_streaming_generator(self, mock_post):
        """Test OpenAI streaming returns a generator."""
        from icharlotte_core.llm import LLMHandler

        # Mock streaming response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.iter_lines.return_value = [
            b'data: {"choices":[{"delta":{"content":"Hello"}}]}',
            b'data: {"choices":[{"delta":{"content":" World"}}]}',
            b'data: [DONE]'
        ]
        mock_post.return_value = mock_response

        result = LLMHandler.generate(
            provider='OpenAI',
            model='gpt-4o',
            system_prompt='You are helpful.',
            user_prompt='Say hello',
            file_contents='',
            settings={'stream': True}
        )

        # Should return a generator
        self.assertTrue(hasattr(result, '__iter__'))

        # Collect all tokens
        tokens = list(result)
        self.assertEqual(tokens, ['Hello', ' World'])

    @patch('icharlotte_core.llm.requests.post')
    def test_claude_streaming_generator(self, mock_post):
        """Test Claude streaming returns a generator."""
        from icharlotte_core.llm import LLMHandler

        # Mock streaming response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.iter_lines.return_value = [
            b'data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Hello"}}',
            b'data: {"type":"content_block_delta","delta":{"type":"text_delta","text":" World"}}',
            b'data: {"type":"message_stop"}'
        ]
        mock_post.return_value = mock_response

        result = LLMHandler.generate(
            provider='Claude',
            model='claude-3.5-sonnet',
            system_prompt='You are helpful.',
            user_prompt='Say hello',
            file_contents='',
            settings={'stream': True}
        )

        # Should return a generator
        self.assertTrue(hasattr(result, '__iter__'))

        # Collect all tokens
        tokens = list(result)
        self.assertEqual(tokens, ['Hello', ' World'])

    @patch('icharlotte_core.llm.requests.post')
    def test_openai_non_streaming(self, mock_post):
        """Test OpenAI non-streaming returns complete text."""
        from icharlotte_core.llm import LLMHandler

        # Mock non-streaming response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'choices': [{'message': {'content': 'Hello World'}}]
        }
        mock_post.return_value = mock_response

        result = LLMHandler.generate(
            provider='OpenAI',
            model='gpt-4o',
            system_prompt='You are helpful.',
            user_prompt='Say hello',
            file_contents='',
            settings={'stream': False}
        )

        self.assertEqual(result, 'Hello World')

    @patch('icharlotte_core.llm.requests.post')
    def test_claude_non_streaming(self, mock_post):
        """Test Claude non-streaming returns complete text."""
        from icharlotte_core.llm import LLMHandler

        # Mock non-streaming response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'content': [{'text': 'Hello World'}]
        }
        mock_post.return_value = mock_response

        result = LLMHandler.generate(
            provider='Claude',
            model='claude-3.5-sonnet',
            system_prompt='You are helpful.',
            user_prompt='Say hello',
            file_contents='',
            settings={'stream': False}
        )

        self.assertEqual(result, 'Hello World')

    @patch('icharlotte_core.llm.requests.post')
    def test_openai_streaming_error_handling(self, mock_post):
        """Test OpenAI streaming handles errors gracefully."""
        from icharlotte_core.llm import LLMHandler

        # Mock error response
        mock_response = Mock()
        mock_response.status_code = 500
        mock_response.text = 'Internal Server Error'
        mock_post.return_value = mock_response

        with self.assertRaises(Exception) as context:
            result = LLMHandler.generate(
                provider='OpenAI',
                model='gpt-4o',
                system_prompt='You are helpful.',
                user_prompt='Say hello',
                file_contents='',
                settings={'stream': True}
            )

        self.assertIn('OpenAI Error', str(context.exception))

    @patch('icharlotte_core.llm.requests.post')
    def test_claude_streaming_error_handling(self, mock_post):
        """Test Claude streaming handles errors gracefully."""
        from icharlotte_core.llm import LLMHandler

        # Mock error response
        mock_response = Mock()
        mock_response.status_code = 401
        mock_response.text = 'Unauthorized'
        mock_post.return_value = mock_response

        with self.assertRaises(Exception) as context:
            result = LLMHandler.generate(
                provider='Claude',
                model='claude-3.5-sonnet',
                system_prompt='You are helpful.',
                user_prompt='Say hello',
                file_contents='',
                settings={'stream': True}
            )

        self.assertIn('Claude Error', str(context.exception))

    def test_streaming_with_history(self):
        """Test that history is passed correctly in streaming mode."""
        from icharlotte_core.llm import LLMHandler

        with patch('icharlotte_core.llm.requests.post') as mock_post:
            mock_response = Mock()
            mock_response.status_code = 200
            mock_response.iter_lines.return_value = [
                b'data: {"choices":[{"delta":{"content":"Response"}}]}',
                b'data: [DONE]'
            ]
            mock_post.return_value = mock_response

            history = [
                {'role': 'user', 'content': 'Previous question'},
                {'role': 'assistant', 'content': 'Previous answer'}
            ]

            result = LLMHandler.generate(
                provider='OpenAI',
                model='gpt-4o',
                system_prompt='You are helpful.',
                user_prompt='Follow-up question',
                file_contents='',
                settings={'stream': True},
                history=history
            )

            # Verify that the request was made with history
            call_args = mock_post.call_args
            request_data = call_args[1]['json']

            # Should have system + history + new message
            self.assertGreater(len(request_data['messages']), 2)


class TestLLMWorkerStreaming(unittest.TestCase):
    """Test cases for LLMWorker streaming functionality."""

    def test_worker_emits_tokens(self):
        """Test that LLMWorker emits tokens during streaming."""
        try:
            from PySide6.QtCore import QCoreApplication
            from icharlotte_core.llm import LLMWorker

            # Create QApplication if needed for Qt signals
            app = QCoreApplication.instance()
            if not app:
                app = QCoreApplication([])

            tokens_received = []

            def on_token(token):
                tokens_received.append(token)

            with patch('icharlotte_core.llm.LLMHandler.generate') as mock_generate:
                # Mock a generator
                mock_generate.return_value = iter(['Hello', ' ', 'World'])

                worker = LLMWorker(
                    provider='OpenAI',
                    model='gpt-4o',
                    system='System prompt',
                    user='User prompt',
                    files='',
                    settings={'stream': True}
                )

                worker.new_token.connect(on_token)
                worker.run()

                self.assertEqual(tokens_received, ['Hello', ' ', 'World'])

        except ImportError:
            self.skipTest("PySide6 not available")

    def test_worker_stop_request(self):
        """Test that worker respects stop requests."""
        try:
            from PySide6.QtCore import QCoreApplication
            from icharlotte_core.llm import LLMWorker

            app = QCoreApplication.instance()
            if not app:
                app = QCoreApplication([])

            tokens_received = []
            stop_after = 2

            def on_token(token):
                tokens_received.append(token)
                if len(tokens_received) >= stop_after:
                    worker.request_stop()

            def slow_generator():
                for word in ['Hello', ' ', 'World', '!']:
                    yield word

            with patch('icharlotte_core.llm.LLMHandler.generate') as mock_generate:
                mock_generate.return_value = slow_generator()

                worker = LLMWorker(
                    provider='OpenAI',
                    model='gpt-4o',
                    system='System prompt',
                    user='User prompt',
                    files='',
                    settings={'stream': True}
                )

                worker.new_token.connect(on_token)
                worker.run()

                # Should have stopped early (at most 3 tokens due to timing)
                self.assertLessEqual(len(tokens_received), 3)

        except ImportError:
            self.skipTest("PySide6 not available")


class TestStreamingEdgeCases(unittest.TestCase):
    """Edge case tests for streaming."""

    def setUp(self):
        """Set up test fixtures."""
        self.api_keys_patch = patch.dict(
            'icharlotte_core.llm.API_KEYS',
            {'Gemini': 'test-key', 'OpenAI': 'test-key', 'Claude': 'test-key'}
        )
        self.api_keys_patch.start()

    def tearDown(self):
        """Clean up test fixtures."""
        self.api_keys_patch.stop()

    @patch('icharlotte_core.llm.requests.post')
    def test_empty_stream(self, mock_post):
        """Test handling of empty stream."""
        from icharlotte_core.llm import LLMHandler

        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.iter_lines.return_value = [
            b'data: [DONE]'
        ]
        mock_post.return_value = mock_response

        result = LLMHandler.generate(
            provider='OpenAI',
            model='gpt-4o',
            system_prompt='You are helpful.',
            user_prompt='Say hello',
            file_contents='',
            settings={'stream': True}
        )

        tokens = list(result)
        self.assertEqual(tokens, [])

    @patch('icharlotte_core.llm.requests.post')
    def test_malformed_json_in_stream(self, mock_post):
        """Test handling of malformed JSON in stream."""
        from icharlotte_core.llm import LLMHandler

        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.iter_lines.return_value = [
            b'data: {"choices":[{"delta":{"content":"Hello"}}]}',
            b'data: not valid json',
            b'data: {"choices":[{"delta":{"content":" World"}}]}',
            b'data: [DONE]'
        ]
        mock_post.return_value = mock_response

        result = LLMHandler.generate(
            provider='OpenAI',
            model='gpt-4o',
            system_prompt='You are helpful.',
            user_prompt='Say hello',
            file_contents='',
            settings={'stream': True}
        )

        # Should skip malformed JSON and continue
        tokens = list(result)
        self.assertEqual(tokens, ['Hello', ' World'])

    @patch('icharlotte_core.llm.requests.post')
    def test_stream_with_empty_deltas(self, mock_post):
        """Test handling of empty deltas in stream."""
        from icharlotte_core.llm import LLMHandler

        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.iter_lines.return_value = [
            b'data: {"choices":[{"delta":{"content":"Hello"}}]}',
            b'data: {"choices":[{"delta":{}}]}',  # Empty delta
            b'data: {"choices":[{"delta":{"content":""}}]}',  # Empty content
            b'data: {"choices":[{"delta":{"content":" World"}}]}',
            b'data: [DONE]'
        ]
        mock_post.return_value = mock_response

        result = LLMHandler.generate(
            provider='OpenAI',
            model='gpt-4o',
            system_prompt='You are helpful.',
            user_prompt='Say hello',
            file_contents='',
            settings={'stream': True}
        )

        tokens = list(result)
        # Should only have non-empty tokens
        self.assertEqual(tokens, ['Hello', ' World'])


if __name__ == '__main__':
    unittest.main()
