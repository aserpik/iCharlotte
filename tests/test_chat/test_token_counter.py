"""
Tests for token counting module.
"""
import os
import sys
import unittest

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from icharlotte_core.chat.token_counter import TokenCounter


class TestTokenCounter(unittest.TestCase):
    """Test cases for TokenCounter class."""

    def test_estimate_tokens_empty(self):
        """Test token estimation for empty string."""
        count = TokenCounter.estimate_tokens("")
        self.assertEqual(count, 0)

        count = TokenCounter.estimate_tokens(None)
        self.assertEqual(count, 0)

    def test_estimate_tokens_simple(self):
        """Test token estimation for simple text."""
        # "Hello world" ~= 11 chars / 4 ~= 2-3 tokens
        count = TokenCounter.estimate_tokens("Hello world")
        self.assertGreater(count, 0)
        self.assertLess(count, 20)  # Should be roughly 2-4 tokens

    def test_estimate_tokens_longer_text(self):
        """Test token estimation for longer text."""
        text = "This is a longer piece of text that should have more tokens. " * 10
        count = TokenCounter.estimate_tokens(text)

        # Should be roughly len(text)/4
        expected_min = len(text) // 6
        expected_max = len(text) // 2
        self.assertGreater(count, expected_min)
        self.assertLess(count, expected_max)

    def test_estimate_tokens_code(self):
        """Test token estimation for code (should be higher ratio)."""
        code = """
        def hello_world():
            print("Hello, World!")
            return {"status": "ok"}
        """
        count = TokenCounter.estimate_tokens(code)
        self.assertGreater(count, 0)

    def test_estimate_tokens_detailed(self):
        """Test detailed token estimation."""
        text = "Hello world! This is a test."
        details = TokenCounter.estimate_tokens_detailed(text)

        self.assertIn('total', details)
        self.assertIn('text', details)
        self.assertIn('whitespace', details)
        self.assertGreater(details['total'], 0)

    def test_get_context_limit_known_model(self):
        """Test getting context limit for known models."""
        # Gemini models
        limit = TokenCounter.get_context_limit('gemini-3-flash-preview')
        self.assertEqual(limit, 1000000)

        limit = TokenCounter.get_context_limit('gemini-1.5-pro')
        self.assertEqual(limit, 2000000)

        # OpenAI models
        limit = TokenCounter.get_context_limit('gpt-4o')
        self.assertEqual(limit, 128000)

        # Claude models
        limit = TokenCounter.get_context_limit('claude-3.5-sonnet')
        self.assertEqual(limit, 200000)

    def test_get_context_limit_partial_match(self):
        """Test context limit with partial model name match."""
        # Should match 'gpt-4' pattern
        limit = TokenCounter.get_context_limit('gpt-4-turbo-2024-04-09')
        self.assertGreater(limit, 0)

    def test_get_context_limit_unknown_with_provider(self):
        """Test context limit fallback for unknown model with provider."""
        limit = TokenCounter.get_context_limit('unknown-model', provider='Gemini')
        self.assertEqual(limit, 1000000)

        limit = TokenCounter.get_context_limit('unknown-model', provider='OpenAI')
        self.assertEqual(limit, 128000)

        limit = TokenCounter.get_context_limit('unknown-model', provider='Claude')
        self.assertEqual(limit, 200000)

    def test_get_context_limit_completely_unknown(self):
        """Test context limit fallback for completely unknown model."""
        limit = TokenCounter.get_context_limit('completely-unknown-model-xyz')
        self.assertEqual(limit, 128000)  # Ultimate fallback

    def test_calculate_context_usage(self):
        """Test calculating total context usage."""
        messages = [
            {'content': 'Hello'},
            {'content': 'World'}
        ]
        system_prompt = "You are a helpful assistant."
        file_content = "Document content here."

        usage = TokenCounter.calculate_context_usage(
            messages, system_prompt, file_content, 'Gemini'
        )

        self.assertIn('system_tokens', usage)
        self.assertIn('file_tokens', usage)
        self.assertIn('message_tokens', usage)
        self.assertIn('total_tokens', usage)

        self.assertGreater(usage['total_tokens'], 0)
        self.assertEqual(
            usage['total_tokens'],
            usage['system_tokens'] + usage['file_tokens'] + usage['message_tokens']
        )

    def test_get_usage_percentage(self):
        """Test calculating usage percentage."""
        # 50% usage
        percentage = TokenCounter.get_usage_percentage(50000, 'gpt-4o')
        # gpt-4o has 128k limit, so 50k should be ~39%
        self.assertGreater(percentage, 30)
        self.assertLess(percentage, 50)

        # 0 tokens
        percentage = TokenCounter.get_usage_percentage(0, 'gpt-4o')
        self.assertEqual(percentage, 0.0)

    def test_format_token_count(self):
        """Test formatting token counts."""
        self.assertEqual(TokenCounter.format_token_count(500), "500")
        self.assertEqual(TokenCounter.format_token_count(1500), "1.5k")
        self.assertEqual(TokenCounter.format_token_count(1500000), "1.5M")
        self.assertEqual(TokenCounter.format_token_count(100000), "100.0k")

    def test_should_summarize(self):
        """Test summarization recommendation."""
        # Below threshold (80%)
        should = TokenCounter.should_summarize(50000, 'gpt-4o')
        self.assertFalse(should)  # 50k/128k = 39%

        # Above threshold
        should = TokenCounter.should_summarize(110000, 'gpt-4o')
        self.assertTrue(should)  # 110k/128k = 86%

        # Custom threshold
        should = TokenCounter.should_summarize(50000, 'gpt-4o', threshold_percent=30.0)
        self.assertTrue(should)  # 39% > 30%

    def test_estimate_remaining_tokens(self):
        """Test estimating remaining tokens."""
        remaining = TokenCounter.estimate_remaining_tokens(50000, 'gpt-4o')
        # 128000 - 50000 - 4096 (reserve) = 73904
        self.assertEqual(remaining, 73904)

        # Custom reserve
        remaining = TokenCounter.estimate_remaining_tokens(50000, 'gpt-4o', reserve_for_response=8000)
        self.assertEqual(remaining, 70000)

        # When near limit
        remaining = TokenCounter.estimate_remaining_tokens(126000, 'gpt-4o')
        self.assertEqual(remaining, 0)  # Should be 0, not negative

    def test_consistency_across_providers(self):
        """Test that token estimation is consistent regardless of provider."""
        text = "This is a test sentence for token counting."

        count_gemini = TokenCounter.estimate_tokens(text, 'Gemini')
        count_openai = TokenCounter.estimate_tokens(text, 'OpenAI')
        count_claude = TokenCounter.estimate_tokens(text, 'Claude')

        # All should give the same result since we use a universal estimation
        self.assertEqual(count_gemini, count_openai)
        self.assertEqual(count_openai, count_claude)


class TestTokenCounterEdgeCases(unittest.TestCase):
    """Edge case tests for TokenCounter."""

    def test_very_long_text(self):
        """Test with very long text."""
        text = "word " * 100000  # 500k+ characters
        count = TokenCounter.estimate_tokens(text)
        self.assertGreater(count, 100000)

    def test_unicode_text(self):
        """Test with unicode characters."""
        text = "This is a multilingual"
        count = TokenCounter.estimate_tokens(text)
        self.assertGreater(count, 0)

    def test_special_characters(self):
        """Test with special characters."""
        text = "!@#$%^&*(){}[]|\\:\";<>?,./~`"
        count = TokenCounter.estimate_tokens(text)
        self.assertGreater(count, 0)

    def test_newlines_and_whitespace(self):
        """Test with various whitespace."""
        text = "Line 1\nLine 2\n\nLine 3\tTabbed"
        count = TokenCounter.estimate_tokens(text)
        self.assertGreater(count, 0)

    def test_empty_messages_list(self):
        """Test context calculation with empty messages."""
        usage = TokenCounter.calculate_context_usage([], "", "", "Gemini")
        self.assertEqual(usage['total_tokens'], 0)

    def test_messages_as_objects(self):
        """Test that messages work whether dict or string content."""
        messages = [
            {'content': 'Hello'},
            'Direct string message'
        ]
        usage = TokenCounter.calculate_context_usage(messages, "", "", "Gemini")
        self.assertGreater(usage['message_tokens'], 0)


if __name__ == '__main__':
    unittest.main()
