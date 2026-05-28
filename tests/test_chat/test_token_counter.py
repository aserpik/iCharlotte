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
        limit = TokenCounter.get_context_limit('gemini-3.5-flash')
        self.assertEqual(limit, 1048576)

        limit = TokenCounter.get_context_limit('gemini-3.1-pro-preview')
        self.assertEqual(limit, 1048576)

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
        self.assertEqual(limit, 1048576)

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
        # 128000 - 50000 - 16384 (reserve) = 61616
        self.assertEqual(remaining, 61616)

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


class TestModelFamilyContextLimits(unittest.TestCase):
    """Verify family-pattern resolution gives the right window for every
    currently-shipping model and adapts to new versions in known families."""

    # ---- OpenAI ----

    def test_gpt_5_family_resolves_to_400k(self):
        # Cases the user actually sees in the dialog (live OpenAI catalog
        # has gpt-5.1/5.2 variants today; users will type/select gpt-5.5-pro
        # once it ships). All should land on 400k, not the 128k OpenAI default.
        for model in (
            'gpt-5',
            'gpt-5-pro',
            'gpt-5.1-instant',
            'gpt-5.1-thinking',
            'gpt-5.2-instant',
            'gpt-5.2-thinking',
            'gpt-5.5-pro',
            'gpt-5-pro-2026-04-01',
        ):
            self.assertEqual(
                TokenCounter.get_context_limit(model),
                400_000,
                f"{model} should resolve to GPT-5 family (400k)",
            )

    def test_gpt_4_1_family_resolves_to_1m(self):
        for model in ('gpt-4.1', 'gpt-4.1-mini', 'gpt-4.1-nano', 'gpt-4.1-2025-04-14'):
            self.assertEqual(
                TokenCounter.get_context_limit(model),
                1_047_576,
                f"{model} should resolve to GPT-4.1 family (1M)",
            )

    def test_gpt_4o_family_resolves_to_128k(self):
        for model in ('gpt-4o', 'gpt-4o-mini', 'gpt-4o-2024-08-06', 'gpt-4-turbo', 'gpt-4-turbo-2024-04-09'):
            self.assertEqual(TokenCounter.get_context_limit(model), 128_000)

    def test_o_series_reasoning_resolves_to_200k(self):
        # o1 / o3 / o4 / future o-series → 200k, except mini variants (overridden)
        for model in ('o1', 'o1-preview', 'o3', 'o4-preview'):
            self.assertEqual(
                TokenCounter.get_context_limit(model),
                200_000,
                f"{model} should be 200k",
            )

    def test_o_series_mini_variants_overridden(self):
        self.assertEqual(TokenCounter.get_context_limit('o1-mini'), 128_000)
        self.assertEqual(TokenCounter.get_context_limit('o3-mini'), 200_000)

    def test_chatgpt_snapshots_resolve_to_128k(self):
        self.assertEqual(TokenCounter.get_context_limit('chatgpt-4o-latest'), 128_000)

    # ---- Claude ----

    def test_claude_sonnet_4_5_plus_resolves_to_1m(self):
        for model in ('claude-sonnet-4-5', 'claude-sonnet-4-6', 'claude-sonnet-4-7',
                      'claude-sonnet-4-5-20250929'):
            self.assertEqual(
                TokenCounter.get_context_limit(model),
                1_000_000,
                f"{model} should resolve to 1M (Sonnet 4.5+)",
            )

    def test_claude_opus_4_6_plus_resolves_to_1m(self):
        for model in ('claude-opus-4-6', 'claude-opus-4-7', 'claude-opus-4-7-20260201'):
            self.assertEqual(
                TokenCounter.get_context_limit(model),
                1_000_000,
                f"{model} should resolve to 1M (Opus 4.6+)",
            )

    def test_claude_4_base_resolves_to_200k(self):
        for model in ('claude-sonnet-4-20250514', 'claude-opus-4-20250514',
                      'claude-haiku-4-20250514', 'claude-haiku-4-5-20251001',
                      'claude-sonnet-4-4'):
            self.assertEqual(
                TokenCounter.get_context_limit(model),
                200_000,
                f"{model} should be 200k (Claude 4 base / pre-1M)",
            )

    def test_claude_3_resolves_to_200k(self):
        for model in ('claude-3-opus', 'claude-3.5-sonnet', 'claude-3-5-sonnet-20241022',
                      'claude-3-haiku-20240307'):
            self.assertEqual(TokenCounter.get_context_limit(model), 200_000)

    def test_future_claude_5_resolves_to_1m(self):
        # Future-proof: Claude 5.x assumed to inherit the 1M ceiling
        for model in ('claude-opus-5', 'claude-sonnet-5-2', 'claude-haiku-5-20270101'):
            self.assertEqual(TokenCounter.get_context_limit(model), 1_000_000)

    # ---- Gemini ----

    def test_gemini_3_resolves_to_1m(self):
        for model in ('gemini-3.1-pro-preview', 'gemini-3.5-flash', 'gemini-3.1-flash-lite',
                      'gemini-3.7-ultra'):
            self.assertEqual(TokenCounter.get_context_limit(model), 1_048_576)

    def test_gemini_2_resolves_to_1m(self):
        for model in ('gemini-2.0-flash', 'gemini-2.5-pro'):
            self.assertEqual(TokenCounter.get_context_limit(model), 1_048_576)

    def test_future_gemini_family_uses_1m_default(self):
        for model in ('gemini-4-pro', 'gemini-10-ultra'):
            self.assertEqual(TokenCounter.get_context_limit(model), 1_048_576)

    # ---- Provider fallback ----

    def test_unknown_openai_model_uses_provider_default(self):
        self.assertEqual(
            TokenCounter.get_context_limit('davinci-codex-007', provider='OpenAI'),
            128_000,
        )

    def test_unknown_claude_model_uses_provider_default(self):
        self.assertEqual(
            TokenCounter.get_context_limit('mystery-anthropic-model', provider='Claude'),
            200_000,
        )

    def test_patterns_are_anchored_to_model_id_start(self):
        """Family patterns must not match if the family token appears mid-string.

        A custom finetune ID like ``ft:org:gpt-4o:custom`` shouldn't be
        treated as a GPT-4o-class model by the regex pass. (It may still
        be resolved later via the legacy substring fallback — that's fine
        and intentional — but the regex stage must stay strict.)"""
        from icharlotte_core.chat.token_counter import MODEL_FAMILY_PATTERNS
        for pattern, _ in MODEL_FAMILY_PATTERNS:
            self.assertIsNone(
                pattern.match("ft:org:gpt-4o:custom"),
                f"pattern {pattern.pattern} should not match finetune-style ID",
            )


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
