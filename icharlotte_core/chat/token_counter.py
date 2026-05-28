"""
Token counting utilities for estimating context usage.

Context-window lookup is layered so it stays accurate as new models ship:

  1. ``MODEL_CONTEXT_LIMITS``  — exact model-ID overrides (use sparingly,
     only when a specific model deviates from its family's typical window).
  2. ``MODEL_FAMILY_PATTERNS`` — ordered list of (regex, limit) tuples that
     catch whole model families (e.g. any ``gpt-5.*`` variant). New versions
     within a known family resolve correctly without code changes.
  3. ``PROVIDER_DEFAULT_LIMITS`` — last-resort fallback by provider name.

When a new model family launches (e.g. Claude 5, GPT-6, Gemini 4) add ONE
entry to ``MODEL_FAMILY_PATTERNS`` — that covers every version/snapshot in
the family. Only add to ``MODEL_CONTEXT_LIMITS`` for one-off exceptions.
"""
from typing import Dict, List, Optional, Pattern, Tuple
import re


# Exact model-ID overrides. Add entries here ONLY for models whose window
# differs from what their family pattern would say (e.g. legacy small variants).
MODEL_CONTEXT_LIMITS: Dict[str, int] = {
    # OpenAI legacy / small variants
    'gpt-4': 8192,
    'gpt-4-32k': 32768,
    'gpt-3.5-turbo': 16385,
    'gpt-3.5-turbo-16k': 16385,
    'gpt-3.5-turbo-instruct': 4096,
    'o1-mini': 128000,
    'o3-mini': 200000,
}


# Family pattern → context window, checked in order. First match wins, so
# put MORE-SPECIFIC patterns before more-general ones. Patterns are anchored
# at the start (``re.match``) so they only match real model-ID prefixes.
#
# Public context-window references (verified against provider docs as of
# 2026-05):
#   - GPT-4.1 family ........ 1,047,576 (1M)
#   - GPT-5 family (all) .... 400,000 (incl. Pro / Thinking / Instant)
#   - GPT-4o / GPT-4 Turbo .. 128,000
#   - o-series (o1/o3/o4) ... 200,000
#   - Claude Sonnet 4.5+ .... 1,000,000 (with 1M-context beta header)
#   - Claude Opus 4.6+ ...... 1,000,000 (with 1M-context beta header)
#   - Other Claude 3.x/4.x .. 200,000
#   - Gemini 2.x+ / 3.x ..... 1,048,576 (1M)
#
# IMPORTANT regex caveat: model IDs frequently carry an 8-digit date suffix
# (e.g. ``claude-sonnet-4-20250514``). A naive ``\d+`` for the minor version
# would swallow the date and mis-classify the model. We constrain minor
# versions to 1-2 digits via ``(?:[N-9]|[1-9]\d)`` followed by ``(?![0-9])``
# so an 8-digit date can never satisfy the version capture.
MODEL_FAMILY_PATTERNS: List[Tuple[Pattern[str], int]] = [
    # ---- OpenAI ----
    # GPT-4.1 family — 1M context
    (re.compile(r'^gpt-4\.1(?![0-9])', re.IGNORECASE), 1_047_576),
    # GPT-5 family and any future GPT-N where N>=5 — 400k context
    # Covers gpt-5, gpt-5-pro, gpt-5.1-thinking, gpt-5.5-pro, gpt-6, etc.
    (re.compile(r'^gpt-(?:[5-9]|[1-9]\d)(?![0-9])', re.IGNORECASE), 400_000),
    # GPT-4o, GPT-4o-mini, GPT-4 Turbo, chatgpt-* snapshots — 128k
    (re.compile(r'^(?:gpt-4o|gpt-4-turbo|chatgpt-)', re.IGNORECASE), 128_000),
    # o-series reasoning models (o1/o3/o4/...) — 200k (mini variants overridden above)
    (re.compile(r'^o[1-9](?![0-9])', re.IGNORECASE), 200_000),

    # ---- Claude ----
    # Claude Sonnet 4.5+ and Opus 4.6+ ship with optional 1M-context beta
    (re.compile(r'^claude-sonnet-4-(?:[5-9]|[1-9]\d)(?![0-9])', re.IGNORECASE), 1_000_000),
    (re.compile(r'^claude-opus-4-(?:[6-9]|[1-9]\d)(?![0-9])', re.IGNORECASE), 1_000_000),
    # Future Claude 5.x default (no public spec yet — assume 1M like 4.5+)
    (re.compile(r'^claude-(?:opus|sonnet|haiku)-(?:[5-9]|[1-9]\d)(?![0-9])', re.IGNORECASE), 1_000_000),
    # All other Claude 3.x / 4.x models → 200k
    (re.compile(r'^claude-(?:3|4|opus|sonnet|haiku)', re.IGNORECASE), 200_000),

    # ---- Gemini ----
    # Gemini 2.x / 3.x and beyond — 1M default
    # (iCharlotte does not support Gemini 1.x; see tests/test_gemini_model_catalog.py)
    (re.compile(r'^gemini-(?:[2-9]|[1-9]\d)(?![0-9])', re.IGNORECASE), 1_048_576),
]


# Provider-level fallback when neither exact match nor family pattern hits.
PROVIDER_DEFAULT_LIMITS: Dict[str, int] = {
    'Gemini': 1_048_576,
    'OpenAI': 128_000,
    'Claude': 200_000,
}


class TokenCounter:
    """Utilities for estimating token counts and managing context limits."""

    @staticmethod
    def estimate_tokens(text: str, provider: str = 'Gemini') -> int:
        """
        Estimate the number of tokens in a text string.

        Uses a simple character-based estimation (~4 chars per token for English).
        This is a rough approximation that works reasonably well for most use cases.

        For more accurate counts:
        - OpenAI: Use tiktoken library
        - Gemini: Use the countTokens API
        - Claude: Use the count_tokens API

        Args:
            text: The text to estimate tokens for
            provider: The LLM provider (affects estimation slightly)

        Returns:
            Estimated token count
        """
        if not text:
            return 0

        # Base estimation: ~4 characters per token for English
        # This is a commonly used approximation
        char_count = len(text)
        base_estimate = char_count / 4

        # Adjust for whitespace and punctuation
        # These typically count as separate tokens
        whitespace_count = len(re.findall(r'\s+', text))
        punct_count = len(re.findall(r'[.,!?;:\'"()\[\]{}]', text))

        # Code tends to have more tokens per character
        code_indicator = len(re.findall(r'[{}\[\]();=<>]', text))
        code_factor = 1.0 + (code_indicator / max(char_count, 1)) * 0.5

        # Final estimate with adjustments
        estimate = base_estimate * code_factor + whitespace_count * 0.3 + punct_count * 0.3

        return int(estimate)

    @staticmethod
    def estimate_tokens_detailed(text: str) -> Dict[str, int]:
        """
        Get a detailed token estimate breakdown.

        Returns:
            Dictionary with token estimates by category
        """
        if not text:
            return {'total': 0, 'text': 0, 'whitespace': 0, 'code': 0}

        char_count = len(text)
        text_tokens = char_count / 4

        # Whitespace
        whitespace = len(re.findall(r'\s+', text))

        # Code patterns
        code_patterns = len(re.findall(r'[{}\[\]();=<>]|def |class |function |import |from ', text))

        return {
            'total': int(text_tokens + whitespace * 0.3 + code_patterns * 0.5),
            'text': int(text_tokens),
            'whitespace': whitespace,
            'code': code_patterns
        }

    @staticmethod
    def get_context_limit(model: str, provider: str = None) -> int:
        """
        Get the context window limit for a model.

        Resolution order:
            1. Exact match in MODEL_CONTEXT_LIMITS
            2. Family-pattern match in MODEL_FAMILY_PATTERNS
            3. Substring match against MODEL_CONTEXT_LIMITS keys (legacy)
            4. Provider default in PROVIDER_DEFAULT_LIMITS
            5. 128k ultimate fallback

        Args:
            model: Model name (e.g., 'gemini-3.5-flash', 'gpt-5.5-pro')
            provider: Optional provider name for fallback

        Returns:
            Context window size in tokens
        """
        if not model:
            if provider and provider in PROVIDER_DEFAULT_LIMITS:
                return PROVIDER_DEFAULT_LIMITS[provider]
            return 128000

        # 1. Exact match
        if model in MODEL_CONTEXT_LIMITS:
            return MODEL_CONTEXT_LIMITS[model]

        # 2. Family-pattern match (handles new model versions automatically)
        for pattern, limit in MODEL_FAMILY_PATTERNS:
            if pattern.match(model):
                return limit

        # 3. Legacy substring match against exact-overrides dict
        model_lower = model.lower()
        for key, limit in MODEL_CONTEXT_LIMITS.items():
            if key.lower() in model_lower:
                return limit

        # 4. Provider default
        if provider and provider in PROVIDER_DEFAULT_LIMITS:
            return PROVIDER_DEFAULT_LIMITS[provider]

        # 5. Ultimate fallback
        return 128000

    @staticmethod
    def calculate_context_usage(
        messages: list,
        system_prompt: str = '',
        file_content: str = '',
        provider: str = 'Gemini'
    ) -> Dict[str, int]:
        """
        Calculate total context usage for a conversation.

        Args:
            messages: List of message dicts with 'content' key
            system_prompt: The system prompt
            file_content: Any attached file content
            provider: The LLM provider

        Returns:
            Dictionary with token counts and percentages
        """
        system_tokens = TokenCounter.estimate_tokens(system_prompt, provider)
        file_tokens = TokenCounter.estimate_tokens(file_content, provider)

        message_tokens = 0
        for msg in messages:
            content = msg.get('content', '') if isinstance(msg, dict) else str(msg)
            message_tokens += TokenCounter.estimate_tokens(content, provider)

        total = system_tokens + file_tokens + message_tokens

        return {
            'system_tokens': system_tokens,
            'file_tokens': file_tokens,
            'message_tokens': message_tokens,
            'total_tokens': total
        }

    @staticmethod
    def get_usage_percentage(
        total_tokens: int,
        model: str,
        provider: str = None
    ) -> float:
        """
        Get context usage as a percentage.

        Args:
            total_tokens: Total tokens used
            model: Model name
            provider: Optional provider for fallback

        Returns:
            Usage percentage (0.0 to 100.0+)
        """
        limit = TokenCounter.get_context_limit(model, provider)
        if limit == 0:
            return 0.0
        return (total_tokens / limit) * 100

    @staticmethod
    def format_token_count(count: int) -> str:
        """
        Format a token count for display.

        Args:
            count: Token count

        Returns:
            Formatted string (e.g., "1.2k", "1.5M")
        """
        if count >= 1000000:
            return f"{count / 1000000:.1f}M"
        elif count >= 1000:
            return f"{count / 1000:.1f}k"
        else:
            return str(count)

    @staticmethod
    def should_summarize(
        total_tokens: int,
        model: str,
        provider: str = None,
        threshold_percent: float = 80.0
    ) -> bool:
        """
        Check if conversation should be summarized based on context usage.

        Args:
            total_tokens: Total tokens used
            model: Model name
            provider: Optional provider for fallback
            threshold_percent: Percentage threshold to trigger summarization

        Returns:
            True if summarization is recommended
        """
        usage = TokenCounter.get_usage_percentage(total_tokens, model, provider)
        return usage >= threshold_percent

    @staticmethod
    def estimate_remaining_tokens(
        total_tokens: int,
        model: str,
        provider: str = None,
        reserve_for_response: int = 16384
    ) -> int:
        """
        Estimate remaining tokens available for new content.

        Args:
            total_tokens: Currently used tokens
            model: Model name
            provider: Optional provider for fallback
            reserve_for_response: Tokens to reserve for model response

        Returns:
            Available tokens for new input
        """
        limit = TokenCounter.get_context_limit(model, provider)
        remaining = limit - total_tokens - reserve_for_response
        return max(0, remaining)
