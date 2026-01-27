"""
Token counting utilities for estimating context usage.
"""
from typing import Dict, Optional
import re


# Context window limits for various models
MODEL_CONTEXT_LIMITS = {
    # Gemini models
    'gemini-3-flash': 1000000,
    'gemini-3-flash-preview': 1000000,
    'gemini-3-pro': 1000000,
    'gemini-3-pro-preview': 1000000,
    'gemini-2.5-flash': 1000000,
    'gemini-2.5-pro': 1000000,
    'gemini-2.0-flash': 1000000,
    'gemini-1.5-flash': 1000000,
    'gemini-1.5-pro': 2000000,
    'gemini-1.0-pro': 32000,

    # OpenAI models
    'gpt-4o': 128000,
    'gpt-4o-mini': 128000,
    'gpt-4-turbo': 128000,
    'gpt-4-turbo-preview': 128000,
    'gpt-4': 8192,
    'gpt-4-32k': 32768,
    'gpt-3.5-turbo': 16385,
    'gpt-3.5-turbo-16k': 16385,
    'o1': 200000,
    'o1-preview': 128000,
    'o1-mini': 128000,
    'o3': 200000,
    'o3-mini': 200000,

    # Claude models
    'claude-3-opus': 200000,
    'claude-3-sonnet': 200000,
    'claude-3-haiku': 200000,
    'claude-3.5-sonnet': 200000,
    'claude-3.5-haiku': 200000,
    'claude-3-opus-20240229': 200000,
    'claude-3-sonnet-20240229': 200000,
    'claude-3-haiku-20240307': 200000,
    'claude-3-5-sonnet-20240620': 200000,
    'claude-3-5-sonnet-20241022': 200000,
    'claude-sonnet-4-20250514': 200000,
    'claude-opus-4-20250514': 200000,
}

# Default limits by provider
PROVIDER_DEFAULT_LIMITS = {
    'Gemini': 1000000,
    'OpenAI': 128000,
    'Claude': 200000,
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

        Args:
            model: Model name (e.g., 'gemini-3-flash-preview', 'gpt-4o')
            provider: Optional provider name for fallback

        Returns:
            Context window size in tokens
        """
        # Try exact match first
        if model in MODEL_CONTEXT_LIMITS:
            return MODEL_CONTEXT_LIMITS[model]

        # Try partial match (model name contains key)
        model_lower = model.lower()
        for key, limit in MODEL_CONTEXT_LIMITS.items():
            if key.lower() in model_lower:
                return limit

        # Fallback to provider default
        if provider and provider in PROVIDER_DEFAULT_LIMITS:
            return PROVIDER_DEFAULT_LIMITS[provider]

        # Ultimate fallback
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
        reserve_for_response: int = 4096
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
