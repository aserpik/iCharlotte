"""
Unit tests for word_assistant prompt assembly helpers and constants.

Covers:
- _is_placeholder_selection heuristic
- _render_insertion_instructions templating (placeholder + cursor)
- Shared system preamble is present in all three system prompts
- ELABORATION rule lives only in the system preamble (not duplicated)
- Redline prefix is appended (begins with whitespace separator)

Run with: python -m pytest tests/test_word_assistant_prompts.py -v
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core.word_hotkey import (
    _is_placeholder_selection,
    _render_insertion_instructions,
    _SHARED_SYSTEM_PREAMBLE,
    DEFAULT_WORD_SYSTEM_PROMPT,
    DEFAULT_WORD_REDLINE_SYSTEM_PROMPT,
    EMAIL_SYSTEM_PROMPT,
    DEFAULT_REDLINE_PREFIX,
    DEFAULT_INSERTION_INSTRUCTIONS,
    DEFAULT_SELECTION_INSTRUCTIONS,
    DEFAULT_PLACEHOLDER_INSTRUCTIONS,
    DEFAULT_CURSOR_INSTRUCTIONS,
)


# ── _is_placeholder_selection ───────────────────────────────────────────────

def test_empty_selection_is_placeholder():
    assert _is_placeholder_selection("") is True
    assert _is_placeholder_selection("   ") is True


def test_filler_chars_only_is_placeholder():
    assert _is_placeholder_selection("___") is True
    assert _is_placeholder_selection("___ ___") is True
    assert _is_placeholder_selection("---") is True
    assert _is_placeholder_selection("—") is True
    assert _is_placeholder_selection("–") is True
    assert _is_placeholder_selection("...") is True
    assert _is_placeholder_selection(". . .") is True
    assert _is_placeholder_selection("-_-") is True


def test_short_alphanumeric_is_NOT_placeholder():
    """Regression: previously len(stripped) <= 3 flagged these as blanks."""
    assert _is_placeholder_selection("Mr.") is False
    assert _is_placeholder_selection("Dr.") is False
    assert _is_placeholder_selection("$10") is False
    assert _is_placeholder_selection("USC") is False
    assert _is_placeholder_selection("X") is False
    assert _is_placeholder_selection("Q1") is False


def test_normal_selection_is_NOT_placeholder():
    assert _is_placeholder_selection("The defendant argued that") is False
    assert _is_placeholder_selection("On March 5, 2026") is False


# ── _render_insertion_instructions ──────────────────────────────────────────

def test_render_placeholder_mode():
    out = _render_insertion_instructions("placeholder")
    assert "position of the blank shown above" in out
    assert "the blank" in out
    # Cursor-only "complete sentence" rule must NOT appear in placeholder mode
    assert "complete, well-formed sentence" not in out


def test_render_cursor_mode():
    out = _render_insertion_instructions("cursor")
    assert "cursor position shown above" in out
    assert "the cursor" in out
    # Cursor mode SHOULD include the start-as-complete-sentence rule
    assert "complete, well-formed sentence" in out


def test_render_falls_back_on_broken_template():
    """If user customized prompt removed required placeholders, fall back to raw template."""
    broken = "Just some text without any placeholders"
    out = _render_insertion_instructions("placeholder", broken)
    assert out == broken


def test_render_falls_back_on_keyerror_template():
    """Template referencing an unknown placeholder should fall back to raw."""
    bad = "{unknown_placeholder} only"
    out = _render_insertion_instructions("cursor", bad)
    assert out == bad


def test_legacy_aliases_match_rendered_output():
    """Backward-compat aliases should equal what _render_insertion_instructions produces."""
    assert DEFAULT_PLACEHOLDER_INSTRUCTIONS == _render_insertion_instructions("placeholder")
    assert DEFAULT_CURSOR_INSTRUCTIONS == _render_insertion_instructions("cursor")


# ── System prompt composition ───────────────────────────────────────────────

def test_shared_preamble_in_all_system_prompts():
    """All three modes must inherit the shared preamble (voice/tone, attachment, etc.)."""
    assert _SHARED_SYSTEM_PREAMBLE in DEFAULT_WORD_SYSTEM_PROMPT
    assert _SHARED_SYSTEM_PREAMBLE in DEFAULT_WORD_REDLINE_SYSTEM_PROMPT
    assert _SHARED_SYSTEM_PREAMBLE in EMAIL_SYSTEM_PROMPT


def test_shared_preamble_contains_universal_rules():
    """The preamble must carry the rules that apply to every mode."""
    p = _SHARED_SYSTEM_PREAMBLE
    assert "VOICE & TONE RULE" in p
    assert "ATTACHMENT RULE" in p
    assert "ELABORATION RULE" in p
    assert "CONTINUATION RULE" in p
    assert "FORMAT BASELINE" in p


def test_mode_deltas_present():
    """Each mode prompt adds its own delta on top of the preamble."""
    assert "CONTEXTUAL WRITING" in DEFAULT_WORD_SYSTEM_PROMPT
    assert "REDLINE MODE" in DEFAULT_WORD_REDLINE_SYSTEM_PROMPT
    assert "EMAIL MODE" in EMAIL_SYSTEM_PROMPT


def test_elaboration_rule_not_duplicated_in_insertion_template():
    """Dedupe goal: the ELABORATION rule should live only in the preamble."""
    # The insertion template should NOT restate "USER DIRECTIVE is not the text"
    # (that's the system preamble's job now).
    assert "USER DIRECTIVE" not in DEFAULT_INSERTION_INSTRUCTIONS
    assert "USER DIRECTIVE" not in _render_insertion_instructions("placeholder")
    assert "USER DIRECTIVE" not in _render_insertion_instructions("cursor")


def test_selection_instructions_includes_fact_preservation():
    """Expanded #7: must call out preserving names/dates/facts verbatim."""
    s = DEFAULT_SELECTION_INSTRUCTIONS
    assert "PRESERVE FACTS VERBATIM" in s or "preserve" in s.lower()
    assert "names" in s.lower()
    assert "dates" in s.lower()


# ── Redline prefix placement ────────────────────────────────────────────────

def test_redline_prefix_starts_with_whitespace_separator():
    """
    Prefix is now appended at END of user message rather than prepended.
    It must start with a separator (newlines) so concatenation produces
    clean spacing rather than running into the prior text.
    """
    assert DEFAULT_REDLINE_PREFIX.startswith("\n\n")


def test_redline_prefix_carries_rules():
    p = DEFAULT_REDLINE_PREFIX
    assert "REDLINE MODE IS ACTIVE" in p
    assert "Track Changes" in p
    assert "paragraph" in p.lower()


if __name__ == "__main__":
    import pytest
    sys.exit(pytest.main([__file__, "-v"]))
