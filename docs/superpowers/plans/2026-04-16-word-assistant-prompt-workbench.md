# Word AI Assistant Prompt Workbench Integration — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make all 16 pipeline prompts (word assistant, legal research, mediation brief) editable from the existing Prompt Engineering Workbench via PromptManager, with versioning and fallback to hardcoded defaults.

**Architecture:** Each hardcoded prompt constant gets a companion getter function that checks PromptManager first, falling back to the constant. A seed function writes all defaults as v1 on first Workbench open. The Workbench UI adds three new agents to its dropdown.

**Tech Stack:** Python, PromptManager (existing), PromptsDialog (existing PyQt6 Workbench)

**Spec:** `docs/superpowers/specs/2026-04-16-word-assistant-prompt-workbench-design.md`

---

### Task 1: Add seed_pipeline_prompts() to PromptManager

**Files:**
- Modify: `icharlotte_core/prompt_manager.py:96-105`
- Test: `tests/test_prompt_manager_seed.py` (create)

- [ ] **Step 1: Write the failing test**

Create `tests/test_prompt_manager_seed.py`:

```python
"""Tests for seed_pipeline_prompts() in PromptManager."""
import os
import shutil
import tempfile
import unittest

from icharlotte_core.prompt_manager import PromptManager


class TestSeedPipelinePrompts(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pm = PromptManager(prompts_dir=self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_seed_creates_word_assistant_passes(self):
        self.pm.seed_pipeline_prompts()
        expected = [
            "system_prompt", "redline_system_prompt", "email_system_prompt",
            "redline_prefix", "placeholder_instructions",
            "cursor_instructions", "selection_instructions",
        ]
        for pass_name in expected:
            text = self.pm.get_prompt("word_assistant", pass_name)
            self.assertIsNotNone(text, f"word_assistant:{pass_name} missing")
            self.assertTrue(len(text) > 20, f"word_assistant:{pass_name} too short")

    def test_seed_creates_legal_research_passes(self):
        self.pm.seed_pipeline_prompts()
        expected = [
            "query_planning", "query_extraction", "synthesis",
            "verification", "relevance_ranking",
            "research_framing", "citation_instruction",
        ]
        for pass_name in expected:
            text = self.pm.get_prompt("legal_research", pass_name)
            self.assertIsNotNone(text, f"legal_research:{pass_name} missing")
            self.assertTrue(len(text) > 20, f"legal_research:{pass_name} too short")

    def test_seed_creates_mediation_brief_passes(self):
        self.pm.seed_pipeline_prompts()
        for pass_name in ["style_guide", "formatting_rules"]:
            text = self.pm.get_prompt("mediation_brief", pass_name)
            self.assertIsNotNone(text, f"mediation_brief:{pass_name} missing")
            self.assertTrue(len(text) > 20, f"mediation_brief:{pass_name} too short")

    def test_seed_is_idempotent(self):
        self.pm.seed_pipeline_prompts()
        v1_text = self.pm.get_prompt("word_assistant", "system_prompt")
        # Seed again — should not overwrite
        self.pm.seed_pipeline_prompts()
        v1_text_again = self.pm.get_prompt("word_assistant", "system_prompt")
        self.assertEqual(v1_text, v1_text_again)
        # Should still only have one version
        versions = self.pm.list_versions("word_assistant", "system_prompt")
        self.assertEqual(len(versions), 1)

    def test_seed_does_not_overwrite_user_edits(self):
        self.pm.seed_pipeline_prompts()
        # Simulate user edit
        self.pm.create_version(
            "word_assistant", "system_prompt",
            "My custom system prompt",
            version="v2", set_as_current=True,
        )
        custom = self.pm.get_prompt("word_assistant", "system_prompt")
        self.assertEqual(custom, "My custom system prompt")
        # Seed again — should NOT revert to default
        self.pm.seed_pipeline_prompts()
        after_seed = self.pm.get_prompt("word_assistant", "system_prompt")
        self.assertEqual(after_seed, "My custom system prompt")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_prompt_manager_seed.py -v`
Expected: FAIL with `AttributeError: 'PromptManager' object has no attribute 'seed_pipeline_prompts'`

- [ ] **Step 3: Add new agent dirs to _ensure_directory_structure**

In `icharlotte_core/prompt_manager.py`, change line 102 from:

```python
        for agent in ['summarize', 'discovery', 'deposition', 'timeline', 'contradiction']:
```

to:

```python
        for agent in ['summarize', 'discovery', 'deposition', 'timeline', 'contradiction',
                      'word_assistant', 'legal_research', 'mediation_brief']:
```

- [ ] **Step 4: Add seed_pipeline_prompts() method**

Add this method to the `PromptManager` class in `icharlotte_core/prompt_manager.py`, right after the `migrate_legacy_prompts` method (after line 420):

```python
    def seed_pipeline_prompts(self):
        """Seed all pipeline prompts (word assistant, legal research, mediation brief).

        Writes each hardcoded default as v1 if no version exists yet.
        Idempotent — safe to call multiple times; never overwrites user edits.
        """
        from icharlotte_core.word_hotkey import (
            DEFAULT_WORD_SYSTEM_PROMPT,
            DEFAULT_WORD_REDLINE_SYSTEM_PROMPT,
            EMAIL_SYSTEM_PROMPT,
            DEFAULT_REDLINE_PREFIX,
            DEFAULT_PLACEHOLDER_INSTRUCTIONS,
            DEFAULT_CURSOR_INSTRUCTIONS,
            DEFAULT_SELECTION_INSTRUCTIONS,
        )
        from icharlotte_core.legal_research.prompts import (
            QUERY_PLANNING_PROMPT,
            QUERY_EXTRACTION_PROMPT,
            SYNTHESIS_PROMPT,
            VERIFICATION_PROMPT,
            RELEVANCE_RANKING_PROMPT,
            RESEARCH_FRAMING_INSTRUCTION,
            CITATION_INSTRUCTION,
        )
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        seeds = [
            # word_assistant
            ("word_assistant", "system_prompt", DEFAULT_WORD_SYSTEM_PROMPT, "Default Word system prompt"),
            ("word_assistant", "redline_system_prompt", DEFAULT_WORD_REDLINE_SYSTEM_PROMPT, "Default redline system prompt"),
            ("word_assistant", "email_system_prompt", EMAIL_SYSTEM_PROMPT, "Default Outlook email system prompt"),
            ("word_assistant", "redline_prefix", DEFAULT_REDLINE_PREFIX, "Prefix prepended in redline mode"),
            ("word_assistant", "placeholder_instructions", DEFAULT_PLACEHOLDER_INSTRUCTIONS, "Instructions for filling blank/placeholder"),
            ("word_assistant", "cursor_instructions", DEFAULT_CURSOR_INSTRUCTIONS, "Instructions for cursor-position insertion"),
            ("word_assistant", "selection_instructions", DEFAULT_SELECTION_INSTRUCTIONS, "Instructions for selected text with full doc"),
            # legal_research
            ("legal_research", "query_planning", QUERY_PLANNING_PROMPT, "Structured JSON query generation"),
            ("legal_research", "query_extraction", QUERY_EXTRACTION_PROMPT, "Extract queries from litigation prompt"),
            ("legal_research", "synthesis", SYNTHESIS_PROMPT, "Synthesize authorities into memo"),
            ("legal_research", "verification", VERIFICATION_PROMPT, "Citation verification"),
            ("legal_research", "relevance_ranking", RELEVANCE_RANKING_PROMPT, "Case relevance ranking"),
            ("legal_research", "research_framing", RESEARCH_FRAMING_INSTRUCTION, "Citation requirements for user prompt"),
            ("legal_research", "citation_instruction", CITATION_INSTRUCTION, "Strict citation rules"),
            # mediation_brief
            ("mediation_brief", "style_guide", MediationBriefGenerator.STYLE_GUIDE, "Defense writing style/tone guide"),
            ("mediation_brief", "formatting_rules", MediationBriefGenerator.FORMATTING_RULES, "Structural formatting rules"),
        ]

        seeded = 0
        for agent, pass_name, content, description in seeds:
            key = self._get_prompt_key(agent, pass_name)
            if key in self._registry.get("prompts", {}):
                continue  # Already exists — do not overwrite
            # Also check if the file already exists on disk
            if os.path.exists(self._get_current_path(agent, pass_name)):
                continue
            self.create_version(
                agent, pass_name, content.strip(),
                version="v1",
                description=description,
                author="system",
                set_as_current=True,
            )
            seeded += 1

        if seeded:
            print(f"[PromptManager] Seeded {seeded} pipeline prompts")
        return seeded
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests/test_prompt_manager_seed.py -v`
Expected: tests will fail because `DEFAULT_REDLINE_PREFIX`, `DEFAULT_PLACEHOLDER_INSTRUCTIONS`, `DEFAULT_CURSOR_INSTRUCTIONS`, `DEFAULT_SELECTION_INSTRUCTIONS` don't exist in word_hotkey.py yet. That's expected — we create them in Task 2. For now, verify the test infrastructure is sound by confirming the import error is from word_hotkey, not from prompt_manager itself.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/prompt_manager.py tests/test_prompt_manager_seed.py
git commit -m "feat(prompt-manager): add seed_pipeline_prompts() and pipeline agent dirs"
```

---

### Task 2: Extract instruction constants from word_hotkey.py

**Files:**
- Modify: `icharlotte_core/word_hotkey.py:170-230` (constants area) and `~3981, ~4020-4095` (inline blocks)

This task extracts the four inline instruction blocks into named constants at the top of the file, right next to the existing system prompt constants. No behavior change — just extraction.

- [ ] **Step 1: Add DEFAULT_REDLINE_PREFIX constant**

In `icharlotte_core/word_hotkey.py`, after the `EMAIL_SYSTEM_PROMPT` constant (after line ~230), add:

```python
DEFAULT_REDLINE_PREFIX = (
    "IMPORTANT — REDLINE MODE IS ACTIVE: Your output will be compared "
    "word-by-word against the original text to generate Track Changes. "
    "You MUST follow these rules:\n"
    "1. Preserve any text that does not need to change EXACTLY as-is — "
    "same wording, same sentence structure, same word order.\n"
    "2. PRESERVE THE EXACT PARAGRAPH STRUCTURE — keep the same number of "
    "paragraphs and blank lines between them. Each paragraph must start "
    "and end at the same boundaries as the original. Do NOT merge paragraphs, "
    "split paragraphs, or remove blank lines.\n"
    "3. Do NOT rephrase, reorganize, or rewrite portions that are already correct.\n"
    "4. Only modify the specific words or sentences that need to change.\n"
    "5. If a sentence is fine as-is, copy it verbatim.\n\n"
)
```

- [ ] **Step 2: Add DEFAULT_PLACEHOLDER_INSTRUCTIONS constant**

Immediately after DEFAULT_REDLINE_PREFIX, add:

```python
DEFAULT_PLACEHOLDER_INSTRUCTIONS = (
    "CRITICAL INSTRUCTIONS:\n"
    "- The USER DIRECTIVE above tells you WHAT to write about — it is "
    "NOT the text to insert. You must compose original, substantive "
    "prose that develops the ideas described. NEVER copy or rephrase "
    "the user's instruction as your output.\n"
    "- Your output will be INSERTED DIRECTLY into the document at "
    "the position of the blank shown above.\n"
    "- Write ACTUAL PROSE that flows naturally from the text before "
    "the blank and connects to the text after it.\n"
    "- You are ghostwriting as the document's author. Match their "
    'voice, tone, and person (e.g., if they write "we", you write "we").\n'
    "- Do NOT write instructions, action items, task descriptions, "
    "or summaries. Write the actual words that belong in the document.\n"
    "- Output ONLY the replacement text — no preamble, no explanation."
)
```

- [ ] **Step 3: Add DEFAULT_CURSOR_INSTRUCTIONS constant**

```python
DEFAULT_CURSOR_INSTRUCTIONS = (
    "CRITICAL INSTRUCTIONS:\n"
    "- The USER DIRECTIVE above tells you WHAT to write about — it is "
    "NOT the text to insert. You must compose original, substantive "
    "prose that develops the ideas described. NEVER copy or rephrase "
    "the user's instruction as your output.\n"
    "- Your output will be INSERTED DIRECTLY into the document at "
    "the cursor position shown above.\n"
    "- Write ACTUAL PROSE that flows naturally from the text before "
    "the cursor and connects to the text after it.\n"
    "- You are ghostwriting as the document's author. Match their "
    'voice, tone, and person (e.g., if they write "we", you write "we").\n'
    "- Do NOT write instructions, action items, task descriptions, "
    "or meta-commentary. Write the actual words that belong in the document.\n"
    "- Your output must start as a complete, well-formed sentence or "
    "paragraph. NEVER begin with a partial word or sentence fragment.\n"
    "- Output ONLY the text to insert — no preamble, no explanation."
)
```

- [ ] **Step 4: Add DEFAULT_SELECTION_INSTRUCTIONS constant**

```python
DEFAULT_SELECTION_INSTRUCTIONS = (
    "Apply the user's instructions to the SELECTED TEXT above. "
    "Use the full document for context and to match the writing style. "
    "Output only the processed version of the selected text."
)
```

- [ ] **Step 5: Replace inline redline prefix with constant**

Find the inline redline prefix block (around line 3981). Replace:

```python
            if self._redline_mode_active:
                redline_prefix = (
                    "IMPORTANT — REDLINE MODE IS ACTIVE: Your output will be compared "
                    "word-by-word against the original text to generate Track Changes. "
                    "You MUST follow these rules:\n"
                    "1. Preserve any text that does not need to change EXACTLY as-is — "
                    "same wording, same sentence structure, same word order.\n"
                    "2. PRESERVE THE EXACT PARAGRAPH STRUCTURE — keep the same number of "
                    "paragraphs and blank lines between them. Each paragraph must start "
                    "and end at the same boundaries as the original. Do NOT merge paragraphs, "
                    "split paragraphs, or remove blank lines.\n"
                    "3. Do NOT rephrase, reorganize, or rewrite portions that are already correct.\n"
                    "4. Only modify the specific words or sentences that need to change.\n"
                    "5. If a sentence is fine as-is, copy it verbatim.\n\n"
                )
```

with:

```python
            if self._redline_mode_active:
                redline_prefix = DEFAULT_REDLINE_PREFIX
```

- [ ] **Step 6: Replace inline placeholder instructions with constant**

In the placeholder branch (around line 4020), replace the entire f-string that builds `full_prompt` with one that uses the constant for the instructions tail:

```python
                        # Placeholder/blank: instruct AI to write replacement prose
                        full_prompt = (
                            f"=== USER DIRECTIVE (describes what to write — NOT the text itself) ===\n"
                            f"{prompt}\n\n"
                            f"=== FULL DOCUMENT ===\n{all_text}\n\n"
                            f"=== IMMEDIATE CONTEXT ===\n"
                            f"Text BEFORE the blank: \"{context_before}\"\n"
                            f"[___BLANK TO FILL___]\n"
                            f"Text AFTER the blank: \"{context_after}\"\n\n"
                            f"{DEFAULT_PLACEHOLDER_INSTRUCTIONS}"
                        )
```

- [ ] **Step 7: Replace inline cursor instructions with constant**

In the cursor-only branch (around line 4071), replace the f-string:

```python
                        full_prompt = (
                            f"=== USER DIRECTIVE (describes what to write — NOT the text itself) ===\n"
                            f"{prompt}\n\n"
                            f"=== FULL DOCUMENT ===\n{all_text}\n\n"
                            f"=== CURSOR POSITION ===\n"
                            f"Text BEFORE cursor: \"{context_before}\"\n"
                            f"[___CURSOR IS HERE___]\n"
                            f"Text AFTER cursor: \"{context_after}\"\n\n"
                            f"{DEFAULT_CURSOR_INSTRUCTIONS}"
                        )
```

- [ ] **Step 8: Replace inline selection instructions with constant**

In the normal selection branch (around line 4045), replace:

```python
                    else:
                        # Normal selection: process/transform the selected text
                        full_prompt = (
                            f"{prompt}\n\n"
                            f"=== FULL DOCUMENT (for context) ===\n{all_text}\n\n"
                            f"=== SELECTED TEXT TO PROCESS ===\n{selected_text}\n\n"
                            f"{DEFAULT_SELECTION_INSTRUCTIONS}"
                        )
```

- [ ] **Step 9: Verify syntax**

Run: `python -m py_compile icharlotte_core/word_hotkey.py && echo OK`
Expected: `OK`

- [ ] **Step 10: Run existing tests to verify no regressions**

Run: `python -m pytest tests/test_quote_insertion.py tests/test_mediation_brief_live.py -v`
Expected: All pass

- [ ] **Step 11: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "refactor(word-hotkey): extract inline instruction blocks into named constants"
```

---

### Task 3: Wire word_hotkey.py to load prompts from PromptManager

**Files:**
- Modify: `icharlotte_core/word_hotkey.py` — system prompt loading (~line 4140), redline prefix (~line 3981), placeholder/cursor/selection branches, and the popup System Prompt tab

- [ ] **Step 1: Add get_prompt import near top of word_hotkey.py**

After the existing imports (around line 36, after the PySide6 imports), add:

```python
from icharlotte_core.prompt_manager import get_prompt as _get_workbench_prompt
```

Use the alias `_get_workbench_prompt` to avoid shadowing any local `get_prompt` usage.

- [ ] **Step 2: Create _load_pipeline_prompt helper**

Add this helper function right after the existing `_force_restore_screen_updating` function (around line 125):

```python
def _load_pipeline_prompt(agent: str, pass_name: str, default: str) -> str:
    """Load a prompt from the Workbench (PromptManager), falling back to default."""
    try:
        custom = _get_workbench_prompt(agent, pass_name)
        if custom:
            return custom
    except Exception:
        pass
    return default
```

- [ ] **Step 3: Wire system prompt loading**

In `_do_execute`, find the system prompt selection block (around line 4138):

```python
            if self._redline_mode_active:
                system_prompt = custom_sp if custom_sp else DEFAULT_WORD_REDLINE_SYSTEM_PROMPT
            else:
                system_prompt = custom_sp if custom_sp else DEFAULT_WORD_SYSTEM_PROMPT
```

Replace with:

```python
            if self._redline_mode_active:
                system_prompt = custom_sp if custom_sp else _load_pipeline_prompt(
                    "word_assistant", "redline_system_prompt", DEFAULT_WORD_REDLINE_SYSTEM_PROMPT)
            else:
                system_prompt = custom_sp if custom_sp else _load_pipeline_prompt(
                    "word_assistant", "system_prompt", DEFAULT_WORD_SYSTEM_PROMPT)
```

- [ ] **Step 4: Wire redline prefix**

In `_do_execute`, find the redline prefix line (around line 3981). Replace:

```python
                redline_prefix = DEFAULT_REDLINE_PREFIX
```

with:

```python
                redline_prefix = _load_pipeline_prompt(
                    "word_assistant", "redline_prefix", DEFAULT_REDLINE_PREFIX)
```

- [ ] **Step 5: Wire placeholder instructions**

In the placeholder branch, replace `{DEFAULT_PLACEHOLDER_INSTRUCTIONS}` in the f-string with a variable:

Before the `full_prompt` f-string, add:

```python
                        _ph_instr = _load_pipeline_prompt(
                            "word_assistant", "placeholder_instructions", DEFAULT_PLACEHOLDER_INSTRUCTIONS)
```

Then in the f-string, change the last line from `f"{DEFAULT_PLACEHOLDER_INSTRUCTIONS}"` to `f"{_ph_instr}"`.

- [ ] **Step 6: Wire cursor instructions**

Same pattern for cursor branch. Before the f-string, add:

```python
                        _cur_instr = _load_pipeline_prompt(
                            "word_assistant", "cursor_instructions", DEFAULT_CURSOR_INSTRUCTIONS)
```

Then change `f"{DEFAULT_CURSOR_INSTRUCTIONS}"` to `f"{_cur_instr}"`.

- [ ] **Step 7: Wire selection instructions**

Same pattern for selection branch. Before the f-string, add:

```python
                        _sel_instr = _load_pipeline_prompt(
                            "word_assistant", "selection_instructions", DEFAULT_SELECTION_INSTRUCTIONS)
```

Then change `f"{DEFAULT_SELECTION_INSTRUCTIONS}"` to `f"{_sel_instr}"`.

- [ ] **Step 8: Wire email system prompt**

Find the Outlook system prompt loading in `_do_execute_outlook` (around line 5639). Replace:

```python
                system_prompt = custom_sp if custom_sp else DEFAULT_WORD_SYSTEM_PROMPT
```

and the EMAIL_SYSTEM_PROMPT reference with:

```python
                system_prompt = custom_sp if custom_sp else _load_pipeline_prompt(
                    "word_assistant", "email_system_prompt", EMAIL_SYSTEM_PROMPT)
```

(Search for all references to `EMAIL_SYSTEM_PROMPT` and `DEFAULT_WORD_SYSTEM_PROMPT` / `DEFAULT_WORD_REDLINE_SYSTEM_PROMPT` in the Outlook path and apply the same pattern.)

- [ ] **Step 9: Rewire popup System Prompt tab to use PromptManager**

In `_load_custom_system_prompt` (line 2857), replace the JSON file approach:

```python
    def _load_custom_system_prompt(self) -> str:
        """Load custom system prompt from PromptManager. Returns empty string if using default."""
        try:
            workbench = _get_workbench_prompt("word_assistant", "system_prompt")
            if workbench and workbench.strip() != DEFAULT_WORD_SYSTEM_PROMPT.strip():
                return workbench
        except Exception:
            pass
        # Legacy fallback: check old JSON file
        if os.path.exists(self.system_prompt_path):
            try:
                with open(self.system_prompt_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get("system_prompt", "")
            except Exception as e:
                print(f"Error loading custom system prompt: {e}")
        return ""
```

In `_save_system_prompt_from_ui` (line 2868), add PromptManager save alongside the JSON file:

After `self._custom_system_prompt = text`, add:

```python
        # Also save to PromptManager for Workbench sync
        try:
            from icharlotte_core.prompt_manager import get_prompt_manager
            pm = get_prompt_manager()
            pm.create_version(
                "word_assistant", "system_prompt", text,
                description="Saved from Word popup",
                set_as_current=True,
            )
        except Exception as e:
            print(f"Error saving system prompt to PromptManager: {e}")
```

- [ ] **Step 10: Verify syntax**

Run: `python -m py_compile icharlotte_core/word_hotkey.py && echo OK`
Expected: `OK`

- [ ] **Step 11: Run existing tests**

Run: `python -m pytest tests/test_quote_insertion.py tests/test_mediation_brief_live.py -v`
Expected: All pass

- [ ] **Step 12: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat(word-hotkey): load all prompts from PromptManager with fallback"
```

---

### Task 4: Wire legal_research/prompts.py to load from PromptManager

**Files:**
- Modify: `icharlotte_core/legal_research/prompts.py`
- Modify: `icharlotte_core/legal_research/engine.py`
- Modify: `icharlotte_core/word_hotkey.py:1228-1232` (import site)

- [ ] **Step 1: Add getter functions to prompts.py**

At the bottom of `icharlotte_core/legal_research/prompts.py` (after the `build_augmented_system_prompt` function), add:

```python
# ---------------------------------------------------------------------------
# Workbench-aware getters — check PromptManager first, fall back to constant
# ---------------------------------------------------------------------------

def _load(pass_name: str, default: str) -> str:
    try:
        from icharlotte_core.prompt_manager import get_prompt
        custom = get_prompt("legal_research", pass_name)
        if custom:
            return custom
    except Exception:
        pass
    return default


def get_query_planning_prompt() -> str:
    return _load("query_planning", QUERY_PLANNING_PROMPT)


def get_query_extraction_prompt() -> str:
    return _load("query_extraction", QUERY_EXTRACTION_PROMPT)


def get_synthesis_prompt() -> str:
    return _load("synthesis", SYNTHESIS_PROMPT)


def get_verification_prompt() -> str:
    return _load("verification", VERIFICATION_PROMPT)


def get_relevance_ranking_prompt() -> str:
    return _load("relevance_ranking", RELEVANCE_RANKING_PROMPT)


def get_research_framing_instruction() -> str:
    return _load("research_framing", RESEARCH_FRAMING_INSTRUCTION)


def get_citation_instruction() -> str:
    return _load("citation_instruction", CITATION_INSTRUCTION)
```

- [ ] **Step 2: Update engine.py imports**

In `icharlotte_core/legal_research/engine.py`, find the import block (lines 11-14):

```python
    QUERY_PLANNING_PROMPT,
    RELEVANCE_RANKING_PROMPT,
    SYNTHESIS_PROMPT,
    VERIFICATION_PROMPT,
```

Replace with:

```python
    get_query_planning_prompt,
    get_relevance_ranking_prompt,
    get_synthesis_prompt,
    get_verification_prompt,
```

- [ ] **Step 3: Update engine.py usage sites**

In `engine.py`, replace each constant reference with a function call:

- Line 130: `SYNTHESIS_PROMPT` → `get_synthesis_prompt()`
- Line 168: `QUERY_PLANNING_PROMPT` → `get_query_planning_prompt()`
- Line 375: `RELEVANCE_RANKING_PROMPT` → `get_relevance_ranking_prompt()`
- Line 479: `VERIFICATION_PROMPT` → `get_verification_prompt()`

- [ ] **Step 4: Update word_hotkey.py legal research imports**

In `icharlotte_core/word_hotkey.py`, find the import block in `_run_legal_research` (around line 1228):

```python
            from icharlotte_core.legal_research.prompts import (
                QUERY_EXTRACTION_PROMPT,
                CITATION_INSTRUCTION,
                RESEARCH_FRAMING_INSTRUCTION,
            )
```

Replace with:

```python
            from icharlotte_core.legal_research.prompts import (
                get_query_extraction_prompt,
                get_citation_instruction,
                get_research_framing_instruction,
            )
```

- [ ] **Step 5: Update word_hotkey.py usage sites in _run_legal_research**

In the same method, replace:
- Line 1253: `QUERY_EXTRACTION_PROMPT, self.task_data.full_prompt[:8000]` → `get_query_extraction_prompt(), self.task_data.full_prompt[:8000]`
- Line 1288: `f"{RESEARCH_FRAMING_INSTRUCTION}\n\n"` → `f"{get_research_framing_instruction()}\n\n"`
- Line 1292: `f"{CITATION_INSTRUCTION}"` → `f"{get_citation_instruction()}"` 
- Line 1298: `f"{CITATION_INSTRUCTION}"` → `f"{get_citation_instruction()}"`

- [ ] **Step 6: Update build_augmented_system_prompt to use getters**

In `icharlotte_core/legal_research/prompts.py`, find `build_augmented_system_prompt` (line 245). It directly uses `RESEARCH_FRAMING_INSTRUCTION` and `CITATION_INSTRUCTION`. Replace those two references with getter calls:

```python
def build_augmented_system_prompt(
    base_system_prompt: str,
    authority_block: str,
    research_memo: str = "",
) -> str:
    parts = [base_system_prompt]

    if research_memo:
        parts.append("")
        parts.append(get_research_framing_instruction())

    parts.append("")
    parts.append(get_citation_instruction())

    if research_memo:
        parts.append("")
        parts.append("[LEGAL RESEARCH MEMO]")
        parts.append(research_memo)
        parts.append("[/LEGAL RESEARCH MEMO]")

    parts.append("")
    parts.append(f"[LEGAL AUTHORITY]\n{authority_block}")

    return "\n".join(parts)
```

- [ ] **Step 7: Verify syntax for all files**

Run: `python -m py_compile icharlotte_core/legal_research/prompts.py && python -m py_compile icharlotte_core/legal_research/engine.py && python -m py_compile icharlotte_core/word_hotkey.py && echo OK`
Expected: `OK`

- [ ] **Step 8: Commit**

```bash
git add icharlotte_core/legal_research/prompts.py icharlotte_core/legal_research/engine.py icharlotte_core/word_hotkey.py
git commit -m "feat(legal-research): load prompts from PromptManager with fallback"
```

---

### Task 5: Wire mediation_brief.py to load from PromptManager

**Files:**
- Modify: `icharlotte_core/mediation_brief.py:158-186, 435`

- [ ] **Step 1: Update _build_system_prompt to load from PromptManager**

In `icharlotte_core/mediation_brief.py`, modify `_build_system_prompt` (line 423) from:

```python
    def _build_system_prompt(self, section_name: str) -> str:
        """Return the hard-coded system prompt for the given section.

        Combines the style guide, formatting rules, and a defense attorney
        persona statement.
        """
        persona = (
            "You are a senior defense litigation attorney with decades of trial experience. "
            "You are writing a confidential mediation brief on behalf of the defendant. "
            "Your goal is to present the strongest possible defense position to the mediator, "
            "supported by the evidence and the facts of the case."
        )
        return f"{persona}\n{self.STYLE_GUIDE}\n{self.FORMATTING_RULES}"
```

to:

```python
    def _build_system_prompt(self, section_name: str) -> str:
        """Return the system prompt for the given section.

        Loads style guide and formatting rules from the Workbench
        (PromptManager) if edited, otherwise uses class constants.
        """
        persona = (
            "You are a senior defense litigation attorney with decades of trial experience. "
            "You are writing a confidential mediation brief on behalf of the defendant. "
            "Your goal is to present the strongest possible defense position to the mediator, "
            "supported by the evidence and the facts of the case."
        )
        try:
            from icharlotte_core.prompt_manager import get_prompt
            style = get_prompt("mediation_brief", "style_guide") or self.STYLE_GUIDE
            formatting = get_prompt("mediation_brief", "formatting_rules") or self.FORMATTING_RULES
        except Exception:
            style = self.STYLE_GUIDE
            formatting = self.FORMATTING_RULES
        return f"{persona}\n{style}\n{formatting}"
```

- [ ] **Step 2: Verify syntax**

Run: `python -m py_compile icharlotte_core/mediation_brief.py && echo OK`
Expected: `OK`

- [ ] **Step 3: Run mediation brief tests**

Run: `python -m pytest tests/test_mediation_brief_live.py -v`
Expected: All pass

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/mediation_brief.py
git commit -m "feat(mediation-brief): load style_guide and formatting_rules from PromptManager"
```

---

### Task 6: Register new agents in the Workbench UI

**Files:**
- Modify: `icharlotte_core/ui/dialogs.py:388-400, 1443-1447, 1465-1473`

- [ ] **Step 1: Add new agents to WORKBENCH_TO_AGENT_ID**

In `icharlotte_core/ui/dialogs.py`, find `WORKBENCH_TO_AGENT_ID` (line 388). Add three entries:

```python
WORKBENCH_TO_AGENT_ID = {
    "summarize": "agent_summarize",
    "discovery": "agent_sum_disc",
    "deposition": "agent_sum_depo",
    "liability": "agent_liability",
    "exposure": "agent_exposure",
    "med_record": "agent_med_rec",
    "med_chron": "agent_med_chron",
    "extraction": "agent_separate",
    "email_update": "func_email_compose",
    "chat": "func_chat",
    "mediation_brief": "agent_mediation_brief",
    "word_assistant": "func_word_assistant",
    "legal_research": "func_legal_research",
}
```

- [ ] **Step 2: Add new agents to predefined list in _populate_agents**

Find the predefined agents list (line 1444). Change from:

```python
        for agent in ['summarize', 'discovery', 'deposition',
                      'liability', 'exposure', 'med_record', 'med_chron', 'extraction',
                      'email_update', 'chat']:
```

to:

```python
        for agent in ['summarize', 'discovery', 'deposition',
                      'liability', 'exposure', 'med_record', 'med_chron', 'extraction',
                      'email_update', 'chat',
                      'word_assistant', 'legal_research', 'mediation_brief']:
```

- [ ] **Step 3: Add seed call to _migrate_if_needed**

Find `_migrate_if_needed` (line 1465). Add the seed call after the existing migration:

```python
    def _migrate_if_needed(self):
        """Run migration if the registry doesn't exist."""
        if not os.path.exists(os.path.join(PROMPTS_DIR, "registry.json")):
            try:
                migrated = self.prompt_manager.migrate_legacy_prompts()
                if migrated > 0:
                    log_event(f"Migrated {migrated} legacy prompts to versioned storage")
            except Exception as e:
                log_event(f"Error migrating prompts: {e}", "error")

        # Seed pipeline prompts (word assistant, legal research, mediation brief)
        try:
            seeded = self.prompt_manager.seed_pipeline_prompts()
            if seeded and seeded > 0:
                log_event(f"Seeded {seeded} pipeline prompts in Workbench")
        except Exception as e:
            log_event(f"Error seeding pipeline prompts: {e}", "error")
```

- [ ] **Step 4: Verify syntax**

Run: `python -m py_compile icharlotte_core/ui/dialogs.py && echo OK`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/dialogs.py
git commit -m "feat(workbench): register word_assistant, legal_research, mediation_brief agents"
```

---

### Task 7: Run seed tests end-to-end and verify

**Files:**
- Test: `tests/test_prompt_manager_seed.py` (already created in Task 1)

- [ ] **Step 1: Run the seed tests**

Run: `python -m pytest tests/test_prompt_manager_seed.py -v`
Expected: All 5 tests pass

- [ ] **Step 2: Run all existing test suites for regressions**

Run: `python -m pytest tests/test_quote_insertion.py tests/test_mediation_brief_live.py tests/test_mediation_brief_quote_search.py -v`
Expected: All pass

- [ ] **Step 3: Verify full import chain works**

Run: `python -c "from icharlotte_core.prompt_manager import get_prompt_manager; pm = get_prompt_manager(); pm.seed_pipeline_prompts(); print('Seeded OK'); text = pm.get_prompt('word_assistant', 'system_prompt'); print(f'system_prompt: {len(text)} chars'); text2 = pm.get_prompt('legal_research', 'query_planning'); print(f'query_planning: {len(text2)} chars'); text3 = pm.get_prompt('mediation_brief', 'style_guide'); print(f'style_guide: {len(text3)} chars')"`

Expected: Three lines showing character counts > 100 for each.

- [ ] **Step 4: Verify word_hotkey loads from PromptManager**

Run: `python -c "from icharlotte_core.word_hotkey import _load_pipeline_prompt, DEFAULT_WORD_SYSTEM_PROMPT; result = _load_pipeline_prompt('word_assistant', 'system_prompt', DEFAULT_WORD_SYSTEM_PROMPT); print(f'Loaded: {len(result)} chars, matches default: {result == DEFAULT_WORD_SYSTEM_PROMPT}')"`

Expected: `Loaded: <N> chars, matches default: True` (since no custom edit has been made)

- [ ] **Step 5: Verify legal_research getters work**

Run: `python -c "from icharlotte_core.legal_research.prompts import get_query_planning_prompt, get_citation_instruction; print(f'query_planning: {len(get_query_planning_prompt())} chars'); print(f'citation: {len(get_citation_instruction())} chars')"`

Expected: Both show character counts > 100.

- [ ] **Step 6: Commit (if any test fixes were needed)**

```bash
git add -A
git commit -m "fix: test fixes for prompt workbench integration"
```

Only commit this step if fixes were needed. Skip if all tests passed on first run.
