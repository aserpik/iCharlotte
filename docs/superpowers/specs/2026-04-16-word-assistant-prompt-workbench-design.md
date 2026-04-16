# Word AI Assistant — Prompt Engineering Workbench Integration

**Date:** 2026-04-16
**Status:** Approved

## Goal

Make every prompt in the Word AI assistant pipeline (system prompts, inline instruction blocks, legal research prompts, mediation brief style/formatting rules) editable from the existing Prompt Engineering Workbench in iCharlotte, with versioning, A/B testing, and rollback support.

## Prompt Registry Structure

### Agent: `word_assistant` (7 passes)

| Pass | Current Source | Description |
|------|---------------|-------------|
| `system_prompt` | `DEFAULT_WORD_SYSTEM_PROMPT` constant (word_hotkey.py:170) | Main system prompt: voice/tone, attachment, contextual writing, format, continuation, elaboration rules |
| `redline_system_prompt` | `DEFAULT_WORD_REDLINE_SYSTEM_PROMPT` constant (word_hotkey.py:214) | Redline/Track Changes system prompt |
| `email_system_prompt` | `EMAIL_SYSTEM_PROMPT` constant (word_hotkey.py:225) | Outlook email system prompt |
| `redline_prefix` | Inline string (word_hotkey.py:3981) | 5-rule prefix prepended to user prompt when redline mode is active |
| `placeholder_instructions` | CRITICAL INSTRUCTIONS block in placeholder/blank branch (word_hotkey.py:~4061) | Instructions when the selection is a blank/placeholder and "Use all text" is checked |
| `cursor_instructions` | CRITICAL INSTRUCTIONS block in cursor-only branch (word_hotkey.py:~4122) | Instructions when inserting at cursor position with no selection |
| `selection_instructions` | Inline instruction at normal-selection branch (word_hotkey.py:~4082) | Instructions when processing selected text with full document context |

### Agent: `legal_research` (7 passes)

| Pass | Current Source | Description |
|------|---------------|-------------|
| `query_planning` | `QUERY_PLANNING_PROMPT` (legal_research/prompts.py:7) | Structured JSON query generation from case analysis |
| `query_extraction` | `QUERY_EXTRACTION_PROMPT` (legal_research/prompts.py:50) | Extract 3-7 research queries from litigation prompt |
| `synthesis` | `SYNTHESIS_PROMPT` (legal_research/prompts.py:78) | Synthesize authorities into research memo |
| `verification` | `VERIFICATION_PROMPT` (legal_research/prompts.py:103) | Citation verification with JSON output |
| `relevance_ranking` | `RELEVANCE_RANKING_PROMPT` (legal_research/prompts.py:149) | Rank and select most relevant cases |
| `research_framing` | `RESEARCH_FRAMING_INSTRUCTION` (legal_research/prompts.py:175) | Citation requirements injected when research results are available |
| `citation_instruction` | `CITATION_INSTRUCTION` (legal_research/prompts.py:224) | Strict citation rules (malpractice-level) |

### Agent: `mediation_brief` (2 new passes, existing agent)

| Pass | Current Source | Description |
|------|---------------|-------------|
| `style_guide` | `MediationBriefGenerator.STYLE_GUIDE` class constant (mediation_brief.py:158) | Voice/tone rules for defense litigation writing |
| `formatting_rules` | `MediationBriefGenerator.FORMATTING_RULES` class constant (mediation_brief.py:172) | Structural formatting: subsection markers, depo quote format, transcript-only restriction |

## Design Decisions

### Editability: Static Instructions Only (Option A)

The f-string prompt templates in `_do_execute` have two parts:
1. **Structural scaffolding** — section headers (`=== FULL DOCUMENT ===`), variable interpolation (`{all_text}`, `{context_before}`), and the overall template layout
2. **Instruction text** — the CRITICAL INSTRUCTIONS blocks, rules, and behavioral directives

Only the instruction text (part 2) is editable via the Workbench. The structural scaffolding stays in code. This prevents broken templates from typos in placeholder tokens while giving full control over the language the LLM sees.

### Organization (Option A)

Prompts are grouped under logical agent names (`word_assistant`, `legal_research`, `mediation_brief`) with descriptive pass names. They appear alongside existing agents in the Workbench dropdown.

### Loading & Fallback

Every prompt consumer uses the pattern:

```python
from icharlotte_core.prompt_manager import get_prompt

text = get_prompt("word_assistant", "system_prompt") or DEFAULT_WORD_SYSTEM_PROMPT
```

- If the Workbench has a saved/current version → use it
- If not → fall back to the hardcoded default constant
- Hardcoded defaults never change — "Reset to Default" is always available
- Edits take effect on the next LLM call (no restart required)

### Seeding

A `seed_pipeline_prompts()` function in `prompt_manager.py`:
- Runs once on first Workbench open (or explicitly via migration)
- For each prompt: if no version exists in PromptManager, writes the hardcoded default as `v1` and sets it as current
- Idempotent — safe to call multiple times

### Word Popup System Prompt Tab

The existing "System Prompt" tab in the Word popup (`word_hotkey.py`) currently uses its own JSON file for persistence. It will be rewired to read/write through PromptManager for `word_assistant:system_prompt`. This means edits made in the popup and edits made in the Workbench stay in sync — single source of truth.

## Files Changed

### `icharlotte_core/prompt_manager.py`
- Add `seed_pipeline_prompts()` function that seeds all 16 prompts as v1 if not already present
- Add `word_assistant`, `legal_research` to `_ensure_directory_structure()` agent directories
- Import default constants from each source module (word_hotkey, legal_research.prompts, mediation_brief)

### `icharlotte_core/word_hotkey.py`
- Extract inline instruction blocks into named default constants:
  - `DEFAULT_PLACEHOLDER_INSTRUCTIONS`
  - `DEFAULT_CURSOR_INSTRUCTIONS`
  - `DEFAULT_SELECTION_INSTRUCTIONS`
  - `DEFAULT_REDLINE_PREFIX`
- Each usage site loads from PromptManager with fallback to the default constant
- Word popup "System Prompt" tab reads/writes via PromptManager instead of standalone JSON
- Keep `DEFAULT_WORD_SYSTEM_PROMPT`, `DEFAULT_WORD_REDLINE_SYSTEM_PROMPT`, `EMAIL_SYSTEM_PROMPT` as fallback constants

### `icharlotte_core/legal_research/prompts.py`
- Add getter functions for each prompt that check PromptManager first:
  - `get_query_planning_prompt()` → `get_prompt("legal_research", "query_planning") or QUERY_PLANNING_PROMPT`
  - Same pattern for all 7 prompts
- Callers switch from `QUERY_PLANNING_PROMPT` to `get_query_planning_prompt()`

### `icharlotte_core/mediation_brief.py`
- `STYLE_GUIDE` and `FORMATTING_RULES` properties load from PromptManager with fallback to class constants
- Rename originals to `_DEFAULT_STYLE_GUIDE` and `_DEFAULT_FORMATTING_RULES` to indicate they're fallbacks

### `icharlotte_core/ui/dialogs.py`
- Add `word_assistant`, `legal_research` to predefined agents in `_populate_agents()`
- Add entries to `WORKBENCH_TO_AGENT_ID` map
- Call `seed_pipeline_prompts()` in `_migrate_if_needed()` alongside existing migration

## What Is NOT Changing

- Word/Outlook dropdown prompt templates (Improve Writing, Fix Grammar, etc.) — already editable via the popup's built-in save/delete UI
- Mediation brief per-section prompts — already managed by PromptManager
- The Workbench UI itself (Editor, LLM Assistant, A/B Testing, Version History, Dashboard tabs) — all existing functionality works unchanged with the new agents
