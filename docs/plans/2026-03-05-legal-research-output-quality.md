# Legal Research Agent Output Quality Improvements

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Improve case law citation depth and legal analysis quality in research-augmented LLM output.

**Architecture:** Four targeted changes: (1) inject the synthesis memo into the final prompt instead of discarding it, (2) increase opinion snippet length so the synthesis LLM has enough text to extract holdings, (3) extract a focused research query from the full prompt, (4) strengthen citation instructions to demand case-law-driven analysis with parentheticals.

**Tech Stack:** Python, existing `icharlotte_core/legal_research/` module, no new dependencies.

---

### Task 1: Increase opinion snippet length for synthesis LLM

**Files:**
- Modify: `icharlotte_core/legal_research/engine.py:27`

**Step 1: Update the snippet length constant**

In `engine.py`, change line 27:

```python
# Was: _OPINION_SNIPPET_LENGTH = 1500
_OPINION_SNIPPET_LENGTH = 5000
```

This gives the synthesis LLM (Phase 5) enough opinion text (~800 words) to extract actual holdings and rule statements. The authority block limits in `models.py` stay as-is since we'll inject the memo instead (Task 3).

**Step 2: Verify no other code depends on the old value**

Run: `grep -rn "_OPINION_SNIPPET_LENGTH\|1500" icharlotte_core/legal_research/`
Expected: Only `engine.py:27` references this constant.

**Step 3: Commit**

```bash
git add icharlotte_core/legal_research/engine.py
git commit -m "feat(legal-research): increase opinion snippet to 5000 chars for better holdings extraction"
```

---

### Task 2: Add query extraction prompt and strengthened citation instruction

**Files:**
- Modify: `icharlotte_core/legal_research/prompts.py`

**Step 1: Add QUERY_EXTRACTION_PROMPT after the existing QUERY_PLANNING_PROMPT**

Add this new prompt constant after `QUERY_PLANNING_PROMPT` (after line 43):

```python
QUERY_EXTRACTION_PROMPT = """\
You are extracting the core legal questions from a litigation analysis prompt. \
The user has a long prompt containing case facts, parties, causes of action, and \
instructions. Your job is to distill this into 1-3 focused legal questions that \
a legal research database should answer.

Output ONLY the legal questions, one per line. Each question should be specific \
enough to find relevant case law. Do not include case facts, party names, or \
instructions.

Examples:
- "What duty does a contractor owe to a property owner for work outside the contracted scope?"
- "Can a property owner's prior knowledge of a dangerous condition sever the causal chain for a contractor's negligence?"
- "What are the elements of breach of oral contract under California law?"

Keep total output under 500 characters."""
```

**Step 2: Replace CITATION_INSTRUCTION with a strengthened version**

Replace the existing `CITATION_INSTRUCTION` (lines 130-135) with:

```python
CITATION_INSTRUCTION = (
    "You MUST cite specific case law to support every legal proposition. For each "
    "major legal point:\n"
    "1. State the rule with a case citation in California format: Case Name (Year) "
    "Volume Reporter Page\n"
    "2. Include a parenthetical explanation of the holding: (holding that...)\n"
    "3. Apply the holding to the facts of this case\n"
    "4. Address any adverse authority from the provided sources and distinguish it\n\n"
    "You may ONLY cite cases and statutes from the [LEGAL RESEARCH MEMO] and "
    "[LEGAL AUTHORITY] sections below. Do NOT fabricate or hallucinate any "
    "citations. If you cannot find sufficient authority, state that expressly.\n\n"
    "Cite statutes as: Code Name, section symbol Section (e.g., Civ. Code, section 1714).\n"
    "Note any negative treatment (overruled, disapproved) when citing a negatively "
    "treated case."
)
```

**Step 3: Update `build_augmented_system_prompt` to accept optional memo**

Replace the function (lines 138-154) with:

```python
def build_augmented_system_prompt(
    base_system_prompt: str,
    authority_block: str,
    research_memo: str = "",
) -> str:
    """Combine a base system prompt with citation instructions, memo, and authority.

    Args:
        base_system_prompt: The base system prompt for the task.
        authority_block: Pre-formatted block of legal authorities.
        research_memo: Optional synthesis memo from the research engine.

    Returns:
        Combined prompt with base, citation instruction, memo, and authority block.
    """
    parts = [base_system_prompt, "", CITATION_INSTRUCTION]

    if research_memo:
        parts.append("")
        parts.append("[LEGAL RESEARCH MEMO]")
        parts.append(research_memo)
        parts.append("[/LEGAL RESEARCH MEMO]")

    parts.append("")
    parts.append(f"[LEGAL AUTHORITY]\n{authority_block}")

    return "\n".join(parts)
```

**Step 4: Run existing tests**

Run: `python -m pytest tests/test_legal_research/ -v`
Expected: All tests pass (may need minor updates if tests assert on `build_augmented_system_prompt` signature).

**Step 5: Fix any test failures from signature change**

If tests call `build_augmented_system_prompt(base, authority)` positionally, they will still work since `research_memo` has a default value. Check:

Run: `grep -rn "build_augmented_system_prompt" tests/`
Fix any issues found.

**Step 6: Commit**

```bash
git add icharlotte_core/legal_research/prompts.py
git commit -m "feat(legal-research): add query extraction prompt and strengthen citation instructions"
```

---

### Task 3: Extract focused research query in Win+V popup

**Files:**
- Modify: `icharlotte_core/word_hotkey.py:2415-2465`

**Step 1: Add query extraction before engine.research()**

Replace the research block (lines 2415-2465) with:

```python
            # Legal Research: if checkbox is checked, run research and augment prompt
            if self.legal_research_checkbox.isChecked():
                engine = self._get_legal_research_engine()
                if engine:
                    self.status_label.setText("Researching legal authority...")
                    QApplication.processEvents()

                    from icharlotte_core.legal_research.prompts import (
                        build_augmented_system_prompt,
                        QUERY_EXTRACTION_PROMPT,
                        CITATION_INSTRUCTION,
                    )

                    def _llm_for_research(system_prompt, user_prompt):
                        """Synchronous LLM call for research sub-steps."""
                        from icharlotte_core.llm import LLMHandler
                        provider, model_id = self._get_selected_model()
                        settings = {
                            'temperature': 0.3,
                            'top_p': 0.95,
                            'max_tokens': -1,
                            'stream': False,
                            'thinking_level': 'None',
                        }
                        return LLMHandler.generate(
                            provider=provider, model=model_id,
                            system_prompt=system_prompt,
                            user_prompt=user_prompt,
                            file_contents="", settings=settings,
                        )

                    try:
                        # Extract focused legal questions from the full prompt
                        self.status_label.setText("Extracting legal questions...")
                        QApplication.processEvents()
                        research_query = _llm_for_research(
                            QUERY_EXTRACTION_PROMPT, full_prompt[:8000]
                        )
                        if not research_query or len(research_query.strip()) < 20:
                            research_query = full_prompt[:2000]

                        research_result = engine.research(
                            query=research_query,
                            llm_callback=_llm_for_research,
                            status_callback=lambda msg: (
                                self.status_label.setText(msg),
                                QApplication.processEvents(),
                            ),
                        )
                        authority_block = research_result.format_authority_block()
                        self._pending_research_result = research_result
                    except Exception as e:
                        print(f"[LegalResearch] Engine error: {e}")
                        self._pending_research_result = None
                        authority_block = ""
                        research_result = None

                    if authority_block and research_result:
                        # Inject both the research memo and authority block
                        memo = research_result.memo or ""
                        if memo:
                            full_prompt = (
                                f"{full_prompt}\n\n"
                                f"[LEGAL RESEARCH MEMO]\n{memo}\n[/LEGAL RESEARCH MEMO]\n\n"
                                f"{authority_block}\n\n"
                                f"{CITATION_INSTRUCTION}"
                            )
                        else:
                            full_prompt = (
                                f"{full_prompt}\n\n"
                                f"{authority_block}\n\n"
                                f"{CITATION_INSTRUCTION}"
                            )

                    self.status_label.setText("Generating response with legal citations...")
                    QApplication.processEvents()
                else:
                    self._pending_research_result = None
                    print("[LegalResearch] No COURTLISTENER_API_TOKEN in .env — skipping research")
            else:
                self._pending_research_result = None
```

**Step 2: Commit**

```bash
git add icharlotte_core/word_hotkey.py
git commit -m "feat(legal-research): extract focused query and inject research memo in Win+V popup"
```

---

### Task 4: Update ChatTab integration to inject research memo

**Files:**
- Modify: `icharlotte_core/ui/tabs.py:1082-1161`

**Step 1: Add query extraction to ChatTab research flow**

Replace lines 1090-1092 (the research query construction):

```python
                # Extract focused legal questions instead of sending raw prompt
                from icharlotte_core.legal_research.prompts import QUERY_EXTRACTION_PROMPT
                self.chat_history.append("<i>  Extracting legal questions...</i>")
                QApplication.processEvents()
                research_query = _llm_for_research(
                    QUERY_EXTRACTION_PROMPT,
                    (user_text + ("\n\nContext:\n" + file_content[:8000] if file_content else ""))[:8000]
                )
                if not research_query or len(research_query.strip()) < 20:
                    research_query = user_text
                    if file_content:
                        research_query += "\n\nContext:\n" + file_content[:8000]
```

**Step 2: Update the system prompt augmentation to include memo**

Replace lines 1154-1161:

```python
        # Build system prompt (augment with legal authority if research was done)
        effective_system_prompt = self.system_prompt
        if research_result:
            from icharlotte_core.legal_research.prompts import build_augmented_system_prompt
            authority = research_result.format_authority_block()
            memo = research_result.memo or ""
            effective_system_prompt = build_augmented_system_prompt(
                self.system_prompt, authority, research_memo=memo
            )
```

**Step 3: Run existing tests**

Run: `python -m pytest tests/test_legal_research/ -v`
Expected: All pass.

**Step 4: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(legal-research): extract focused query and inject memo in ChatTab"
```

---

### Task 5: Manual integration test

**Step 1: Test Win+V popup with legal research checkbox**

1. Open Word with a liability analysis document
2. Select text containing causes of action
3. Check "Perform Legal Research" checkbox
4. Choose a liability analysis prompt
5. Submit

**Step 2: Verify output quality**

Check that the response:
- [ ] Cites specific case names with full California citations
- [ ] Includes parenthetical explanations of holdings
- [ ] Applies case holdings to the facts
- [ ] References more than just statutes
- [ ] Does not contain `[UNVERIFIED]` flags (ideally)

**Step 3: Test ChatTab with legal research checkbox**

1. Open iCharlotte ChatTab
2. Check "Legal Research" checkbox
3. Ask a legal question about a case
4. Verify same quality criteria

**Step 4: Commit any fixes**

```bash
git add -u
git commit -m "fix(legal-research): integration test fixes"
```
