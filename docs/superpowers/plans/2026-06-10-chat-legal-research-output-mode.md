# Chat Legal Research Output Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a selectable Chat Legal Research output mode so users can choose the current Quick Answer behavior or a polished Research Memo format without weakening verified-authority citation guardrails.

**Architecture:** Store the output mode on `ChatResearchSettings`, which already flows through the UI, worker, research service, packet, and final prompt builder. Keep retrieval and quote verification unchanged; only the final prompt and UI settings change. Research Memo mode adds presentation instructions and role hints inside the verified authority block, while Quick Answer remains the default.

**Tech Stack:** Python 3, PySide6/QSettings/QActionGroup, pytest, existing iCharlotte chat legal-research service.

---

## Files

- Modify: `icharlotte_core/chat/legal_research.py`
  - Add `ChatResearchOutputMode`.
  - Add `output_mode` to `ChatResearchSettings`.
  - Preserve output mode through normalization.
  - Add Research Memo prompt instructions.
  - Remove the shared instruction that forces the model to generate a full `Research Basis` section.
  - Add deterministic presentation role hints for memo mode only.
- Modify: `icharlotte_core/chat/__init__.py`
  - Re-export `ChatResearchOutputMode`.
- Modify: `icharlotte_core/ui/tabs.py`
  - Import the new enum.
  - Persist and restore output mode with `QSettings`.
  - Add output mode actions to the legal research sources menu.
  - Include output mode in debug details.
- Modify: `tests/test_chat/test_legal_research_service.py`
  - Cover enum defaults, prompt differences, and memo role hints.
- Modify: `tests/test_chat/test_legal_research_ui.py`
  - Cover persistence, selected settings, async worker payload, and debug metadata.

---

### Task 1: Domain Model and Prompt Guardrails

**Files:**
- Modify: `icharlotte_core/chat/legal_research.py`
- Modify: `icharlotte_core/chat/__init__.py`
- Test: `tests/test_chat/test_legal_research_service.py`

- [ ] **Step 1: Write failing service tests for output-mode defaults and parsing**

Add these imports in `tests/test_chat/test_legal_research_service.py`:

```python
from icharlotte_core.chat.legal_research import (
    ChatAuthorityCandidate,
    ChatLegalResearchService,
    ChatResearchOutputMode,
    ChatResearchPacket,
    ChatResearchSettings,
    ChatResearchSource,
    ChatSelectedAuthority,
    CourtListenerMode,
    normalize_settings,
)
```

If the file already imports these names in a parenthesized import block, add only `ChatResearchOutputMode` to that block.

Add these tests near the existing settings tests:

```python
def test_research_output_mode_defaults_to_quick_answer():
    settings = ChatResearchSettings.default()

    assert settings.output_mode == ChatResearchOutputMode.QUICK_ANSWER


def test_research_output_mode_from_values_accepts_research_memo():
    settings = ChatResearchSettings.from_values(output_mode="research_memo")

    assert settings.output_mode == ChatResearchOutputMode.RESEARCH_MEMO


def test_unknown_research_output_mode_uses_default():
    settings = ChatResearchSettings.from_values(output_mode="unknown")

    assert settings.output_mode == ChatResearchOutputMode.QUICK_ANSWER


def test_normalize_settings_preserves_output_mode():
    settings = ChatResearchSettings(
        firm_authority=False,
        local_corpus=False,
        courtlistener_mode=CourtListenerMode.FALLBACK_CURRENT_LAW,
        output_mode=ChatResearchOutputMode.RESEARCH_MEMO,
    )

    normalized = normalize_settings(settings)

    assert normalized.courtlistener_mode == CourtListenerMode.ALWAYS_SEARCH
    assert normalized.output_mode == ChatResearchOutputMode.RESEARCH_MEMO
```

- [ ] **Step 2: Run the new settings tests and verify they fail**

Run:

```powershell
.\.venv\Scripts\python.exe -m pytest `
  tests/test_chat/test_legal_research_service.py::test_research_output_mode_defaults_to_quick_answer `
  tests/test_chat/test_legal_research_service.py::test_research_output_mode_from_values_accepts_research_memo `
  tests/test_chat/test_legal_research_service.py::test_unknown_research_output_mode_uses_default `
  tests/test_chat/test_legal_research_service.py::test_normalize_settings_preserves_output_mode `
  -q
```

Expected: FAIL because `ChatResearchOutputMode` and `output_mode` do not exist yet.

- [ ] **Step 3: Implement output-mode enum and settings plumbing**

In `icharlotte_core/chat/legal_research.py`, add this enum after `CourtListenerMode`:

```python
class ChatResearchOutputMode(str, Enum):
    QUICK_ANSWER = "quick_answer"
    RESEARCH_MEMO = "research_memo"
```

Update `ChatResearchSettings`:

```python
@dataclass(frozen=True)
class ChatResearchSettings:
    firm_authority: bool = True
    local_corpus: bool = True
    courtlistener_mode: CourtListenerMode = CourtListenerMode.FALLBACK_CURRENT_LAW
    output_mode: ChatResearchOutputMode = ChatResearchOutputMode.QUICK_ANSWER
```

Update `from_values` to accept and parse `output_mode`:

```python
    @classmethod
    def from_values(
        cls,
        *,
        firm_authority: Any = True,
        local_corpus: Any = True,
        courtlistener_mode: Any = CourtListenerMode.FALLBACK_CURRENT_LAW,
        output_mode: Any = ChatResearchOutputMode.QUICK_ANSWER,
    ) -> "ChatResearchSettings":
        if isinstance(courtlistener_mode, CourtListenerMode):
            mode = courtlistener_mode
        else:
            try:
                mode = CourtListenerMode(str(courtlistener_mode))
            except ValueError:
                mode = CourtListenerMode.FALLBACK_CURRENT_LAW
        if isinstance(output_mode, ChatResearchOutputMode):
            selected_output_mode = output_mode
        else:
            try:
                selected_output_mode = ChatResearchOutputMode(str(output_mode))
            except ValueError:
                selected_output_mode = ChatResearchOutputMode.QUICK_ANSWER
        return normalize_settings(
            cls(
                firm_authority=_bool_value(firm_authority, True),
                local_corpus=_bool_value(local_corpus, True),
                courtlistener_mode=mode,
                output_mode=selected_output_mode,
            )
        )
```

Update `normalize_settings` so the fallback-to-always conversion preserves output mode:

```python
        return ChatResearchSettings(
            firm_authority=False,
            local_corpus=False,
            courtlistener_mode=CourtListenerMode.ALWAYS_SEARCH,
            output_mode=settings.output_mode,
        )
```

Update `icharlotte_core/chat/__init__.py`:

```python
from .legal_research import (
    ChatResearchError,
    ChatResearchOutputMode,
    ChatResearchPacket,
    ChatResearchSettings,
    CourtListenerMode,
)
```

and add `'ChatResearchOutputMode'` to `__all__`.

- [ ] **Step 4: Run the settings tests and verify they pass**

Run the same command from Step 2.

Expected: PASS.

- [ ] **Step 5: Write failing service tests for Quick Answer vs Research Memo prompts**

In `tests/test_chat/test_legal_research_service.py`, update the existing `test_build_augmented_prompt_includes_research_basis_and_authorities` so it no longer expects the shared prompt to force a `Research Basis` heading. The test should assert the appended/audit helper still contains the basis, while the main prompt does not force the model to write that section:

```python
def test_quick_answer_prompt_uses_guardrails_without_memo_format():
    packet = ChatResearchPacket(
        query="duty rule",
        settings=ChatResearchSettings.default(),
        propositions=["duty rule"],
        selected_authorities=[
            ChatSelectedAuthority(
                id="cap:1",
                proposition="duty rule",
                case_name="Duty v. Care",
                citation="30 Cal. 4th 43",
                year="2020",
                reason="It states the rule.",
                supports="Duty controls negligence.",
                quote="The duty rule controls the negligence analysis.",
                sources=[ChatResearchSource(kind="local_corpus", label="Local California corpus")],
            )
        ],
    )

    prompt = packet.build_augmented_system_prompt("Base prompt.")
    basis_html = "\n".join(packet.format_research_basis_html())

    assert "Base prompt." in prompt
    assert "cite only authorities" in prompt.lower()
    assert "Duty v. Care" in prompt
    assert "Research Memo" not in prompt
    assert "Best Supporting Cases" not in prompt
    assert "Include a concise section titled \"Research Basis\"" not in prompt
    assert "Legal Research Basis" in basis_html
```

Add this Research Memo prompt test:

```python
def test_research_memo_prompt_adds_memo_format_and_role_hints():
    packet = ChatResearchPacket(
        query="duty rule",
        settings=ChatResearchSettings(output_mode=ChatResearchOutputMode.RESEARCH_MEMO),
        propositions=["duty rule"],
        selected_authorities=[
            ChatSelectedAuthority(
                id="cap:1",
                proposition="duty rule",
                case_name="Seminal v. Rule",
                citation="45 Cal. 2d 265",
                year="1955",
                court="California Supreme Court",
                reason="It states the foundational rule.",
                supports="An occupant of land may recover annoyance and discomfort damages.",
                quote="an occupant of land may recover damages for annoyance and discomfort",
                sources=[ChatResearchSource(kind="local_corpus", label="Local California corpus")],
            )
        ],
    )

    prompt = packet.build_augmented_system_prompt("Base prompt.")

    assert "Research Memo mode is enabled" in prompt
    assert "Summary" in prompt
    assert "Governing Rule" in prompt
    assert "Best Supporting Cases" in prompt
    assert "Limitations / Adverse Authority" in prompt
    assert "Suggested Argument Framing" in prompt
    assert "Presentation role: foundational" in prompt
    assert "Do not add citations from memory" in prompt
```

- [ ] **Step 6: Run the prompt tests and verify they fail**

Run:

```powershell
.\.venv\Scripts\python.exe -m pytest `
  tests/test_chat/test_legal_research_service.py::test_quick_answer_prompt_uses_guardrails_without_memo_format `
  tests/test_chat/test_legal_research_service.py::test_research_memo_prompt_adds_memo_format_and_role_hints `
  -q
```

Expected: FAIL because the old shared prompt still requires `Research Basis`, and memo instructions/role hints do not exist.

- [ ] **Step 7: Implement prompt split and role hints**

In `icharlotte_core/chat/legal_research.py`, replace `RESEARCH_PROMPT_INSTRUCTION` with:

```python
RESEARCH_PROMPT_INSTRUCTION = """LEGAL RESEARCH MODE IS ENABLED.

You must cite only authorities in [CHAT LEGAL RESEARCH AUTHORITY].
Do not invent, recall, or add citations from memory.
If the selected authorities do not support a requested proposition, say that the selected sources did not provide support.
Write a direct answer to the user's question using the verified authorities below.
"""
```

Add this constant after it:

```python
RESEARCH_MEMO_PROMPT_INSTRUCTION = """Research Memo mode is enabled.

Organize the main answer as polished attorney work product. Use this structure unless the user's prompt clearly requests a different structure:
1. Summary
2. Governing Rule
3. Best Supporting Cases
4. Limitations / Adverse Authority
5. Suggested Argument Framing

Use the Presentation role hints in the authority block to organize the discussion. Do not repeat the full Legal Research Basis appendix inside the main answer. The application will append the audit trail separately.
"""
```

Add these helpers near `_normalize_ws`:

```python
def _is_california_supreme_court(authority: ChatSelectedAuthority) -> bool:
    court = (authority.court or "").lower()
    citation = authority.citation or ""
    if "supreme" in court:
        return True
    return bool(re.search(r"\bCal\.\s*(?:2d|3d|4th|5th)\b", citation)) and "Cal.App" not in citation


def _authority_presentation_role(authority: ChatSelectedAuthority) -> str:
    caveat = (authority.caveat or "").lower()
    combined = " ".join(
        [
            authority.reason or "",
            authority.supports or "",
            authority.proposition or "",
            caveat,
        ]
    ).lower()
    if any(term in combined for term in ("adverse", "contrary", "distinguish", "cuts against")):
        return "adverse"
    if caveat:
        return "limiting"
    if _is_california_supreme_court(authority):
        return "foundational"
    if any(term in combined for term in ("directly", "squarely", "direct support")):
        return "direct"
    return "background"
```

Update `format_authority_block()` so role hints are included only in Research Memo mode:

```python
                if self.settings.output_mode == ChatResearchOutputMode.RESEARCH_MEMO:
                    lines.append(
                        f"  Presentation role: {_authority_presentation_role(authority)}"
                    )
```

Place that immediately after the `Source:` line.

Update `build_augmented_system_prompt()`:

```python
    def build_augmented_system_prompt(self, base_system_prompt: str) -> str:
        parts = [
            base_system_prompt,
            RESEARCH_PROMPT_INSTRUCTION,
        ]
        if self.settings.output_mode == ChatResearchOutputMode.RESEARCH_MEMO:
            parts.append(RESEARCH_MEMO_PROMPT_INSTRUCTION)
        parts.append(self.format_authority_block())
        return "\n\n".join(parts)
```

- [ ] **Step 8: Run prompt tests and service settings tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m pytest `
  tests/test_chat/test_legal_research_service.py::test_research_output_mode_defaults_to_quick_answer `
  tests/test_chat/test_legal_research_service.py::test_research_output_mode_from_values_accepts_research_memo `
  tests/test_chat/test_legal_research_service.py::test_unknown_research_output_mode_uses_default `
  tests/test_chat/test_legal_research_service.py::test_normalize_settings_preserves_output_mode `
  tests/test_chat/test_legal_research_service.py::test_quick_answer_prompt_uses_guardrails_without_memo_format `
  tests/test_chat/test_legal_research_service.py::test_research_memo_prompt_adds_memo_format_and_role_hints `
  -q
```

Expected: PASS.

- [ ] **Step 9: Commit Task 1**

Run:

```powershell
git diff -- icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py tests/test_chat/test_legal_research_service.py
git add icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py tests/test_chat/test_legal_research_service.py
git diff --cached --check
git commit -m "Add chat legal research output modes"
```

Before committing, verify the cached diff contains only Task 1 changes and not unrelated pre-existing worktree edits.

---

### Task 2: Chat UI Selector and Persistence

**Files:**
- Modify: `icharlotte_core/ui/tabs.py`
- Test: `tests/test_chat/test_legal_research_ui.py`

- [ ] **Step 1: Write failing UI tests for output mode persistence**

In `tests/test_chat/test_legal_research_ui.py`, update the import:

```python
from icharlotte_core.chat.legal_research import ChatResearchOutputMode, CourtListenerMode
```

Update `_clear_chat_research_settings()` to remove:

```python
"chat_tab/legal_research_output_mode",
```

Add this test near the existing source persistence tests:

```python
def test_chat_research_output_mode_persists(qtbot, monkeypatch):
    _app()
    _clear_chat_research_settings()

    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.research_memo_action.setChecked(True)
    tab._on_research_source_changed()

    tab2 = _make_chat_tab(qtbot, monkeypatch)
    settings = tab2._current_chat_research_settings()

    assert settings.output_mode == ChatResearchOutputMode.RESEARCH_MEMO
    assert tab2.research_memo_action.isChecked() is True
    assert "Memo" in tab2.research_sources_btn.text()
```

- [ ] **Step 2: Run the new UI persistence test and verify it fails**

Run:

```powershell
.\.venv\Scripts\python.exe -m pytest `
  tests/test_chat/test_legal_research_ui.py::test_chat_research_output_mode_persists `
  -q
```

Expected: FAIL because `research_memo_action` does not exist.

- [ ] **Step 3: Implement UI output mode actions**

In `icharlotte_core/ui/tabs.py`, update imports from `..chat`:

```python
from ..chat import (
    ChatPersistence,
    TokenCounter,
    Message,
    Conversation,
    BUILTIN_PROMPTS,
    TRANSCRIBE_PROMPT,
    ChatResearchError,
    ChatResearchOutputMode,
    ChatResearchSettings,
    CourtListenerMode,
)
```

Update `_load_chat_research_settings()`:

```python
            output_mode=settings.value(
                "chat_tab/legal_research_output_mode",
                ChatResearchOutputMode.QUICK_ANSWER.value,
            ),
```

Update `_save_chat_research_settings()`:

```python
        settings.setValue(
            "chat_tab/legal_research_output_mode",
            research_settings.output_mode.value,
        )
```

In `_build_research_sources_menu()`, after the CourtListener action setup and checked-state logic, add a separator and output mode group:

```python
        menu.addSeparator()
        output_group = QActionGroup(menu)
        output_group.setExclusive(True)
        self.quick_answer_action = QAction("Output: Quick Answer", menu)
        self.research_memo_action = QAction("Output: Research Memo", menu)
        for action in (self.quick_answer_action, self.research_memo_action):
            action.setCheckable(True)
            output_group.addAction(action)
            menu.addAction(action)

        if current.output_mode == ChatResearchOutputMode.RESEARCH_MEMO:
            self.research_memo_action.setChecked(True)
        else:
            self.quick_answer_action.setChecked(True)
```

Update the action connection tuple to include both output actions:

```python
            self.quick_answer_action,
            self.research_memo_action,
```

Update `_current_chat_research_settings()`:

```python
        output_mode = (
            ChatResearchOutputMode.RESEARCH_MEMO
            if self.research_memo_action.isChecked()
            else ChatResearchOutputMode.QUICK_ANSWER
        )
        return ChatResearchSettings.from_values(
            firm_authority=self.firm_authority_action.isChecked(),
            local_corpus=self.local_corpus_action.isChecked(),
            courtlistener_mode=mode.value,
            output_mode=output_mode.value,
        )
```

Update `_on_research_source_changed()` after the CourtListener checked-state normalization:

```python
        if current.output_mode == ChatResearchOutputMode.RESEARCH_MEMO:
            self.research_memo_action.setChecked(True)
        else:
            self.quick_answer_action.setChecked(True)
```

Update `_refresh_research_sources_label()`:

```python
        output = "Memo" if current.output_mode == ChatResearchOutputMode.RESEARCH_MEMO else "Quick"
        self.research_sources_btn.setText(
            "Sources: " + " + ".join(parts) + f" | {output}"
        )
```

- [ ] **Step 4: Run the UI persistence test and verify it passes**

Run the command from Step 2.

Expected: PASS.

- [ ] **Step 5: Write failing UI tests for settings propagation and debug metadata**

Update `test_chat_research_source_defaults` expected button text:

```python
assert tab.research_sources_btn.text() == "Sources: Firm + Local + CL Fallback | Quick"
```

Update `test_run_chat_legal_research_passes_selected_settings` by setting memo mode before `_on_research_source_changed()`:

```python
tab.research_memo_action.setChecked(True)
```

and add:

```python
assert captured["settings"].output_mode == ChatResearchOutputMode.RESEARCH_MEMO
```

Update `test_send_message_dispatches_legal_research_in_background` by setting memo mode before `send_message()`:

```python
tab.research_memo_action.setChecked(True)
tab._on_research_source_changed()
```

and add:

```python
assert captured["research_settings"].output_mode == ChatResearchOutputMode.RESEARCH_MEMO
```

Add this test near `test_run_chat_legal_research_records_task_debug_events`:

```python
def test_run_chat_legal_research_records_output_mode_in_debug_details(
    qtbot,
    monkeypatch,
    tmp_path,
):
    _app()
    _clear_chat_research_settings()
    from icharlotte_core import task_debug
    from icharlotte_core.ui import tabs

    class FakeService:
        @classmethod
        def from_environment(cls, *, llm_callback):
            return cls()

        def research(self, *, user_text, context_text, settings, status_callback, debug_callback=None):
            return SimpleNamespace(
                selected_authorities=[],
                get_known_case_names=lambda: [],
                build_augmented_system_prompt=lambda base: base + "\nAUGMENTED",
                format_research_basis_html=lambda: ["<b>Legal Research Basis</b>"],
            )

    monkeypatch.setattr(tabs, "ChatLegalResearchService", FakeService)
    task_debug.reset_for_tests(trace_dir=tmp_path / "traces", max_events=20)
    tab = _make_chat_tab(qtbot, monkeypatch)
    tab.research_memo_action.setChecked(True)
    tab._on_research_source_changed()

    packet = tab._run_chat_legal_research("research this", "context text")

    events = task_debug.get_events()
    assert packet is not None
    assert events[0].details["research_settings"]["output_mode"] == "research_memo"
```

- [ ] **Step 6: Run the affected UI tests and verify they fail where implementation is incomplete**

Run:

```powershell
.\.venv\Scripts\python.exe -m pytest `
  tests/test_chat/test_legal_research_ui.py::test_chat_research_source_defaults `
  tests/test_chat/test_legal_research_ui.py::test_run_chat_legal_research_passes_selected_settings `
  tests/test_chat/test_legal_research_ui.py::test_send_message_dispatches_legal_research_in_background `
  tests/test_chat/test_legal_research_ui.py::test_run_chat_legal_research_records_output_mode_in_debug_details `
  -q
```

Expected: FAIL until debug metadata and propagation are fully updated.

- [ ] **Step 7: Add output mode to debug metadata**

In `_start_chat_legal_research()`, add `output_mode` to the `research_settings` detail dict:

```python
                    "output_mode": research_settings.output_mode.value,
```

In `_run_chat_legal_research()`, add the same key in its `task_debug.start_run()` detail dict:

```python
                    "output_mode": research_settings.output_mode.value,
```

- [ ] **Step 8: Run the affected UI tests**

Run the command from Step 6.

Expected: PASS.

- [ ] **Step 9: Commit Task 2**

Run:

```powershell
git diff -- icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_ui.py
git add icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_ui.py
git diff --cached --check
git commit -m "Add chat legal research output selector"
```

Before committing, verify the cached diff contains only Task 2 changes and not unrelated pre-existing worktree edits.

---

### Task 3: Focused Regression Verification

**Files:**
- Verify: `icharlotte_core/chat/legal_research.py`
- Verify: `icharlotte_core/ui/tabs.py`
- Verify: `tests/test_chat/test_legal_research_service.py`
- Verify: `tests/test_chat/test_legal_research_ui.py`

- [ ] **Step 1: Run focused chat legal-research tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m pytest `
  tests/test_chat/test_legal_research_service.py `
  tests/test_chat/test_legal_research_ui.py `
  -q
```

Expected: PASS. If unrelated pre-existing failures appear, rerun the narrower tests from Tasks 1 and 2 and record the failing test names and error text.

- [ ] **Step 2: Compile changed Python files**

Run:

```powershell
.\.venv\Scripts\python.exe -m py_compile `
  icharlotte_core/chat/legal_research.py `
  icharlotte_core/chat/__init__.py `
  icharlotte_core/ui/tabs.py
```

Expected: command exits with code 0 and no output.

- [ ] **Step 3: Inspect final changed subset**

Run:

```powershell
git status --short
git diff -- icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py
```

Expected: only intentional modifications for the output mode feature appear in this subset. Other dirty files may remain in the worktree and must not be staged unless the user asks.

- [ ] **Step 4: Optional final commit if Task 1 and Task 2 were not already committed**

If Tasks 1 and 2 were committed separately, skip this step.

If they were not committed, run:

```powershell
git add icharlotte_core/chat/legal_research.py icharlotte_core/chat/__init__.py icharlotte_core/ui/tabs.py tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py
git diff --cached --check
git commit -m "Add selectable chat legal research output mode"
```

Expected: commit succeeds and includes only the output-mode implementation files.
