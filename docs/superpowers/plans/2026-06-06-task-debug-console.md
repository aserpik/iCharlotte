# Task Debug Console Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an always-capturing task debug event stream, floating Debug Console window, and granular chat legal-research diagnostics.

**Architecture:** Add a small process-wide task-debug recorder with JSONL traces and a Qt signal bridge. Add a floating PySide6 debug console subscribed to the recorder. Wire generic wizard task containers, custom wizard task tabs, and chat legal research into the recorder without changing existing concise status UI behavior.

**Tech Stack:** Python 3, PySide6, pytest, pytest-qt, SQLite-free file JSONL logging.

---

## File Structure

- Create `icharlotte_core/task_debug.py`: dataclass event model, bounded recorder, JSONL writer, run lifecycle helpers, Qt signal bridge, test reset helpers.
- Create `icharlotte_core/ui/task_debug_window.py`: floating debug console UI with filters, search, pause autoscroll, copy, clear, and open trace folder.
- Modify `iCharlotte.py`: add `View > Debug Console` action and lazy window opener. Preserve the existing uncommitted case-switching block around `load_case_by_number`.
- Modify `icharlotte_core/ui/wizard/task_tab.py`: mirror generic subprocess task status/progress/finish/fail/cancel/file-completed events to the recorder.
- Modify `icharlotte_core/ui/wizard/in_process_task_tab.py`: mirror in-process worker progress/warning/finish/fail/cancel events to the recorder.
- Modify custom wizard task pages:
  - `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
  - `icharlotte_core/ui/wizard/pages/generate_motion_page.py`
  - `icharlotte_core/ui/wizard/pages/mediation_brief_page.py`
  - `icharlotte_core/ui/wizard/pages/separate_page.py`
  - `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py`
- Modify `icharlotte_core/chat/legal_research.py`: add optional granular debug callback while keeping `status_callback` behavior.
- Modify `icharlotte_core/ui/tabs.py`: start/finish a `chat_legal_research` run and pass a debug callback into `ChatLegalResearchService`.
- Create tests:
  - `tests/test_task_debug.py`
  - `tests/test_ui_task_debug_window.py`
  - `tests/test_wizard/test_task_debug_instrumentation.py`
- Modify existing chat tests:
  - `tests/test_chat/test_legal_research_service.py`
  - `tests/test_chat/test_legal_research_ui.py`

## Task 1: Core Task Debug Recorder

**Files:**
- Create: `icharlotte_core/task_debug.py`
- Test: `tests/test_task_debug.py`

- [ ] **Step 1: Write failing tests for event serialization, redaction, buffer, run lifecycle, JSONL, and subscriber bridge**

```python
def test_start_emit_finish_records_buffer_and_jsonl(tmp_path, monkeypatch):
    import icharlotte_core.task_debug as td
    td.reset_for_tests(trace_dir=tmp_path)

    run_id = td.start_run(
        task_id="chat_legal_research",
        task_title="Chat Legal Research",
        source="ChatTab",
        details={"api_token": "secret", "file_count": 2},
    )
    td.emit_event(
        run_id=run_id,
        task_id="chat_legal_research",
        task_title="Chat Legal Research",
        phase="search",
        level="info",
        message="Searching local corpus",
        source="Local California corpus",
        details={"hits": 3},
    )
    td.finish_run(run_id, status="success", message="Research complete")

    events = td.get_events()
    assert [e.phase for e in events] == ["start", "search", "finish"]
    assert events[0].details["api_token"] == "[REDACTED]"
    trace_files = list(tmp_path.glob("*.jsonl"))
    assert len(trace_files) == 1
    text = trace_files[0].read_text(encoding="utf-8")
    assert "Searching local corpus" in text
    assert "secret" not in text

def test_recorder_coerces_non_json_details_and_bounds_buffer(tmp_path):
    import icharlotte_core.task_debug as td
    td.reset_for_tests(trace_dir=tmp_path, max_events=2)
    run_id = td.start_run(task_id="t", task_title="Task")
    td.emit_event(run_id=run_id, task_id="t", task_title="Task", phase="one", message="one", details={"bad": object()})
    td.emit_event(run_id=run_id, task_id="t", task_title="Task", phase="two", message="two")
    td.emit_event(run_id=run_id, task_id="t", task_title="Task", phase="three", message="three")
    assert [event.phase for event in td.get_events()] == ["two", "three"]
```

- [ ] **Step 2: Run tests and verify they fail because `icharlotte_core.task_debug` does not exist**

Run: `python -m pytest tests/test_task_debug.py -q`

Expected: FAIL with `ModuleNotFoundError: No module named 'icharlotte_core.task_debug'`.

- [ ] **Step 3: Implement minimal recorder**

Implement `TaskDebugEvent`, `_TaskDebugRecorder`, `start_run`, `emit_event`, `finish_run`, `get_events`, `clear_events`, `get_trace_dir`, and `reset_for_tests`. Use `collections.deque(maxlen=...)`, a `threading.RLock`, JSONL writes with `encoding="utf-8"`, and a PySide6 `QObject` bridge with `event_emitted = Signal(object)`.

- [ ] **Step 4: Run tests and verify green**

Run: `python -m pytest tests/test_task_debug.py -q`

Expected: PASS.

## Task 2: Debug Console Window

**Files:**
- Create: `icharlotte_core/ui/task_debug_window.py`
- Test: `tests/test_ui_task_debug_window.py`

- [ ] **Step 1: Write failing Qt tests for loading buffered events and filtering**

```python
def test_debug_console_loads_buffer_and_filters(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow
    td.reset_for_tests(trace_dir=tmp_path)
    run_id = td.start_run(task_id="task_a", task_title="Task A")
    td.emit_event(run_id=run_id, task_id="task_a", task_title="Task A", phase="search", message="Local hit", source="Local")
    td.emit_event(run_id=run_id, task_id="task_a", task_title="Task A", phase="verify", message="Verifier warning", level="warning", source="Verifier")

    window = TaskDebugWindow()
    qtbot.addWidget(window)
    assert window.table.rowCount() == 3

    window.level_filter.setCurrentText("warning")
    assert window.table.rowCount() == 1
    assert "Verifier warning" in window.table.item(0, window.COL_MESSAGE).text()
```

- [ ] **Step 2: Run test and verify it fails because `TaskDebugWindow` does not exist**

Run: `python -m pytest tests/test_ui_task_debug_window.py -q`

Expected: FAIL with `ModuleNotFoundError` or import error for `task_debug_window`.

- [ ] **Step 3: Implement the floating console**

Create a `QMainWindow` subclass with filter row (`QComboBox` for task/source/level, `QLineEdit` search, `QCheckBox` pause autoscroll, `QPushButton` clear/copy/open folder) and a `QTableWidget`. Subscribe to `task_debug.get_bridge().event_emitted`, load `task_debug.get_events()` on init, and rebuild visible rows on filter changes.

- [ ] **Step 4: Run test and verify green**

Run: `python -m pytest tests/test_ui_task_debug_window.py -q`

Expected: PASS.

## Task 3: Main Window View Menu Integration

**Files:**
- Modify: `iCharlotte.py`
- Test: `tests/test_wizard/test_task_debug_instrumentation.py`

- [ ] **Step 1: Write failing test for `View > Debug Console` action**

```python
def test_main_window_view_menu_has_debug_console_action(monkeypatch, qtbot):
    import iCharlotte as app_mod

    called = []
    class FakeDebugWindow:
        def __init__(self, parent=None):
            called.append(("init", parent is not None))
        def show(self):
            called.append(("show", True))
        def raise_(self):
            called.append(("raise", True))
        def activateWindow(self):
            called.append(("activate", True))

    monkeypatch.setattr("icharlotte_core.ui.task_debug_window.TaskDebugWindow", FakeDebugWindow)
    window = object.__new__(app_mod.MainWindow)
    window.view_menu = app_mod.QMenu()
    window._task_debug_window = None
    app_mod.MainWindow._add_debug_console_action(window)

    actions = [action.text() for action in window.view_menu.actions()]
    assert "Debug Console" in actions
    window.view_menu.actions()[-1].trigger()
    assert ("show", True) in called
```

- [ ] **Step 2: Run test and verify it fails because helper/action does not exist**

Run: `python -m pytest tests/test_wizard/test_task_debug_instrumentation.py::test_main_window_view_menu_has_debug_console_action -q`

Expected: FAIL with `AttributeError: type object 'MainWindow' has no attribute '_add_debug_console_action'`.

- [ ] **Step 3: Add menu helper and opener**

In `setup_view_menu`, after existing tab actions, call `self._add_debug_console_action()`. Add `_add_debug_console_action(self)` and `show_task_debug_console(self)` methods. Lazy-import `TaskDebugWindow`.

- [ ] **Step 4: Run test and verify green**

Run: `python -m pytest tests/test_wizard/test_task_debug_instrumentation.py::test_main_window_view_menu_has_debug_console_action -q`

Expected: PASS.

## Task 4: Generic Wizard Task Instrumentation

**Files:**
- Modify: `icharlotte_core/ui/wizard/task_tab.py`
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py`
- Test: `tests/test_wizard/test_task_debug_instrumentation.py`

- [ ] **Step 1: Write failing tests for generic subprocess and in-process mirroring**

```python
def test_task_tab_mirrors_status_progress_and_failure_to_debug(qtbot, tmp_path, monkeypatch):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.wizard.task_tab import TaskTab
    from icharlotte_core.ui.wizard.registry import get_task
    td.reset_for_tests(trace_dir=tmp_path)

    tab = TaskTab(get_task("summarize"), files=["/tmp/a.pdf"], case_path="/tmp/case", file_number="1000.001")
    qtbot.addWidget(tab)
    tab._debug_run_id = td.start_run(task_id="summarize", task_title="Summarize")
    tab._on_worker_status("Extracting text")
    tab._on_worker_progress(42)
    tab._on_worker_failed("boom")

    messages = [event.message for event in td.get_events()]
    assert "Extracting text" in messages
    assert any(event.phase == "progress" and event.details["percent"] == 42 for event in td.get_events())
    assert any(event.level == "error" and "boom" in event.message for event in td.get_events())
```

- [ ] **Step 2: Run test and verify it fails because `_on_worker_status` and `_on_worker_progress` do not exist**

Run: `python -m pytest tests/test_wizard/test_task_debug_instrumentation.py::test_task_tab_mirrors_status_progress_and_failure_to_debug -q`

Expected: FAIL with `AttributeError`.

- [ ] **Step 3: Implement status/progress wrappers and run lifecycle in `TaskTab`**

Add `_debug_run_id`, `_start_debug_run`, `_emit_debug`, `_finish_debug_run`, `_on_worker_status`, and `_on_worker_progress`. Replace direct `worker.status.connect(self.status_page.on_status)` and `worker.progress.connect(self.status_page.on_progress)` connections with wrappers where TaskTab owns signals. Emit file-completed, finish, failure, cancel, awaiting-input, and phase2-resume events.

- [ ] **Step 4: Add in-process task wrappers**

Add `_debug_run_id`, `_start_debug_run`, `_emit_debug`, `_finish_debug_run`, `_on_worker_progress`, and `_on_worker_warning` to `InProcessTaskTab`. Replace lambda/direct progress wiring with wrappers.

- [ ] **Step 5: Run focused tests and verify green**

Run: `python -m pytest tests/test_wizard/test_task_debug_instrumentation.py tests/test_wizard/test_task_tab.py tests/test_wizard/test_in_process_task_tab.py -q`

Expected: PASS.

## Task 5: Custom Wizard Task Instrumentation

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
- Modify: `icharlotte_core/ui/wizard/pages/generate_motion_page.py`
- Modify: `icharlotte_core/ui/wizard/pages/mediation_brief_page.py`
- Modify: `icharlotte_core/ui/wizard/pages/separate_page.py`
- Modify: `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py`
- Test: `tests/test_wizard/test_task_debug_instrumentation.py`

- [ ] **Step 1: Write failing tests for one custom string-progress task and one dual-status task**

```python
def test_mediation_brief_task_mirrors_progress_to_debug(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.wizard.pages.mediation_brief_page import MediationBriefTaskTab
    from icharlotte_core.ui.wizard.registry import get_task
    td.reset_for_tests(trace_dir=tmp_path)
    tab = MediationBriefTaskTab(get_task("mediation_brief"), case_path="/tmp/case", file_number="1000.001")
    qtbot.addWidget(tab)
    tab._debug_run_id = td.start_run(task_id="mediation_brief", task_title="Mediation Brief")
    tab._on_worker_progress("Reading source documents")
    assert any(event.message == "Reading source documents" for event in td.get_events())
```

- [ ] **Step 2: Run test and verify it fails because custom progress wrapper does not exist**

Run: `python -m pytest tests/test_wizard/test_task_debug_instrumentation.py::test_mediation_brief_task_mirrors_progress_to_debug -q`

Expected: FAIL with `AttributeError`.

- [ ] **Step 3: Add custom tab debug wrappers**

For each custom task tab, add these methods with task-specific titles and phases:

```python
def _start_debug_run(self, phase: str, *, details: dict | None = None) -> None:
    self._debug_run_id = task_debug.start_run(
        task_id=self._spec.task_id,
        task_title=self._spec.title,
        source=type(self).__name__,
        details={"file_number": self._file_number, "phase": phase, **(details or {})},
    )

def _emit_debug(self, phase: str, message: str, *, level: str = "info", details: dict | None = None) -> None:
    if not getattr(self, "_debug_run_id", ""):
        return
    task_debug.emit_event(
        run_id=self._debug_run_id,
        task_id=self._spec.task_id,
        task_title=self._spec.title,
        phase=phase,
        level=level,
        message=message,
        source=type(self).__name__,
        details=details or {},
    )

def _on_worker_progress(self, message: str) -> None:
    self.status_page.on_status(message)
    level = "warning" if str(message).upper().startswith("WARNING:") else "info"
    self._emit_debug("status", str(message), level=level)

def _finish_debug_run(self, status: str, message: str) -> None:
    if getattr(self, "_debug_run_id", ""):
        task_debug.finish_run(self._debug_run_id, status=status, message=message)
        self._debug_run_id = ""
```

Preserve existing `status_page.on_status(...)` behavior. For dual-flow pages such as Oppose Motion and Generate Motion, call `_start_debug_run("analysis")` for analysis workers and `_start_debug_run("draft")` for draft workers while keeping the same task id and title.

- [ ] **Step 4: Run focused custom wizard tests**

Run: `python -m pytest tests/test_wizard/test_oppose_motion_page.py tests/test_wizard/test_generate_motion_worker.py tests/test_wizard/test_mediation_brief_page.py tests/test_wizard/test_separate_task.py tests/test_wizard/test_case_intake_docket_page.py tests/test_wizard/test_task_debug_instrumentation.py -q`

Expected: PASS.

## Task 6: Chat Legal Research Granular Events

**Files:**
- Modify: `icharlotte_core/chat/legal_research.py`
- Modify: `icharlotte_core/ui/tabs.py`
- Modify: `tests/test_chat/test_legal_research_service.py`
- Modify: `tests/test_chat/test_legal_research_ui.py`

- [ ] **Step 1: Write failing service test for granular debug events**

```python
def test_research_emits_granular_debug_events_for_selected_sources():
    local = FakeCorpusClient(
        results=[_case(text="The duty rule controls the negligence analysis.")],
        text_by_id={"cap:1": "The duty rule controls the negligence analysis."},
    )
    events = []
    def debug(event):
        events.append(event)
    def llm(system_prompt, user_prompt):
        if "extracting focused" in system_prompt:
            return '{"propositions":["duty rule"]}'
        return '{"selections":[{"id":"cap:1","reason":"It states the rule.","supports":"Duty controls.","quote":"The duty rule controls the negligence analysis.","caveat":""}]}'

    service = ChatLegalResearchService(llm_callback=llm, local_corpus=local)
    service.research(
        user_text="research duty rule",
        context_text="",
        settings=ChatResearchSettings(firm_authority=False, local_corpus=True, courtlistener_mode=CourtListenerMode.OFF),
        debug_callback=debug,
    )

    phases = [event["phase"] for event in events]
    assert "propositions" in phases
    assert "search" in phases
    assert any(event["source"] == "Local California corpus" and event["details"]["candidate_count"] == 1 for event in events)
    assert any(event["phase"] == "select" and event["details"]["selected_count"] == 1 for event in events)
```

- [ ] **Step 2: Run test and verify it fails because `debug_callback` is not accepted**

Run: `python -m pytest tests/test_chat/test_legal_research_service.py::test_research_emits_granular_debug_events_for_selected_sources -q`

Expected: FAIL with `TypeError: research() got an unexpected keyword argument 'debug_callback'`.

- [ ] **Step 3: Add debug callback support to service**

Add `DebugCallback = Optional[Callable[[dict[str, Any]], None]]` and `_emit_debug(callback, phase, message, level="info", source="", details=None)`. Pass `debug_callback` through `research`, `collect_candidates`, `_collect_firm`, and `_collect_case_client`. Emit selected settings, proposition count, source start/finish, fallback decisions, candidate counts, warnings, selection counts, and completion.

- [ ] **Step 4: Update UI test fake service signature and add run lifecycle assertion**

In `tests/test_chat/test_legal_research_ui.py`, update fake `research(...)` signatures to accept `debug_callback=None`. Add a test that resets `task_debug`, calls `_run_chat_legal_research`, and asserts `chat_legal_research` start/status/finish events are recorded.

- [ ] **Step 5: Modify `ChatTab._run_chat_legal_research`**

Start a task-debug run before calling the service. Pass a debug callback that forwards service event dictionaries into `task_debug.emit_event(...)`. Finish the run on success, `ChatResearchError`, and generic exception. Keep existing chat-history italic status messages.

- [ ] **Step 6: Run focused chat tests**

Run: `python -m pytest tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py -q`

Expected: PASS.

## Task 7: Final Verification

**Files:**
- All changed files.

- [ ] **Step 1: Run focused task-debug and adjacent wizard/chat suites**

Run: `python -m pytest tests/test_task_debug.py tests/test_ui_task_debug_window.py tests/test_wizard/test_task_debug_instrumentation.py tests/test_wizard/test_task_tab.py tests/test_wizard/test_in_process_task_tab.py tests/test_chat/test_legal_research_service.py tests/test_chat/test_legal_research_ui.py -q`

Expected: PASS.

- [ ] **Step 2: Run py_compile on changed Python files**

Run: `python -m py_compile icharlotte_core/task_debug.py icharlotte_core/ui/task_debug_window.py icharlotte_core/ui/wizard/task_tab.py icharlotte_core/ui/wizard/in_process_task_tab.py icharlotte_core/chat/legal_research.py icharlotte_core/ui/tabs.py iCharlotte.py`

Expected: exit code 0.

- [ ] **Step 3: Run diff whitespace check**

Run: `git diff --check`

Expected: no output and exit code 0.

- [ ] **Step 4: Inspect working tree**

Run: `git status --short`

Expected: changed task-debug feature files plus pre-existing unrelated dirty/untracked files. Do not stage unrelated files.

## Implementation Notes

- The current workspace has a pre-existing uncommitted change in `iCharlotte.py` around the change-file case switch path. Do not revert or overwrite it.
- Avoid dumping full prompts, full document text, full opinion text, or API keys into debug details.
- The debug recorder must never raise into task execution.
- Use explicit pytest file lists in PowerShell.
