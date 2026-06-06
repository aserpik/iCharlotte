import pytest

pytest.importorskip("pytestqt")

from PySide6.QtWidgets import QApplication


def _messages(window):
    return [
        window.table.item(row, window.COL_MESSAGE).text()
        for row in range(window.table.rowCount())
    ]


def test_debug_console_loads_buffer_and_filters(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path)
    run_id = td.start_run(task_id="task_a", task_title="Task A")
    td.emit_event(
        run_id=run_id,
        task_id="task_a",
        task_title="Task A",
        phase="search",
        message="Local hit",
        source="Local",
    )
    td.emit_event(
        run_id=run_id,
        task_id="task_a",
        task_title="Task A",
        phase="verify",
        message="Verifier warning",
        level="warning",
        source="Verifier",
    )

    window = TaskDebugWindow()
    qtbot.addWidget(window)
    assert window.table.rowCount() == 3

    window.level_filter.setCurrentText("warning")
    assert window.table.rowCount() == 1
    assert "Verifier warning" in window.table.item(0, window.COL_MESSAGE).text()


def test_debug_console_filters_by_task_and_source(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path)
    run_a = td.start_run(task_id="task_a", task_title="Task A", source="Runner")
    td.emit_event(
        run_id=run_a,
        task_id="task_a",
        task_title="Task A",
        phase="search",
        message="local task a",
        source="Local",
    )
    run_b = td.start_run(task_id="task_b", task_title="Task B", source="Runner")
    td.emit_event(
        run_id=run_b,
        task_id="task_b",
        task_title="Task B",
        phase="verify",
        message="verifier task b",
        source="Verifier",
    )

    window = TaskDebugWindow()
    qtbot.addWidget(window)

    window.task_filter.setCurrentText("Task B (task_b)")
    assert window.table.rowCount() == 2
    assert all("Task B" in window.table.item(row, window.COL_TASK).text() for row in range(2))

    window.source_filter.setCurrentText("Verifier")
    assert window.table.rowCount() == 1
    assert _messages(window) == ["verifier task b"]


def test_debug_console_search_matches_details_text(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path)
    run_id = td.start_run(task_id="chat", task_title="Chat")
    td.emit_event(
        run_id=run_id,
        task_id="chat",
        task_title="Chat",
        phase="search",
        message="plain status",
        details={"citation": "Needle v. Haystack"},
    )
    td.emit_event(
        run_id=run_id,
        task_id="chat",
        task_title="Chat",
        phase="search",
        message="unrelated",
        details={"citation": "Other v. Case"},
    )

    window = TaskDebugWindow()
    qtbot.addWidget(window)
    window.search_input.setText("needle")

    assert window.table.rowCount() == 1
    assert _messages(window) == ["plain status"]


def test_debug_console_copies_visible_rows_as_tab_separated_text(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path)
    run_id = td.start_run(task_id="task_a", task_title="Task A")
    td.emit_event(
        run_id=run_id,
        task_id="task_a",
        task_title="Task A",
        phase="verify",
        level="warning",
        message="copy me",
        source="Verifier",
    )
    td.emit_event(
        run_id=run_id,
        task_id="task_a",
        task_title="Task A",
        phase="search",
        message="do not copy",
        source="Local",
    )

    window = TaskDebugWindow()
    qtbot.addWidget(window)
    window.level_filter.setCurrentText("warning")
    window.copy_btn.click()

    clipboard_text = QApplication.clipboard().text()
    assert clipboard_text.count("\n") == 0
    assert "\t" in clipboard_text
    assert "copy me" in clipboard_text
    assert "do not copy" not in clipboard_text


def test_debug_console_mirrors_bounded_recorder_buffer_on_new_events(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path, max_events=2)
    window = TaskDebugWindow()
    qtbot.addWidget(window)
    run_id = td.start_run(task_id="task_a", task_title="Task A")
    for message in ["one", "two", "three"]:
        td.emit_event(
            run_id=run_id,
            task_id="task_a",
            task_title="Task A",
            phase="status",
            message=message,
        )

    assert len(td.get_events()) == 2
    assert window.table.rowCount() == 2
    assert _messages(window) == ["two", "three"]


def test_debug_console_pause_autoscroll_suppresses_scroll_to_bottom(
    qtbot, tmp_path, monkeypatch
):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path)
    window = TaskDebugWindow()
    qtbot.addWidget(window)
    calls = []
    monkeypatch.setattr(window.table, "scrollToBottom", lambda: calls.append("scroll"))
    run_id = td.start_run(task_id="chat", task_title="Chat")
    calls.clear()

    td.emit_event(
        run_id=run_id,
        task_id="chat",
        task_title="Chat",
        phase="status",
        message="first",
    )
    assert calls == ["scroll"]

    window.pause_autoscroll_check.setChecked(True)
    td.emit_event(
        run_id=run_id,
        task_id="chat",
        task_title="Chat",
        phase="status",
        message="second",
    )
    assert calls == ["scroll"]


def test_debug_console_open_folder_tolerates_startfile_failure(
    qtbot, tmp_path, monkeypatch
):
    import icharlotte_core.task_debug as td
    import icharlotte_core.ui.task_debug_window as mod
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path)
    window = TaskDebugWindow()
    qtbot.addWidget(window)

    monkeypatch.setattr(
        mod.os,
        "startfile",
        lambda _path: (_ for _ in ()).throw(OSError("cannot open")),
        raising=False,
    )
    window.open_folder_btn.click()


def test_debug_console_receives_new_events_and_clear_clears_recorder(qtbot, tmp_path):
    import icharlotte_core.task_debug as td
    from icharlotte_core.ui.task_debug_window import TaskDebugWindow

    td.reset_for_tests(trace_dir=tmp_path)
    window = TaskDebugWindow()
    qtbot.addWidget(window)
    run_id = td.start_run(task_id="chat", task_title="Chat")
    td.emit_event(
        run_id=run_id,
        task_id="chat",
        task_title="Chat",
        phase="status",
        message="working",
    )
    assert any(
        "working" in window.table.item(row, window.COL_MESSAGE).text()
        for row in range(window.table.rowCount())
    )
    window.clear_btn.click()
    assert window.table.rowCount() == 0
    assert td.get_events() == []
