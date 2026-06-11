import pytest

pytest.importorskip("pytestqt")


def _messages(events):
    return [event.message for event in events]


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

    monkeypatch.setattr(
        "icharlotte_core.ui.task_debug_window.TaskDebugWindow",
        FakeDebugWindow,
    )
    window = app_mod.MainWindow.__new__(app_mod.MainWindow)
    window.view_menu = app_mod.QMenu()
    window._task_debug_window = None

    app_mod.MainWindow._add_debug_console_action(window)
    app_mod.MainWindow._add_debug_console_action(window)

    actions = [action.text() for action in window.view_menu.actions()]
    assert actions.count("Debug Console") == 1
    window.view_menu.actions()[-1].trigger()
    window.view_menu.actions()[-1].trigger()
    assert called.count(("init", True)) == 1
    assert called.count(("show", True)) == 2
    assert called.count(("raise", True)) == 2
    assert called.count(("activate", True)) == 2
    assert ("init", True) in called
    assert ("show", True) in called
    assert ("raise", True) in called
    assert ("activate", True) in called


def test_subprocess_task_tab_records_debug_lifecycle(monkeypatch, qtbot, tmp_path):
    from PySide6.QtCore import QObject, Signal

    from icharlotte_core import task_debug
    from icharlotte_core.ui.wizard.registry import get_task
    from icharlotte_core.ui.wizard.task_tab import TaskTab
    import icharlotte_core.ui.wizard.runners.subprocess_worker as sw_mod

    task_debug.reset_for_tests(trace_dir=tmp_path / "traces", max_events=20)
    output_path = str(tmp_path / "summary.txt")

    class FakeWorker(QObject):
        status = Signal(str)
        progress = Signal(int)
        awaiting_input = Signal(str)
        finished = Signal(str)
        failed = Signal(str)
        cancelled = Signal()

        def __init__(self, *args, **kwargs):
            super().__init__()

        def start(self):
            self.status.emit("Reading input file")
            self.progress.emit(35)
            self.finished.emit(output_path)

        def cancel(self):
            pass

    monkeypatch.setattr(sw_mod, "SubprocessWorker", FakeWorker)

    tab = TaskTab(
        get_task("summarize_documents"),
        files=[str(tmp_path / "source.pdf")],
        case_path=str(tmp_path),
        file_number="1234.001",
    )
    qtbot.addWidget(tab)

    tab._start_run({"include_timeline": True})

    events = task_debug.get_events()
    assert _messages(events) == [
        "Task started",
        "Starting Summarize Documents\u2026",
        "Reading input file",
        "Progress 35%",
        "Task complete",
    ]
    progress = next(event for event in events if event.phase == "progress")
    assert progress.details["progress"] == 35
    assert events[-1].details["output_path"] == output_path


def test_in_process_task_tab_records_progress_warning_and_finish(qtbot, tmp_path):
    from PySide6.QtCore import QObject, Signal
    from PySide6.QtWidgets import QWidget

    from icharlotte_core import task_debug
    from icharlotte_core.ui.wizard.in_process_task_tab import InProcessTaskTab
    from icharlotte_core.ui.wizard.pages.output_page import OutputPage
    from icharlotte_core.ui.wizard.registry import get_task

    class SettingsWidget(QWidget):
        run_requested = Signal(dict)

    class FakeWorker(QObject):
        progress = Signal(str)
        warning = Signal(str)
        finished_result = Signal(bool, str)

        def start(self):
            self.progress.emit("Loading records")
            self.warning.emit("Skipped encrypted attachment")
            self.finished_result.emit(True, str(tmp_path / "result.txt"))

    task_debug.reset_for_tests(trace_dir=tmp_path / "traces", max_events=20)
    tab = InProcessTaskTab(
        spec=get_task("subpoena_tracker"),
        case_path=str(tmp_path),
        file_number="1234.001",
        settings_widget=SettingsWidget(),
        output_widget=OutputPage(),
        worker_factory=lambda cp, fn, settings, parent: FakeWorker(parent),
    )
    qtbot.addWidget(tab)

    tab._on_run({"source_count": 2})

    events = task_debug.get_events()
    assert "Task started" in _messages(events)
    assert "Loading records" in _messages(events)
    warning = next(event for event in events if "encrypted" in event.message)
    assert warning.level == "warning"
    assert events[-1].phase == "finish"
    assert events[-1].details["status"] == "success"


def test_separate_task_records_debug_events(monkeypatch, qtbot, tmp_path):
    pypdf = pytest.importorskip("pypdf")
    from PySide6.QtCore import QObject, Signal

    from icharlotte_core import task_debug
    from icharlotte_core.ui.wizard.pages import separate_page as page_mod
    from icharlotte_core.ui.wizard.registry import get_task

    pdf_path = tmp_path / "source.pdf"
    writer = pypdf.PdfWriter()
    writer.add_blank_page(width=200, height=200)
    with pdf_path.open("wb") as handle:
        writer.write(handle)

    class FakeWorker(QObject):
        progress = Signal(str)
        finished_analysis = Signal(bool, object)
        finished = Signal()

        def __init__(self, pdf_path, sensitivity, parent=None):
            super().__init__(parent)
            self.pdf_path = pdf_path
            self.sensitivity = sensitivity

        def start(self):
            self.progress.emit("Scanning source PDF")
            self.finished_analysis.emit(
                True,
                [{"id": "1", "title": "Document", "date": "", "start": 1, "end": 1}],
            )
            self.finished.emit()

        def isRunning(self):
            return False

    monkeypatch.setattr(page_mod, "SeparateAnalysisWorker", FakeWorker)
    task_debug.reset_for_tests(trace_dir=tmp_path / "traces", max_events=20)

    tab = page_mod.SeparateTaskTab(
        get_task("separate"),
        case_path=str(tmp_path),
        file_number="1234.001",
        pdf_path=str(pdf_path),
    )
    qtbot.addWidget(tab)

    tab._start_analysis(2)

    events = task_debug.get_events()
    assert "Task started" in _messages(events)
    assert "Analyzing..." in _messages(events)
    assert "Scanning source PDF" in _messages(events)
    assert events[-1].phase == "analysis_complete"

    tab._on_processing_complete({"output_folder": str(tmp_path / "separated")})

    events = task_debug.get_events()
    assert events[-1].phase == "finish"
    assert events[-1].details["status"] == "success"
    assert events[-1].details["output_folder"] == str(tmp_path / "separated")


def test_oppose_motion_task_records_draft_debug_events(monkeypatch, qtbot, tmp_path):
    from PySide6.QtCore import QObject, Signal

    from icharlotte_core import task_debug
    from icharlotte_core.opposition.models import DraftDocument
    from icharlotte_core.ui.wizard.pages import oppose_motion_page as page_mod

    class FakeWorker(QObject):
        progress = Signal(str)
        finished_result = Signal(bool, object)
        finished = Signal()

        def __init__(self, case_path, file_number, settings, parent=None):
            super().__init__(parent)
            self.settings = settings

        def start(self):
            self.progress.emit("Drafting legal argument")
            self.finished_result.emit(
                True,
                DraftDocument(
                    title="Opposition",
                    body_text="Argument.",
                    preview_path=str(tmp_path / "oppose.docx"),
                ),
            )
            self.finished.emit()

    monkeypatch.setattr(page_mod, "OpposeMotionWorker", FakeWorker)
    task_debug.reset_for_tests(trace_dir=tmp_path / "traces", max_events=20)

    tab = page_mod.OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path=str(tmp_path),
        file_number="1234.001",
        motion_file=str(tmp_path / "motion.pdf"),
        context_files=[],
    )
    qtbot.addWidget(tab)

    tab._on_run({"motion_file": str(tmp_path / "motion.pdf"), "outline": []})

    events = task_debug.get_events()
    assert "Task started" in _messages(events)
    assert "Drafting opposition memorandum..." in _messages(events)
    assert "Drafting legal argument" in _messages(events)
    assert events[-1].phase == "finish"
    assert events[-1].details["status"] == "success"


def test_oppose_motion_status_cancel_records_debug_event(monkeypatch, qtbot, tmp_path):
    from PySide6.QtCore import QObject, Signal

    from icharlotte_core import task_debug
    from icharlotte_core.ui.wizard.pages import oppose_motion_page as page_mod

    class FakeWorker(QObject):
        progress = Signal(str)
        finished_result = Signal(bool, object)
        finished = Signal()

        def __init__(self, case_path, file_number, settings, parent=None):
            super().__init__(parent)

        def start(self):
            self.progress.emit("Drafting legal argument")

        def isRunning(self):
            return False

    monkeypatch.setattr(page_mod, "OpposeMotionWorker", FakeWorker)
    task_debug.reset_for_tests(trace_dir=tmp_path / "traces", max_events=20)

    tab = page_mod.OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path=str(tmp_path),
        file_number="1234.001",
        motion_file=str(tmp_path / "motion.pdf"),
        context_files=[],
    )
    qtbot.addWidget(tab)

    tab._on_run({"motion_file": str(tmp_path / "motion.pdf"), "outline": []})
    tab.status_page.cancel_requested.emit()

    events = task_debug.get_events()
    assert any(event.phase == "cancel" for event in events)
    assert "cannot be cancelled" in tab.status_page.status_label.text()
