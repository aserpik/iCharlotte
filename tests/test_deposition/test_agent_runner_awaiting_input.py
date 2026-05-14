"""Tests for AgentRunner's AWAITING_INPUT handling."""

from unittest.mock import MagicMock

import pytest

pytest.importorskip("pytestqt")  # PySide6 + pytest-qt required

from icharlotte_core.ui.widgets import AgentRunner


def test_agent_runner_emits_awaiting_input_signal_on_token(qtbot):
    runner = AgentRunner("python", ["--phase=topics", "X.pdf"])
    received = []
    runner.awaiting_input.connect(received.append)

    runner.parse_progress("AWAITING_INPUT:C:\\tmp\\session.json\n")

    assert received == ["C:\\tmp\\session.json"]
    assert runner.session_path == "C:\\tmp\\session.json"


def test_agent_runner_does_not_emit_finished_when_paused(qtbot):
    runner = AgentRunner("python", ["--phase=topics", "X.pdf"])
    finished_calls = []
    awaiting_calls = []
    runner.finished.connect(finished_calls.append)
    runner.awaiting_input.connect(awaiting_calls.append)

    # Simulate phase-1 token then a normal exit(0).
    runner.parse_progress("AWAITING_INPUT:C:\\tmp\\session.json\n")
    runner.process = MagicMock()
    runner.process.deleteLater = MagicMock()
    from PySide6.QtCore import QProcess
    runner.handle_finished(0, QProcess.ExitStatus.NormalExit)

    assert awaiting_calls == ["C:\\tmp\\session.json"]
    assert finished_calls == []
    assert runner.success is None  # still running from the UI's perspective


def test_agent_runner_emits_finished_when_no_pause(qtbot):
    """Sanity check: without an AWAITING_INPUT token, exit 0 still emits finished(True)."""
    runner = AgentRunner("python", ["X.pdf"])
    finished_calls = []
    runner.finished.connect(finished_calls.append)

    runner.process = MagicMock()
    runner.process.deleteLater = MagicMock()
    from PySide6.QtCore import QProcess
    runner.handle_finished(0, QProcess.ExitStatus.NormalExit)

    assert finished_calls == [True]


def test_resume_with_config_starts_phase_two_process(qtbot, monkeypatch):
    runner = AgentRunner("python", ["--phase=topics", "X.pdf"])
    runner.session_path = r"C:\tmp\session.json"

    started_with = {}

    class FakeProcess:
        # Class-level MagicMocks so each instance shares the spec but instance methods still work.
        # The .connect on these will be a callable MagicMock, matching what real QProcess signals support.
        def __init__(self):
            self.readyReadStandardOutput = MagicMock()
            self.readyReadStandardError = MagicMock()
            self.finished = MagicMock()

        def start(self, cmd, args):
            started_with["cmd"] = cmd
            started_with["args"] = args

        def state(self):
            return 0

        def kill(self):
            pass

        def deleteLater(self):
            pass

    monkeypatch.setattr("icharlotte_core.ui.widgets.QProcess", FakeProcess)

    runner.resume_with_config(r"C:\tmp\session.json")

    assert started_with["cmd"] == "python"
    assert "--phase=summary" in started_with["args"]
    assert r"C:\tmp\session.json" in started_with["args"]
