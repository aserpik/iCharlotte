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
