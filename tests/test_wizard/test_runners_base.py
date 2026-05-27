"""Tests for BaseWorker contract — cancel flag, signal surface."""
import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.runners.base import BaseWorker


def test_cancel_sets_flag(qtbot):
    w = BaseWorker(case_path="/tmp/case", file_number="0000.000", files=[], settings={})
    assert w.is_cancel_requested is False
    w.cancel()
    assert w.is_cancel_requested is True


def test_signals_present(qtbot):
    w = BaseWorker(case_path="/tmp/case", file_number="0000.000", files=[], settings={})
    # Just touch them to ensure the attributes exist.
    assert w.status is not None
    assert w.progress is not None
    assert w.finished is not None
    assert w.failed is not None
    assert w.cancelled is not None
