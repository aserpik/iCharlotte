"""Tests for ModeController — global Advanced/Wizard mode persistence."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import QSettings, QCoreApplication

from icharlotte_core.ui.wizard.mode_controller import ModeController, MODE_ADVANCED, MODE_WIZARD


@pytest.fixture(autouse=True)
def _qsettings_org(monkeypatch, tmp_path):
    """Isolate QSettings to a temp file per test."""
    QCoreApplication.setOrganizationName("iCharlotteTest")
    QCoreApplication.setApplicationName(f"WizardTest-{tmp_path.name}")
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    s = QSettings()
    s.clear()
    s.sync()
    yield
    s.clear()
    s.sync()


def test_first_run_defaults_to_wizard():
    ctrl = ModeController()
    assert ctrl.mode == MODE_WIZARD


def test_set_mode_persists_via_qsettings():
    ctrl1 = ModeController()
    ctrl1.set_mode(MODE_ADVANCED)
    # New instance must read the persisted value.
    ctrl2 = ModeController()
    assert ctrl2.mode == MODE_ADVANCED


def test_set_mode_emits_signal(qtbot):
    ctrl = ModeController()
    with qtbot.waitSignal(ctrl.mode_changed, timeout=500) as blocker:
        ctrl.set_mode(MODE_ADVANCED)
    assert blocker.args == [MODE_ADVANCED]


def test_set_mode_does_not_emit_when_unchanged(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_ADVANCED)  # change once
    # Now setting to the same value should NOT fire the signal.
    with qtbot.assertNotEmitted(ctrl.mode_changed, wait=200):
        ctrl.set_mode(MODE_ADVANCED)


def test_invalid_mode_raises():
    ctrl = ModeController()
    with pytest.raises(ValueError):
        ctrl.set_mode("nonsense")


def test_is_wizard_helper():
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    assert ctrl.is_wizard is True
    ctrl.set_mode(MODE_ADVANCED)
    assert ctrl.is_wizard is False
