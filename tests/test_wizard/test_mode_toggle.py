"""Smoke tests for ModeToggle segmented control."""
import pytest

pytest.importorskip("pytestqt")
from PySide6.QtCore import QCoreApplication, QSettings, Qt

from icharlotte_core.ui.wizard.mode_controller import ModeController, MODE_ADVANCED, MODE_WIZARD
from icharlotte_core.ui.wizard.mode_toggle import ModeToggle


@pytest.fixture(autouse=True)
def _qsettings_org(tmp_path):
    QCoreApplication.setOrganizationName("iCharlotteTest")
    QCoreApplication.setApplicationName(f"WizardTest-{tmp_path.name}")
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    s = QSettings()
    s.clear()
    s.sync()
    yield
    s.clear()
    s.sync()


def test_toggle_reflects_initial_mode(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    w = ModeToggle(ctrl)
    qtbot.addWidget(w)
    assert w.wizard_button.isChecked() is True
    assert w.advanced_button.isChecked() is False


def test_clicking_advanced_updates_controller(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    w = ModeToggle(ctrl)
    qtbot.addWidget(w)
    qtbot.mouseClick(w.advanced_button, Qt.MouseButton.LeftButton)
    assert ctrl.mode == MODE_ADVANCED


def test_external_mode_change_updates_buttons(qtbot):
    ctrl = ModeController()
    ctrl.set_mode(MODE_WIZARD)
    w = ModeToggle(ctrl)
    qtbot.addWidget(w)
    ctrl.set_mode(MODE_ADVANCED)
    assert w.advanced_button.isChecked() is True
    assert w.wizard_button.isChecked() is False
