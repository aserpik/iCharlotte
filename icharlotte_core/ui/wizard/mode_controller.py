"""ModeController — coordinates Advanced vs Wizard mode app-wide.

Mode is stored globally in QSettings (one toggle for the whole app, not
per-case). On change, emits `mode_changed(str)` so MainWindow can update
tab visibility.
"""
from PySide6.QtCore import QObject, QSettings, Signal


MODE_ADVANCED = "advanced"
MODE_WIZARD = "wizard"
_VALID_MODES = {MODE_ADVANCED, MODE_WIZARD}

_SETTINGS_KEY = "app/mode"
_DEFAULT_MODE = MODE_WIZARD


class ModeController(QObject):
    """Global mode coordinator. Backed by QSettings."""

    mode_changed = Signal(str)  # emits new mode value

    def __init__(self, parent: QObject | None = None):
        super().__init__(parent)
        self._settings = QSettings()
        self._mode = self._read_mode()

    def _read_mode(self) -> str:
        value = self._settings.value(_SETTINGS_KEY, _DEFAULT_MODE)
        if value not in _VALID_MODES:
            return _DEFAULT_MODE
        return value

    @property
    def mode(self) -> str:
        return self._mode

    @property
    def is_wizard(self) -> bool:
        return self._mode == MODE_WIZARD

    @property
    def is_advanced(self) -> bool:
        return self._mode == MODE_ADVANCED

    def set_mode(self, mode: str) -> None:
        if mode not in _VALID_MODES:
            raise ValueError(f"Invalid mode: {mode!r}. Must be one of {_VALID_MODES}.")
        if mode == self._mode:
            return
        self._mode = mode
        self._settings.setValue(_SETTINGS_KEY, mode)
        self._settings.sync()
        self.mode_changed.emit(mode)
