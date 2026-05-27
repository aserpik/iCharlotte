"""ModeToggle — segmented control for Advanced/Wizard mode selection."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QButtonGroup, QHBoxLayout, QPushButton, QWidget

from .mode_controller import ModeController, MODE_ADVANCED, MODE_WIZARD


_SEGMENTED_STYLE = """
QPushButton {
    background-color: #f5f5f5;
    color: #555;
    font-size: 12px;
    font-weight: 500;
    padding: 6px 16px;
    border: 1px solid #ccc;
    min-width: 110px;
    height: 28px;
}
QPushButton:checked {
    background-color: #1976D2;
    color: white;
    font-weight: 600;
    border-color: #0D47A1;
}
QPushButton:hover:!checked {
    background-color: #e8e8e8;
}
"""


class ModeToggle(QWidget):
    """Two-button segmented control bound to a ModeController.

    Listens to the controller's mode_changed signal so external mode
    changes (e.g. keyboard shortcut, programmatic) keep the UI in sync.
    """

    def __init__(self, controller: ModeController, parent: QWidget | None = None):
        super().__init__(parent)
        self._controller = controller

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.advanced_button = QPushButton("Advanced Mode")
        self.advanced_button.setCheckable(True)
        self.advanced_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.advanced_button.setStyleSheet(
            _SEGMENTED_STYLE + "QPushButton { border-top-left-radius: 4px; border-bottom-left-radius: 4px; border-right: none; }"
        )

        self.wizard_button = QPushButton("Wizard Mode")
        self.wizard_button.setCheckable(True)
        self.wizard_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.wizard_button.setStyleSheet(
            _SEGMENTED_STYLE + "QPushButton { border-top-right-radius: 4px; border-bottom-right-radius: 4px; }"
        )

        layout.addWidget(self.advanced_button)
        layout.addWidget(self.wizard_button)

        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._group.addButton(self.advanced_button)
        self._group.addButton(self.wizard_button)

        self.advanced_button.clicked.connect(lambda: self._controller.set_mode(MODE_ADVANCED))
        self.wizard_button.clicked.connect(lambda: self._controller.set_mode(MODE_WIZARD))

        self._controller.mode_changed.connect(self._sync_from_controller)
        self._sync_from_controller(self._controller.mode)

    def _sync_from_controller(self, mode: str) -> None:
        self.advanced_button.setChecked(mode == MODE_ADVANCED)
        self.wizard_button.setChecked(mode == MODE_WIZARD)
