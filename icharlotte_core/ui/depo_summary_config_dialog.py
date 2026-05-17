"""Modal dialog the user fills out after phase 1 of the deposition agent."""

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QVBoxLayout,
)

from icharlotte_core.ui.depo_summary_config_form import DepoSummaryConfigForm


class DepoSummaryConfigDialog(QDialog):
    def __init__(self, session_path, parent=None):
        super().__init__(parent)
        self.session_path = Path(session_path)

        self.setWindowTitle("Configure Deposition Summary")
        self.setModal(True)
        self.setWindowModality(Qt.ApplicationModal)
        self.resize(800, 750)

        root = QVBoxLayout(self)

        self.form = DepoSummaryConfigForm(session_path, parent=self)
        root.addWidget(self.form, 1)

        # Buttons
        buttons = QDialogButtonBox()
        buttons.addButton("Cancel", QDialogButtonBox.RejectRole)
        generate_btn = buttons.addButton("Generate Summary", QDialogButtonBox.AcceptRole)
        generate_btn.setDefault(True)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def accept(self):
        if not self.form.commit_user_config():
            return  # validation failed — form already showed error, don't close
        super().accept()
