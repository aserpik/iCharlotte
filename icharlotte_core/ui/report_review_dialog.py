"""
Report Review Dialog — popup for configuring report review options.

Shown when user clicks "Review Report" in the AI Assistant panel.
"""

import os
from typing import Dict, List

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QCheckBox, QRadioButton, QButtonGroup, QFrame, QFileDialog,
    QWidget, QScrollArea, QDialogButtonBox,
)
from PySide6.QtCore import Qt


class ReportReviewDialog(QDialog):
    """Dialog for configuring report review options before starting."""

    # Dark theme matching AI Assistant panel style
    STYLE = """
        QDialog {
            background-color: #1e1e2e;
            color: #cdd6f4;
        }
        QLabel {
            color: #cdd6f4;
        }
        QRadioButton, QCheckBox {
            color: #cdd6f4;
            font-size: 11px;
        }
        QRadioButton::indicator, QCheckBox::indicator {
            width: 14px; height: 14px;
        }
        QPushButton {
            background: #313244;
            color: #cdd6f4;
            border: 1px solid #45475a;
            border-radius: 4px;
            padding: 6px 14px;
            font-size: 11px;
        }
        QPushButton:hover {
            background: #45475a;
            border-color: #6c63ff;
        }
        QPushButton#startBtn {
            background: #6c63ff;
            color: #ffffff;
            font-weight: bold;
        }
        QPushButton#startBtn:hover {
            background: #7c73ff;
        }
        QFrame#separator {
            background-color: #45475a;
            max-height: 1px;
        }
        QFrame#chipFrame {
            background: #313244;
            border: 1px solid #45475a;
            border-radius: 3px;
            padding: 2px 6px;
        }
    """

    def __init__(self, has_selection: bool = False, case_detected: bool = False,
                 parent=None):
        """
        Args:
            has_selection: Whether text is currently selected in Word
            case_detected: Whether a case was detected from the document path
            parent: Parent widget
        """
        super().__init__(parent)
        self.setWindowTitle("Report Review")
        self.setFixedWidth(420)
        self.setStyleSheet(self.STYLE)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)

        self._extra_doc_paths: List[str] = []
        self._has_selection = has_selection
        self._case_detected = case_detected

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Title
        title = QLabel("Report Review Options")
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #cdd6f4;")
        layout.addWidget(title)

        # --- Review Scope ---
        scope_label = QLabel("Review scope:")
        scope_label.setStyleSheet("font-size: 11px; color: #a6adc8;")
        layout.addWidget(scope_label)

        self._scope_group = QButtonGroup(self)
        self._full_radio = QRadioButton("Full document")
        self._selection_radio = QRadioButton("Selected text only")
        self._scope_group.addButton(self._full_radio, 0)
        self._scope_group.addButton(self._selection_radio, 1)

        # Auto-set based on Word selection
        if self._has_selection:
            self._selection_radio.setChecked(True)
        else:
            self._full_radio.setChecked(True)
            self._selection_radio.setEnabled(False)
            self._selection_radio.setToolTip("Select text in Word first to enable this option")

        layout.addWidget(self._full_radio)
        layout.addWidget(self._selection_radio)

        # --- Separator ---
        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)

        # --- Additional Context ---
        ctx_label = QLabel("Additional Context")
        ctx_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #cdd6f4;")
        layout.addWidget(ctx_label)

        # Case data checkbox
        self._case_data_check = QCheckBox("Include case data (AI OUTPUT folder)")
        self._case_data_check.setToolTip(
            "Load agent outputs from the case's AI OUTPUT folder\n"
            "to check for missing facts and details"
        )
        if not self._case_detected:
            self._case_data_check.setEnabled(False)
            self._case_data_check.setToolTip(
                "No case detected from document path.\n"
                "Save the document in a case folder first."
            )
        layout.addWidget(self._case_data_check)

        # --- Additional Documents ---
        docs_label = QLabel("Attach additional documents:")
        docs_label.setStyleSheet("font-size: 11px; color: #a6adc8;")
        layout.addWidget(docs_label)

        # Add files button
        add_row = QHBoxLayout()
        self._add_btn = QPushButton("+ Add files...")
        self._add_btn.setFixedWidth(110)
        self._add_btn.clicked.connect(self._browse_files)
        add_row.addWidget(self._add_btn)
        add_row.addStretch()
        layout.addLayout(add_row)

        # Chips area (scrollable)
        self._chips_widget = QWidget()
        self._chips_layout = QVBoxLayout(self._chips_widget)
        self._chips_layout.setContentsMargins(0, 0, 0, 0)
        self._chips_layout.setSpacing(4)

        self._chips_scroll = QScrollArea()
        self._chips_scroll.setWidgetResizable(True)
        self._chips_scroll.setWidget(self._chips_widget)
        self._chips_scroll.setMaximumHeight(100)
        self._chips_scroll.setStyleSheet(
            "QScrollArea { border: none; background: transparent; }"
        )
        self._chips_scroll.setVisible(False)
        layout.addWidget(self._chips_scroll)

        # --- Separator ---
        sep2 = QFrame()
        sep2.setObjectName("separator")
        sep2.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep2)

        # --- Buttons ---
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        start_btn = QPushButton("Start Review")
        start_btn.setObjectName("startBtn")
        start_btn.clicked.connect(self.accept)
        btn_row.addWidget(start_btn)

        layout.addLayout(btn_row)

    def _browse_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Attach Documents", "",
            "Documents (*.doc *.docx *.pdf);;All Files (*)"
        )
        for path in paths:
            if path not in self._extra_doc_paths:
                self._extra_doc_paths.append(path)
                self._add_chip(path)

    def _add_chip(self, file_path: str):
        """Add a file chip with remove button."""
        chip = QFrame()
        chip.setObjectName("chipFrame")
        chip_layout = QHBoxLayout(chip)
        chip_layout.setContentsMargins(6, 2, 2, 2)
        chip_layout.setSpacing(4)

        name_label = QLabel(os.path.basename(file_path))
        name_label.setStyleSheet("font-size: 10px; color: #cdd6f4;")
        name_label.setToolTip(file_path)
        chip_layout.addWidget(name_label, 1)

        remove_btn = QPushButton("\u2715")  # ✕
        remove_btn.setFixedSize(18, 18)
        remove_btn.setStyleSheet(
            "QPushButton { background: transparent; border: none; "
            "color: #f38ba8; font-size: 12px; padding: 0; }"
            "QPushButton:hover { color: #ff0000; }"
        )
        remove_btn.clicked.connect(lambda: self._remove_chip(file_path, chip))
        chip_layout.addWidget(remove_btn)

        self._chips_layout.addWidget(chip)
        self._chips_scroll.setVisible(True)

    def _remove_chip(self, file_path: str, chip: QFrame):
        """Remove a file chip and its path."""
        if file_path in self._extra_doc_paths:
            self._extra_doc_paths.remove(file_path)
        chip.deleteLater()
        if not self._extra_doc_paths:
            self._chips_scroll.setVisible(False)

    def get_options(self) -> Dict:
        """Return the configured review options."""
        return {
            "scope": "selection" if self._selection_radio.isChecked() else "full",
            "include_case_data": self._case_data_check.isChecked(),
            "extra_doc_paths": list(self._extra_doc_paths),
        }
