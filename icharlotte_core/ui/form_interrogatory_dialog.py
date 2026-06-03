"""Manual confirmation dialog for Form Interrogatory selection.

Auto-detecting which checkboxes are marked on a flattened FROG PDF is
best-effort and can never be perfectly reliable across every way the form gets
filled and flattened. This dialog shows every interrogatory found on the form
with the auto-detected ones pre-checked, so the attorney confirms (or fixes)
exactly which interrogatories were propounded before responses are drafted.

Implementation note: this uses real ``QCheckBox`` rows inside a scroll area
rather than checkable ``QListWidget`` items. The wizard theme styles
``QListWidget::item`` via QSS, which makes Qt stop drawing item check
indicators (checked rows render blank, clicks look dead). Standalone
``QCheckBox`` widgets are not affected and match the rest of the wizard.
"""
from __future__ import annotations

from typing import Dict, Iterable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


def _number_key(number: str) -> tuple[int, ...]:
    parts = []
    for piece in str(number).split("."):
        try:
            parts.append(int(piece))
        except ValueError:
            parts.append(0)
    return tuple(parts)


class FormInterrogatorySelectionDialog(QDialog):
    """Checklist of FROG interrogatories; pre-checked from auto-detection."""

    def __init__(
        self,
        interrogatories: Iterable,
        parent: QWidget | None = None,
        *,
        title: str = "Select propounded Form Interrogatories",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(640, 560)

        # number -> its checkbox, in display (natural-sorted) order.
        self.checkboxes: Dict[str, QCheckBox] = {}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(10)

        header = QLabel(
            "Check the Form Interrogatories that were propounded. The boxes "
            "checked below were auto-detected from the PDF — review and adjust "
            "them; only checked interrogatories will be answered."
        )
        header.setWordWrap(True)
        outer.addWidget(header)

        btn_row = QHBoxLayout()
        self.all_btn = QPushButton("Select all")
        self.all_btn.clicked.connect(lambda: self.set_all_checked(True))
        self.none_btn = QPushButton("Select none")
        self.none_btn.clicked.connect(lambda: self.set_all_checked(False))
        btn_row.addWidget(self.all_btn)
        btn_row.addWidget(self.none_btn)
        btn_row.addStretch()
        outer.addLayout(btn_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        rows_host = QWidget()
        rows_layout = QVBoxLayout(rows_host)
        rows_layout.setContentsMargins(4, 4, 4, 4)
        rows_layout.setSpacing(2)

        for item in sorted(interrogatories, key=lambda s: _number_key(s.number)):
            number = str(item.number)
            text = (getattr(item, "text", "") or "").strip()
            display = f"{number}   {text}".strip() if text else number
            if len(display) > 120:
                display = display[:117] + "…"
            cb = QCheckBox(display)
            cb.setChecked(bool(getattr(item, "checked", False)))
            if text:
                cb.setToolTip(f"{number}: {text}")
            cb.toggled.connect(self._refresh_count)
            rows_layout.addWidget(cb)
            self.checkboxes[number] = cb
        rows_layout.addStretch(1)

        scroll.setWidget(rows_host)
        outer.addWidget(scroll, 1)

        self.count_label = QLabel("")
        self.count_label.setStyleSheet("color: #666;")
        outer.addWidget(self.count_label)
        self._refresh_count()

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        outer.addWidget(button_box)

    def set_all_checked(self, checked: bool) -> None:
        for cb in self.checkboxes.values():
            cb.setChecked(checked)

    def selected_numbers(self) -> List[str]:
        chosen = [number for number, cb in self.checkboxes.items() if cb.isChecked()]
        return sorted(chosen, key=_number_key)

    def _refresh_count(self, *_args) -> None:
        total = len(self.checkboxes)
        chosen = sum(1 for cb in self.checkboxes.values() if cb.isChecked())
        self.count_label.setText(f"{chosen} of {total} interrogatories selected")

    @classmethod
    def get_selected_numbers(
        cls,
        interrogatories: Iterable,
        parent: QWidget | None = None,
        *,
        title: str = "Select propounded Form Interrogatories",
    ) -> Optional[List[str]]:
        """Show modally. Return the confirmed numbers on OK, ``None`` on Cancel."""
        dlg = cls(interrogatories, parent, title=title)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            return dlg.selected_numbers()
        return None
