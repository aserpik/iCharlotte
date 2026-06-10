"""SettingsPage — pre-run configuration for a task tab.

The generic page renders a source-review workbench for file-driven tasks:
  - A source queue with Add Files / Remove actions and readiness badges.
  - A task setup panel with file counts, start folder, and output status.
  - A Continue button bottom-right.

Subclasses still receive the legacy simple layout because several specialized
settings pages re-parent those base widgets into custom flows.
"""
import os
from typing import List

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from .. import theme
from ..file_picker import resolve_default_folder
from ..registry import TaskSpec


class SettingsPage(QWidget):
    """Configurable inputs + Proceed button. Emits proceed_requested(settings_dict)."""

    proceed_requested = Signal(dict)  # settings dict (placeholder)

    def __init__(
        self,
        spec: TaskSpec,
        files: List[str],
        case_root: str | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self._spec = spec
        self._files: List[str] = list(files)
        self._case_root: str | None = case_root
        self._workbench_enabled = type(self) is SettingsPage
        self._files_label_title = "Source Queue" if self._workbench_enabled else "Files"

        outer = QVBoxLayout(self)
        outer.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        outer.setSpacing(theme.SPACE_LG)

        if self._workbench_enabled:
            self._build_source_review_workbench(outer)
        else:
            self._build_legacy_layout(outer)

        self._update_proceed_enabled()

    def _build_legacy_layout(self, outer: QVBoxLayout) -> None:
        """Build the original simple layout used by SettingsPage subclasses."""
        # One-line instruction so the page is self-explanatory.
        self.instruction_label = theme.helper_text(
            "Choose the documents to include, then click Continue."
        )
        outer.addWidget(self.instruction_label)

        # Files section
        files_label = theme.section_header(self._format_files_label())
        outer.addWidget(files_label)
        self.files_label = files_label

        # Add Files... / Remove button row
        file_btn_row = QHBoxLayout()
        file_btn_row.setSpacing(theme.SPACE_SM)
        self.add_files_btn = theme.secondary_button("Add Files…")
        self.add_files_btn.clicked.connect(self._on_add_files)
        file_btn_row.addWidget(self.add_files_btn)
        self.remove_btn = theme.secondary_button("Remove")
        self.remove_btn.setEnabled(False)
        self.remove_btn.clicked.connect(self._on_remove_files)
        file_btn_row.addWidget(self.remove_btn)
        file_btn_row.addStretch()
        outer.addLayout(file_btn_row)

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(150)
        self.files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.files_list.itemSelectionChanged.connect(self._on_selection_changed)
        outer.addWidget(self.files_list)

        # Empty-state hint shown under the list when no files are selected.
        self.empty_hint = theme.caption("No files yet — click “Add Files…” to get started.")
        self.empty_hint.setStyleSheet(
            f"font-size: {theme.FONT_CAPTION}px; color: {theme.TEXT_FAINT}; font-style: italic;"
        )
        outer.addWidget(self.empty_hint)

        self._refresh_files_list()

        # Placeholder body
        body = QLabel(f"Settings for {self._spec.title} — to be defined.")
        body.setStyleSheet(
            f"color: {theme.TEXT_MUTED}; font-style: italic; padding: {theme.SPACE_XL}px;"
        )
        outer.addWidget(body, 1)

        # Continue button bottom-right
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.proceed_btn = theme.primary_button("Continue")
        self.proceed_btn.clicked.connect(self._on_proceed)
        btn_row.addWidget(self.proceed_btn)
        outer.addLayout(btn_row)

    def _build_source_review_workbench(self, outer: QVBoxLayout) -> None:
        self.instruction_label = theme.helper_text(
            "Review the source queue and task setup, then click Continue."
        )
        outer.addWidget(self.instruction_label)

        self.workbench_frame = QFrame()
        self.workbench_frame.setObjectName("sourceReviewWorkbench")
        self.workbench_frame.setStyleSheet(self._workbench_stylesheet())

        workbench_layout = QGridLayout(self.workbench_frame)
        workbench_layout.setContentsMargins(0, 0, 0, 0)
        workbench_layout.setHorizontalSpacing(theme.SPACE_LG)
        workbench_layout.setVerticalSpacing(0)
        workbench_layout.setColumnStretch(0, 3)
        workbench_layout.setColumnStretch(1, 2)

        self.source_queue_panel = self._build_source_queue_panel()
        self.task_setup_panel = self._build_task_setup_panel()
        workbench_layout.addWidget(self.source_queue_panel, 0, 0)
        workbench_layout.addWidget(self.task_setup_panel, 0, 1)

        outer.addWidget(self.workbench_frame, 1)

        footer_row = QHBoxLayout()
        footer_row.setContentsMargins(0, 0, 0, 0)
        self.footer_status = theme.caption("")
        self.footer_status.setObjectName("settingsFooterStatus")
        footer_row.addWidget(self.footer_status)
        footer_row.addStretch()
        self.proceed_btn = theme.primary_button("Continue")
        self.proceed_btn.clicked.connect(self._on_proceed)
        footer_row.addWidget(self.proceed_btn)
        outer.addLayout(footer_row)

        self._refresh_files_list()

    def _build_source_queue_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("sourceQueuePanel")
        panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(
            theme.SPACE_LG, theme.SPACE_LG, theme.SPACE_LG, theme.SPACE_LG
        )
        layout.setSpacing(theme.SPACE_MD)

        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        self.files_label = theme.section_header(self._format_files_label())
        header_row.addWidget(self.files_label)
        header_row.addStretch()

        self.add_files_btn = theme.primary_button("Add Files...")
        self.add_files_btn.clicked.connect(self._on_add_files)
        header_row.addWidget(self.add_files_btn)

        self.remove_btn = theme.secondary_button("Remove")
        self.remove_btn.setEnabled(False)
        self.remove_btn.clicked.connect(self._on_remove_files)
        header_row.addWidget(self.remove_btn)
        layout.addLayout(header_row)

        self.files_list = QListWidget()
        self.files_list.setObjectName("sourceQueueList")
        self.files_list.setMinimumHeight(210)
        self.files_list.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.files_list.setWordWrap(True)
        self.files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.files_list.itemSelectionChanged.connect(self._on_selection_changed)
        layout.addWidget(self.files_list, 1)

        self.empty_hint = QLabel()
        self.empty_hint.setObjectName("sourceDropZoneHint")
        self.empty_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_hint.setWordWrap(True)
        self.empty_hint.setMinimumHeight(86)
        layout.addWidget(self.empty_hint)

        return panel

    def _build_task_setup_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("taskSetupPanel")
        panel.setMinimumWidth(280)
        panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(
            theme.SPACE_LG, theme.SPACE_LG, theme.SPACE_LG, theme.SPACE_LG
        )
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.section_header("Task Setup"))

        self.source_count_value = self._metric_value_label()
        layout.addLayout(self._metric_row("Selected sources", self.source_count_value))

        self.ready_count_value = self._metric_value_label()
        layout.addLayout(self._metric_row("Ready files", self.ready_count_value))

        self.missing_count_value = self._metric_value_label()
        layout.addLayout(self._metric_row("Missing files", self.missing_count_value))

        self.start_folder_value = self._metric_value_label()
        self.start_folder_value.setWordWrap(True)
        layout.addLayout(self._metric_row("Start folder", self.start_folder_value))

        self.output_location_value = self._metric_value_label("AI OUTPUT")
        layout.addLayout(self._metric_row("Output", self.output_location_value))

        self.setup_status_label = QLabel()
        self.setup_status_label.setObjectName("setupStatusLabel")
        self.setup_status_label.setWordWrap(True)
        layout.addWidget(self.setup_status_label)
        layout.addStretch()
        return panel

    def _metric_value_label(self, text: str = "") -> QLabel:
        label = QLabel(text)
        label.setObjectName("metricValue")
        label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        return label

    def _metric_row(self, label_text: str, value_label: QLabel) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(theme.SPACE_MD)
        label = QLabel(label_text)
        label.setObjectName("metricLabel")
        row.addWidget(label)
        row.addStretch()
        row.addWidget(value_label)
        return row

    def _workbench_stylesheet(self) -> str:
        return f"""
        QFrame#sourceReviewWorkbench {{
            background: transparent;
        }}
        QFrame#sourceQueuePanel, QFrame#taskSetupPanel {{
            background-color: {theme.BG_CARD};
            border: 1px solid {theme.BORDER};
            border-radius: {theme.RADIUS_MD}px;
        }}
        QListWidget#sourceQueueList {{
            border: 1px solid {theme.BORDER_LIGHT};
            border-radius: {theme.RADIUS_MD}px;
            background-color: #FFFFFF;
            padding: 4px;
        }}
        QListWidget#sourceQueueList::item {{
            border: 1px solid {theme.BORDER_LIGHT};
            border-radius: {theme.RADIUS_SM}px;
            margin: 3px;
            padding: 8px;
        }}
        QListWidget#sourceQueueList::item:selected {{
            background-color: {theme.PRIMARY_SUBTLE};
            color: {theme.TEXT};
            border-color: {theme.PRIMARY};
        }}
        QLabel#sourceFileIcon {{
            background-color: #EEF3F8;
            border-radius: {theme.RADIUS_SM}px;
            color: {theme.PRIMARY};
            font-weight: 700;
            font-size: {theme.FONT_CAPTION}px;
        }}
        QLabel#sourceFileName {{
            color: {theme.TEXT};
            font-weight: 600;
            font-size: {theme.FONT_BODY}px;
        }}
        QLabel#sourceFileDetail {{
            color: {theme.TEXT_MUTED};
            font-size: {theme.FONT_CAPTION}px;
        }}
        QLabel#sourceFileStatusReady {{
            background-color: {theme.SUCCESS_BG};
            border-radius: {theme.RADIUS_SM}px;
            color: {theme.SUCCESS};
            font-size: {theme.FONT_CAPTION}px;
            font-weight: 700;
            padding: 3px 7px;
        }}
        QLabel#sourceFileStatusMissing {{
            background-color: {theme.ERROR_BG};
            border-radius: {theme.RADIUS_SM}px;
            color: {theme.ERROR};
            font-size: {theme.FONT_CAPTION}px;
            font-weight: 700;
            padding: 3px 7px;
        }}
        QLabel#sourceDropZoneHint {{
            border: 1px dashed #AEB9C6;
            border-radius: {theme.RADIUS_MD}px;
            background-color: #F8FBFF;
            color: {theme.TEXT_MUTED};
            padding: {theme.SPACE_LG}px;
        }}
        QLabel#metricLabel {{
            color: {theme.TEXT_MUTED};
            font-size: {theme.FONT_BODY}px;
        }}
        QLabel#metricValue {{
            color: {theme.TEXT};
            font-size: {theme.FONT_BODY}px;
            font-weight: 600;
        }}
        QLabel#setupStatusLabel {{
            background-color: {theme.PRIMARY_SUBTLE};
            border-radius: {theme.RADIUS_MD}px;
            color: {theme.TEXT_BODY};
            padding: {theme.SPACE_MD}px;
        }}
        """

    def _format_files_label(self) -> str:
        return f"{self._files_label_title} ({len(self._files)})"

    def _refresh_files_list(self) -> None:
        from PySide6.QtGui import QColor

        self.files_list.clear()
        for path in self._files:
            display = os.path.basename(path)
            if self._workbench_enabled:
                item_text = ""
            else:
                item_text = display
            item = QListWidgetItem(item_text)
            item.setToolTip(path)
            missing = not os.path.exists(path)
            if missing:
                if not self._workbench_enabled:
                    item.setText(f"{display}  (missing)")
                item.setForeground(QColor(theme.ERROR))
            if self._workbench_enabled:
                item.setSizeHint(QSize(0, 58))
            self.files_list.addItem(item)
            if self._workbench_enabled:
                self.files_list.setItemWidget(
                    item, self._build_source_file_row(path, missing)
                )
        self.files_label.setText(self._format_files_label())
        if hasattr(self, "empty_hint"):
            if self._workbench_enabled:
                if self._files:
                    self.empty_hint.setText(
                        "Add more documents to the queue with Add Files..."
                    )
                else:
                    self.empty_hint.setText(
                        "Add documents to this queue\nPDF, Word, and text files are supported."
                    )
                self.empty_hint.setVisible(True)
            else:
                self.empty_hint.setVisible(len(self._files) == 0)
        if self._workbench_enabled:
            self._refresh_workbench_summary()
        self._update_proceed_enabled()

    def _refresh_workbench_summary(self) -> None:
        selected_count = len(self._files)
        missing_count = sum(1 for path in self._files if not os.path.exists(path))
        ready_count = selected_count - missing_count

        self.source_count_value.setText(self._format_count(selected_count))
        self.ready_count_value.setText(self._format_count(ready_count))
        self.missing_count_value.setText(str(missing_count))

        start_dir = self._file_dialog_start_dir()
        start_label = self._format_start_folder(start_dir)
        self.start_folder_value.setText(start_label)
        self.start_folder_value.setToolTip(start_dir)

        if selected_count == 0:
            status = "Add at least one source document to continue."
        elif missing_count:
            status = "One or more selected files are missing. Remove or replace them before running."
        else:
            status = "Ready to run with the selected source documents."
        self.setup_status_label.setText(status)
        if hasattr(self, "footer_status"):
            self.footer_status.setText(status)

    def _format_count(self, count: int) -> str:
        suffix = "file" if count == 1 else "files"
        return f"{count} {suffix}"

    def _format_start_folder(self, start_dir: str) -> str:
        if not start_dir:
            return "Default folder"
        if self._case_root and os.path.normcase(
            os.path.normpath(start_dir)
        ) == os.path.normcase(os.path.normpath(self._case_root)):
            return "Case root"
        return os.path.basename(os.path.normpath(start_dir)) or start_dir

    def _build_source_file_row(self, path: str, missing: bool) -> QWidget:
        display = os.path.basename(path) or path
        folder = os.path.basename(os.path.dirname(path)) or os.path.dirname(path)
        ext = os.path.splitext(display)[1].lstrip(".")[:3].upper() or "FILE"

        row = QWidget()
        row.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(
            theme.SPACE_SM, theme.SPACE_XS, theme.SPACE_SM, theme.SPACE_XS
        )
        layout.setSpacing(theme.SPACE_SM)

        icon = QLabel(ext)
        icon.setObjectName("sourceFileIcon")
        icon.setFixedSize(34, 34)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon)

        text_col = QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(1)
        name_label = QLabel(display)
        name_label.setObjectName("sourceFileName")
        name_label.setToolTip(path)
        detail_label = QLabel(folder or path)
        detail_label.setObjectName("sourceFileDetail")
        detail_label.setToolTip(path)
        text_col.addWidget(name_label)
        text_col.addWidget(detail_label)
        layout.addLayout(text_col, 1)

        status = QLabel("Missing" if missing else "Ready")
        status.setObjectName(
            "sourceFileStatusMissing" if missing else "sourceFileStatusReady"
        )
        status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(status)
        return row

    def _update_proceed_enabled(self) -> None:
        if not hasattr(self, "proceed_btn"):
            return
        self.proceed_btn.setEnabled(len(self._files) > 0)

    def _on_selection_changed(self) -> None:
        self.remove_btn.setEnabled(len(self.files_list.selectedItems()) > 0)

    def _on_add_files(self) -> None:
        start_dir = self._file_dialog_start_dir()
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add files", start_dir, "All files (*.*)"
        )
        existing = set(self._files)
        for p in paths:
            if p not in existing:
                self._files.append(p)
                existing.add(p)
        self._refresh_files_list()

    def _file_dialog_start_dir(self) -> str:
        if not self._case_root:
            return ""
        return resolve_default_folder(self._case_root, self._spec.default_folders)

    def _on_remove_files(self) -> None:
        selected_rows = {idx.row() for idx in self.files_list.selectedIndexes()}
        self._files = [p for i, p in enumerate(self._files) if i not in selected_rows]
        self._refresh_files_list()

    def _on_proceed(self) -> None:
        self.proceed_requested.emit(self.to_dict())

    # ---- Persistence-friendly API ----

    @property
    def files(self) -> list[str]:
        return list(self._files)

    def to_dict(self) -> dict:
        """Placeholder settings dict. Real per-task settings will override."""
        return {}

    def from_dict(self, data: dict) -> None:
        """Placeholder — real subclasses will restore form state."""
        return None

    def attach_worker(self, worker) -> bool:
        """Override to take control of worker signals (e.g., for speculative runs).

        Return True if the settings page is handling the worker; in that case
        TaskTab will skip wiring status/progress/awaiting_input to the status page.
        """
        return False
