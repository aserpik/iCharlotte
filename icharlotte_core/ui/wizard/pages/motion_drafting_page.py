"""Unified Wizard task page for motion drafting."""

from __future__ import annotations

import os
from datetime import datetime

from PySide6.QtWidgets import (
    QFileDialog,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.motion_drafting.taxonomy import (
    DRAFT_KIND_LABELS,
    DRAFT_KIND_MOTION,
    DRAFT_KIND_OPPOSITION,
    DRAFT_KIND_REPLY,
    MotionTypeOption,
    list_motion_type_options,
)
from icharlotte_core.opposition.extraction import (
    is_supported_context_file,
    is_supported_motion_file,
)
from icharlotte_core.opposition.models import DraftDocument, MotionMetadata, OutlineNode
from icharlotte_core.ui.context_files_dialog import ContextFilesDialog
from icharlotte_core.ui.wizard.file_drop import enable_file_drop
from icharlotte_core.ui.wizard.pages.generate_motion_page import (
    GenerateMotionOutputPage,
    GenerateMotionSettingsPage,
    GenerateMotionTaskTab,
    TASK_PAGE_OUTPUT,
    TASK_PAGE_SETTINGS,
    TASK_PAGE_STATUS,
)
from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
    OpposeMotionAnalysisWorker,
    OpposeMotionWorker,
)
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.task_debug_helpers import (
    finish_debug_run,
    record_status,
    start_debug_run,
)
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer


class MotionDraftingSettingsPage(GenerateMotionSettingsPage):
    """Settings page shared by motion, opposition, and reply drafting."""

    def _build_intake_page(self) -> None:
        self.draft_kind_combo = QComboBox()
        for kind, label in DRAFT_KIND_LABELS.items():
            self.draft_kind_combo.addItem(label, kind)
        self.draft_kind_combo.currentIndexChanged.connect(self._on_draft_kind_changed)

        self.motion_type_combo = QComboBox()
        self.motion_type_combo.currentIndexChanged.connect(self._on_type_changed)

        # The remainder intentionally mirrors GenerateMotionSettingsPage's
        # intake controls so the existing generate-motion worker can be reused.
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(8)

        row = QHBoxLayout()
        row.addWidget(QLabel("Draft"))
        row.addWidget(self.draft_kind_combo)
        row.addSpacing(16)
        row.addWidget(QLabel("Type"))
        row.addWidget(self.motion_type_combo, 1)
        layout.addLayout(row)

        self.motion_file_label = QLabel()
        enable_file_drop(
            self.motion_file_label,
            self.add_motion_files,
            path_filter=is_supported_motion_file,
        )
        layout.addWidget(self.motion_file_label)

        motion_row = QHBoxLayout()
        self.select_motion_btn = QPushButton("Select Motion")
        self.select_motion_btn.clicked.connect(self._on_select_motion_file)
        enable_file_drop(
            self.select_motion_btn,
            self.add_motion_files,
            path_filter=is_supported_motion_file,
        )
        motion_row.addWidget(self.select_motion_btn)
        motion_row.addStretch()
        layout.addLayout(motion_row)

        self.custom_name_edit = QLineEdit()
        self.custom_name_edit.setPlaceholderText("Name the motion type")
        self.custom_name_edit.setVisible(False)
        layout.addWidget(self.custom_name_edit)

        self.guidance_label = QLabel()
        self.guidance_label.setWordWrap(True)
        self.guidance_label.setStyleSheet("color: #6b7480;")
        layout.addWidget(self.guidance_label)

        files_row = QHBoxLayout()
        files_row.addWidget(QLabel("Context documents"))
        files_row.addStretch()
        self.add_files_btn = QPushButton("Add Files...")
        self.add_files_btn.clicked.connect(self._on_add_files)
        files_row.addWidget(self.add_files_btn)
        self.remove_files_btn = QPushButton("Remove")
        self.remove_files_btn.clicked.connect(self._on_remove_files)
        files_row.addWidget(self.remove_files_btn)
        layout.addLayout(files_row)

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(110)
        enable_file_drop(
            self.files_list,
            self.add_target_files,
            path_filter=is_supported_context_file,
        )
        enable_file_drop(page, self.add_target_files, path_filter=is_supported_context_file)
        layout.addWidget(self.files_list)

        layout.addWidget(QLabel("Relief requested"))
        self.user_relief_edit = QLineEdit()
        self.user_relief_edit.setPlaceholderText("What you are asking the court to order")
        layout.addWidget(self.user_relief_edit)

        layout.addWidget(QLabel("Arguments / grounds to include (one per line)"))
        self.user_arguments_edit = QPlainTextEdit()
        self.user_arguments_edit.setPlaceholderText(
            "Type the specific arguments or grounds you want included. The AI "
            "will analyze your documents and propose additional grounds to merge."
        )
        layout.addWidget(self.user_arguments_edit, 1)

        actions = QHBoxLayout()
        actions.addStretch()
        self.analyze_btn = QPushButton("Analyze & Continue")
        self.analyze_btn.clicked.connect(self._on_analyze_continue)
        actions.addWidget(self.analyze_btn)
        layout.addLayout(actions)

        self.motion_file = ""
        self._populate_motion_type_combo()
        self._refresh_motion_file_controls()
        self.addWidget(page)

    def current_draft_kind(self) -> str:
        return self.draft_kind_combo.currentData() or DRAFT_KIND_MOTION

    def current_motion_option(self) -> MotionTypeOption:
        option = self.motion_type_combo.currentData()
        if isinstance(option, MotionTypeOption):
            return option
        return MotionTypeOption(label="Motion", draft_kind=self.current_draft_kind())

    def current_motion_type_id(self) -> str:
        return self.current_motion_option().engine_type_id or "generic"

    def current_motion_type_name(self) -> str:
        label = self.current_motion_option().label or "Motion"
        if self.current_draft_kind() == DRAFT_KIND_REPLY:
            return f"Reply in Support of {label}"
        return label

    def selected_motion_type_label(self) -> str:
        return self.current_motion_option().label or ""

    def current_motion_type_source_path(self) -> str:
        return self.current_motion_option().source_path or ""

    def _on_draft_kind_changed(self, *_args) -> None:
        self._populate_motion_type_combo()
        self._refresh_motion_file_controls()
        self._refresh_outline_labels()

    def _populate_motion_type_combo(self) -> None:
        current_label = self.selected_motion_type_label()
        self.motion_type_combo.blockSignals(True)
        self.motion_type_combo.clear()
        for option in list_motion_type_options(self.current_draft_kind()):
            self.motion_type_combo.addItem(option.label, option)
        idx = self._find_motion_type_label(current_label)
        if idx >= 0:
            self.motion_type_combo.setCurrentIndex(idx)
        self.motion_type_combo.blockSignals(False)
        self._on_type_changed()

    def _find_motion_type_label(self, label: str) -> int:
        if not label:
            return -1
        for index in range(self.motion_type_combo.count()):
            option = self.motion_type_combo.itemData(index)
            if isinstance(option, MotionTypeOption) and option.label == label:
                return index
        return -1

    def _on_type_changed(self, *_args) -> None:
        option = self.current_motion_option()
        if option.engine_type_id == "generic":
            self.guidance_label.setText("Add the documents and arguments needed for this draft.")
        else:
            from icharlotte_core.motion_generation.config import get_motion_config

            self.guidance_label.setText(get_motion_config(option.engine_type_id).target_doc_guidance)

    def _refresh_motion_file_controls(self) -> None:
        show_motion_file = self.current_draft_kind() == DRAFT_KIND_OPPOSITION
        self.motion_file_label.setVisible(show_motion_file)
        self.select_motion_btn.setVisible(show_motion_file)
        self._refresh_motion_file_label()
        if show_motion_file:
            self.add_files_btn.setText("Add Context...")
        else:
            self.add_files_btn.setText("Add Files...")

    def _refresh_outline_labels(self) -> None:
        if not hasattr(self, "generate_btn"):
            return
        if self.current_draft_kind() == DRAFT_KIND_REPLY:
            self.generate_btn.setText("Generate Reply")
        elif self.current_draft_kind() == DRAFT_KIND_OPPOSITION:
            self.generate_btn.setText("Generate Opposition")
        else:
            self.generate_btn.setText("Generate Motion")

    def _refresh_motion_file_label(self) -> None:
        motion_name = os.path.basename(self.motion_file) or "(no motion selected)"
        self.motion_file_label.setText(f"Motion to oppose: {motion_name}")

    def _on_select_motion_file(self) -> None:
        motion_file, _ = QFileDialog.getOpenFileName(
            self,
            "Select motion",
            self.case_root or "",
            "Motion files (*.pdf *.docx)",
        )
        if motion_file:
            self.set_motion_file(motion_file)

    def set_motion_file(self, motion_file: str) -> bool:
        if not is_supported_motion_file(motion_file):
            QMessageBox.warning(self, "Unsupported motion file", "Select a PDF or DOCX motion.")
            return False
        self.motion_file = motion_file
        self._refresh_motion_file_label()
        return True

    def add_motion_files(self, paths: list[str]) -> None:
        for path in paths:
            if self.set_motion_file(path):
                return

    def _on_add_files(self) -> None:
        title = (
            "Select context document(s)"
            if self.current_draft_kind() == DRAFT_KIND_OPPOSITION
            else "Select context document(s) for the draft"
        )
        picked = ContextFilesDialog.get_files(
            self,
            title=title,
            start_dir=self.case_root or "",
            file_filter="Documents (*.pdf *.docx *.txt *.msg);;All files (*.*)",
        )
        self.add_target_files(picked or [])

    def intake_settings(self) -> dict:
        settings = super().intake_settings()
        settings.update(
            {
                "draft_kind": self.current_draft_kind(),
                "motion_type_label": self.selected_motion_type_label(),
                "motion_type_source_path": self.current_motion_type_source_path(),
                "motion_file": self.motion_file,
                "context_files": self.current_target_files(),
            }
        )
        return settings

    def to_dict(self) -> dict:
        data = super().to_dict()
        data.update(
            {
                "draft_kind": self.current_draft_kind(),
                "motion_type_label": self.selected_motion_type_label(),
                "motion_type_source_path": self.current_motion_type_source_path(),
                "motion_file": self.motion_file,
                "context_files": self.current_target_files(),
            }
        )
        return data

    def from_dict(self, data: dict | None) -> None:
        if not isinstance(data, dict):
            data = {}
        draft_kind = data.get("draft_kind") or DRAFT_KIND_MOTION
        idx = self.draft_kind_combo.findData(draft_kind)
        if idx >= 0:
            self.draft_kind_combo.setCurrentIndex(idx)
        self._populate_motion_type_combo()
        type_label = data.get("motion_type_label") or data.get("motion_type_name") or ""
        type_idx = self._find_motion_type_label(type_label)
        if type_idx >= 0:
            self.motion_type_combo.setCurrentIndex(type_idx)
        self.motion_file = data.get("motion_file", self.motion_file)
        self._refresh_motion_file_controls()

        self.files_list.clear()
        for path in data.get("target_files", data.get("context_files", [])) or []:
            self.files_list.addItem(path)

        self.set_metadata(MotionMetadata.from_dict(data.get("metadata")))
        self.set_outline(
            [
                item if isinstance(item, OutlineNode) else OutlineNode.from_dict(item)
                for item in data.get("outline", [])
                if isinstance(item, (OutlineNode, dict))
            ]
        )

    def set_metadata(self, metadata: MotionMetadata) -> None:
        if self.selected_motion_type_label():
            metadata = MotionMetadata.from_dict(metadata.to_dict())
            metadata.motion_type = self.current_motion_type_name()
        super().set_metadata(metadata)

    def _on_analyze_continue(self) -> None:
        settings = self.intake_settings()
        if self.current_draft_kind() == DRAFT_KIND_OPPOSITION and not self.motion_file:
            QMessageBox.warning(self, "Select a motion", "Select the motion to oppose before analyzing.")
            return
        if (
            self.current_draft_kind() != DRAFT_KIND_OPPOSITION
            and not settings["target_files"]
            and not settings["user_arguments"]
        ):
            QMessageBox.warning(
                self,
                "Add documents or arguments",
                "Add at least one context document, or type at least one argument.",
            )
            return
        self.analyze_requested.emit(settings)


class MotionDraftingOutputPage(GenerateMotionOutputPage):
    default_title = "Motion Draft"

    def empty_citations_message(self) -> str:
        return (
            "No citations were detected in this draft. Review the memorandum "
            "for any legal support that may need strengthening."
        )


class MotionDraftingTaskTab(GenerateMotionTaskTab):
    """Task tab that dispatches to the existing motion/opposition workers."""

    def __init__(self, spec, case_path: str, file_number: str, parent: QWidget | None = None):
        WizardTaskContainer.__init__(self, spec, parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._worker = None
        self._last_settings: dict = {}
        self._finishing_worker = None
        self._analysis_worker = None
        self._finishing_analysis_worker = None
        self._debug_run_id = None

        self.settings_page = MotionDraftingSettingsPage(case_path, file_number)
        self.status_page = StatusPage()
        self.output_page = MotionDraftingOutputPage()

        self.addWidget(self.settings_page)
        self.addWidget(self.status_page)
        self.addWidget(self.output_page)
        self.settings_page.analyze_requested.connect(self._start_analysis)
        self.settings_page.run_requested.connect(self._on_run)
        self.status_page.cancel_requested.connect(self._on_cancel)

    @property
    def files(self) -> list[str]:
        return self._current_files(self.settings_page.to_dict())

    def _current_files(self, settings: dict) -> list[str]:
        if settings.get("draft_kind") == DRAFT_KIND_OPPOSITION:
            return [
                path
                for path in [settings.get("motion_file", "")]
                + list(settings.get("context_files", []) or [])
                if path
            ]
        return list(settings.get("target_files", []) or [])

    def _start_analysis(self, intake_settings: dict) -> None:
        if (intake_settings or {}).get("draft_kind") != DRAFT_KIND_OPPOSITION:
            return super()._start_analysis(intake_settings)
        if self._analysis_worker is not None or self._finishing_analysis_worker is not None:
            return
        if not intake_settings.get("motion_file"):
            QMessageBox.warning(self, "Select a motion", "Select the motion to oppose before analyzing.")
            return
        self.status_page.reset()
        start_debug_run(
            self,
            source="wizard.motion_drafting.analysis",
            details={
                "case_path": self._case_path,
                "file_number": self._file_number,
                "settings": dict(intake_settings or {}),
            },
        )
        self.status_page.on_status("Analyzing selected motion and context...")
        record_status(self, "Analyzing selected motion and context...", source="wizard.ui")
        self.status_page.progress_bar.setRange(0, 0)
        self.setCurrentIndex(TASK_PAGE_STATUS)
        worker = OpposeMotionAnalysisWorker(
            settings={
                "motion_file": intake_settings.get("motion_file", ""),
                "context_files": list(intake_settings.get("context_files", []) or []),
                "case_root": self._case_path,
            },
            parent=None,
        )
        worker.progress.connect(
            lambda message: self._on_worker_progress(
                message,
                source="wizard.motion_drafting.analysis",
            )
        )
        worker.finished_analysis.connect(self._on_analysis_finished)
        worker.finished.connect(lambda w=worker: self._on_analysis_thread_finished(w))
        worker.finished.connect(worker.deleteLater)
        self._analysis_worker = worker
        worker.start()

    def _on_run(self, settings: dict) -> None:
        if (settings or {}).get("draft_kind") != DRAFT_KIND_OPPOSITION:
            return super()._on_run(settings)
        if self._analysis_worker is not None or self._finishing_analysis_worker is not None:
            self.status_page.on_status("Motion analysis is still running.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        if self._worker is not None or self._finishing_worker is not None:
            self.status_page.on_status("Motion draft is already running.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        self.status_page.reset()
        self._last_settings = dict(settings or {})
        start_debug_run(
            self,
            source="wizard.motion_drafting.draft",
            details={
                "case_path": self._case_path,
                "file_number": self._file_number,
                "files": self._current_files(self._last_settings),
                "settings": dict(self._last_settings),
            },
        )
        self.status_page.on_status("Drafting opposition memorandum...")
        record_status(self, "Drafting opposition memorandum...", source="wizard.ui")
        self.status_page.progress_bar.setRange(0, 0)
        self.setCurrentIndex(TASK_PAGE_STATUS)
        worker = OpposeMotionWorker(
            case_path=self._case_path,
            file_number=self._file_number,
            settings=self._last_settings,
            parent=None,
        )
        worker.progress.connect(
            lambda message: self._on_worker_progress(
                message,
                source="wizard.motion_drafting.draft",
            )
        )
        worker.finished_result.connect(self._on_worker_finished)
        worker.finished.connect(lambda w=worker: self._on_worker_thread_finished(w))
        worker.finished.connect(worker.deleteLater)
        self._worker = worker
        worker.start()

    def _on_worker_finished(self, success: bool, payload: object) -> None:
        if self.sender() is not None and self.sender() is not self._worker:
            return
        if self.sender() is self._worker:
            self._finishing_worker = self._worker
            self._worker = None
        elif self._worker is not None:
            self._worker = None
        if not success:
            finish_debug_run(
                self,
                status="error",
                message=f"Draft failed: {payload}",
                details={"error": str(payload)},
            )
            self.status_page.on_status(f"FAILED: {payload}")
            return

        draft = payload if isinstance(payload, DraftDocument) else DraftDocument()
        finish_debug_run(
            self,
            status="success",
            message="Task complete",
            details={
                "output_path": draft.preview_path,
                "citation_count": len(draft.citations or []),
                "diagnostics": dict(draft.diagnostics or {}),
            },
        )
        self.output_page.show_result(draft)
        self.setCurrentIndex(TASK_PAGE_OUTPUT)
        settings = self.settings_page.to_dict()
        self.task_completed.emit(
            {
                "task_id": self._spec.task_id,
                "title": self._spec.title,
                "files": self._current_files(settings),
                "settings": settings,
                "output_path": draft.preview_path,
                "completed_at": datetime.now().isoformat(timespec="seconds"),
            }
        )


def build_motion_drafting_tab(spec, case_path: str, file_number: str, parent: QWidget | None):
    return MotionDraftingTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
