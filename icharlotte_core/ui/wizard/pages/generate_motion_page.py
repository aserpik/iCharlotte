"""Custom Wizard task page for generating motions from scratch.

Flow: an intake settings page collects the motion type (one of the configured
types or a custom/"Other" type), context files, and the user's own
arguments/relief. On Analyze & Continue, an LLM proposes additional grounds from
the documents; these are merged with the user's typed grounds and shown in an
editable review step, then an outline, then drafting. Reuses the opposition
research/verify spine and the motion_generation drafter/assembler.
"""
from __future__ import annotations

import os

from PySide6.QtCore import QThread, QUrl, Qt, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QStackedWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.discovery.assembler import DiscoveryAssembler
from icharlotte_core.opposition.argument_research import research_arguments
from icharlotte_core.opposition.citation_parser import extract_citations
from icharlotte_core.opposition.extraction import (
    extract_context_bundle,
    is_supported_context_file,
)
from icharlotte_core.opposition.models import (
    DraftDocument,
    MotionMetadata,
    OutlineNode,
)
from icharlotte_core.opposition.outline import selected_section_plan
from icharlotte_core.opposition.verifier import (
    build_local_opposition_verifier,
    build_opposition_verifier,
    enrich_with_pool_signals,
    pool_membership_check,
)
from icharlotte_core.ui.wizard.pages.citation_review import CitationReviewOutputPage
from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
    _firm_style_exemplars,
    _make_firm_provider,
    _make_local_corpus,
    _research_targets,
)
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer
from icharlotte_core.ui.context_files_dialog import ContextFilesDialog

from icharlotte_core.motion_generation.analyzer import (
    analyze_target,
    generate_motion_outline,
    merge_intake_with_analysis,
    outline_from_config,
)
from icharlotte_core.motion_generation.assembler import assemble_motion_preview
from icharlotte_core.motion_generation.config import (
    get_motion_config,
    list_motion_types,
)
from icharlotte_core.motion_generation.drafter import draft_motion
from icharlotte_core.firm_briefs.motion_taxonomy import normalize_motion_type


SETTINGS_PAGE_INTAKE = 0
SETTINGS_PAGE_REVIEW = 1
SETTINGS_PAGE_OUTLINE = 2
TASK_PAGE_SETTINGS = 0
TASK_PAGE_STATUS = 1
TASK_PAGE_OUTPUT = 2

# Sentinel combo entry that enables a custom motion-type name.
_OTHER = "__other__"


class GenerateMotionSettingsPage(QStackedWidget):
    """Intake → editable review → outline screens for motion generation."""

    analyze_requested = Signal(dict)
    run_requested = Signal(dict)

    def __init__(self, case_root: str, file_number: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.case_root = case_root
        self.file_number = file_number
        self.metadata = MotionMetadata()
        self.outline: list[OutlineNode] = []
        self._build_intake_page()
        self._build_review_page()
        self._build_outline_page()
        self._on_type_changed()

    # ---- Intake page ---------------------------------------------------- #

    def _build_intake_page(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Motion type"))
        self.type_combo = QComboBox()
        for cfg in list_motion_types():
            if cfg.type_id == "generic":
                continue  # the generic engine is reached via "Other (specify…)"
            self.type_combo.addItem(cfg.display_name, cfg.type_id)
        self.type_combo.addItem("Other (specify…)", _OTHER)
        self.type_combo.currentIndexChanged.connect(self._on_type_changed)
        layout.addWidget(self.type_combo)

        self.custom_name_edit = QLineEdit()
        self.custom_name_edit.setPlaceholderText("Name the motion (e.g. Motion for Protective Order)")
        layout.addWidget(self.custom_name_edit)

        self.guidance_label = QLabel()
        self.guidance_label.setWordWrap(True)
        self.guidance_label.setStyleSheet("color: #6b7480;")
        layout.addWidget(self.guidance_label)

        files_row = QHBoxLayout()
        files_row.addWidget(QLabel("Context documents"))
        files_row.addStretch()
        self.add_files_btn = QPushButton("Add Files…")
        self.add_files_btn.clicked.connect(self._on_add_files)
        files_row.addWidget(self.add_files_btn)
        self.remove_files_btn = QPushButton("Remove")
        self.remove_files_btn.clicked.connect(self._on_remove_files)
        files_row.addWidget(self.remove_files_btn)
        layout.addLayout(files_row)

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(110)
        layout.addWidget(self.files_list)

        layout.addWidget(QLabel("Relief requested"))
        self.user_relief_edit = QLineEdit()
        self.user_relief_edit.setPlaceholderText("What you are asking the court to order")
        layout.addWidget(self.user_relief_edit)

        layout.addWidget(QLabel("Arguments / grounds to include (one per line)"))
        self.user_arguments_edit = QPlainTextEdit()
        self.user_arguments_edit.setPlaceholderText(
            "Type the specific arguments or grounds you want in the motion. The AI "
            "will analyze your documents and propose additional grounds to merge."
        )
        layout.addWidget(self.user_arguments_edit, 1)

        row = QHBoxLayout()
        row.addStretch()
        self.analyze_btn = QPushButton("Analyze & Continue")
        self.analyze_btn.clicked.connect(self._on_analyze_continue)
        row.addWidget(self.analyze_btn)
        layout.addLayout(row)
        self.addWidget(page)

    def _build_review_page(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(10)

        self.review_type_label = QLabel()
        layout.addWidget(self.review_type_label)

        layout.addWidget(QLabel("Relief requested (you can edit)"))
        self.relief_edit = QLineEdit()
        layout.addWidget(self.relief_edit)

        layout.addWidget(QLabel("Grounds (merged; edit or add your own, one per line)"))
        self.arguments_edit = QPlainTextEdit()
        layout.addWidget(self.arguments_edit, 1)

        row = QHBoxLayout()
        self.back_btn = QPushButton("Back")
        self.back_btn.clicked.connect(lambda: self.setCurrentIndex(SETTINGS_PAGE_INTAKE))
        row.addWidget(self.back_btn)
        row.addStretch()
        self.continue_btn = QPushButton("Review Outline")
        self.continue_btn.clicked.connect(self._on_continue_to_outline)
        row.addWidget(self.continue_btn)
        layout.addLayout(row)
        self.addWidget(page)

    def _build_outline_page(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(10)

        layout.addWidget(QLabel("Motion Outline"))
        self.outline_tree = QTreeWidget()
        self.outline_tree.setHeaderLabels(["Include / Heading"])
        layout.addWidget(self.outline_tree, 1)

        row = QHBoxLayout()
        self.add_heading_btn = QPushButton("Add Heading")
        self.add_heading_btn.clicked.connect(self._add_heading)
        row.addWidget(self.add_heading_btn)
        row.addStretch()
        self.generate_btn = QPushButton("Generate Motion")
        self.generate_btn.clicked.connect(self._emit_run_requested)
        row.addWidget(self.generate_btn)
        layout.addLayout(row)
        self.addWidget(page)

    # ---- Type / intake helpers ----------------------------------------- #

    def current_motion_type_id(self) -> str:
        data = self.type_combo.currentData()
        return "generic" if data == _OTHER else (data or "generic")

    def current_motion_type_name(self) -> str:
        if self.type_combo.currentData() == _OTHER:
            return self.custom_name_edit.text().strip() or "Motion"
        return get_motion_config(self.current_motion_type_id()).display_name

    def current_target_files(self) -> list[str]:
        return [self.files_list.item(i).text() for i in range(self.files_list.count())]

    def _on_type_changed(self, *_args) -> None:
        is_other = self.type_combo.currentData() == _OTHER
        self.custom_name_edit.setVisible(is_other)
        cfg = get_motion_config(self.current_motion_type_id())
        self.guidance_label.setText(cfg.target_doc_guidance)

    def _on_add_files(self) -> None:
        picked = ContextFilesDialog.get_files(
            self,
            title="Select context document(s) for the motion",
            start_dir=self.case_root or "",
            file_filter="Documents (*.pdf *.docx *.txt *.msg);;All files (*.*)",
        )
        existing = set(self.current_target_files())
        for path in picked or []:
            if is_supported_context_file(path) and path not in existing:
                self.files_list.addItem(path)
                existing.add(path)

    def _on_remove_files(self) -> None:
        for item in self.files_list.selectedItems():
            self.files_list.takeItem(self.files_list.row(item))

    def intake_settings(self) -> dict:
        return {
            "motion_type_id": self.current_motion_type_id(),
            "motion_type_name": self.current_motion_type_name(),
            "target_files": self.current_target_files(),
            "user_relief": self.user_relief_edit.text().strip(),
            "user_arguments": [
                line.strip()
                for line in self.user_arguments_edit.toPlainText().splitlines()
                if line.strip()
            ],
        }

    def _on_analyze_continue(self) -> None:
        if self.type_combo.currentData() == _OTHER and not self.custom_name_edit.text().strip():
            QMessageBox.warning(self, "Name the motion", "Enter a name for the motion type.")
            return
        settings = self.intake_settings()
        if not settings["target_files"] and not settings["user_arguments"]:
            QMessageBox.warning(
                self,
                "Add documents or arguments",
                "Add at least one context document, or type at least one argument.",
            )
            return
        self.analyze_requested.emit(settings)

    # ---- Review / outline ---------------------------------------------- #

    def set_metadata(self, metadata: MotionMetadata) -> None:
        self.metadata = metadata
        self.review_type_label.setText(f"<b>{metadata.motion_type or 'Motion'}</b>")
        self.relief_edit.setText(metadata.relief_requested)
        self.arguments_edit.setPlainText("\n".join(metadata.principal_arguments))

    def current_metadata(self) -> MotionMetadata:
        metadata = MotionMetadata.from_dict(self.metadata.to_dict())
        metadata.motion_type = self.current_motion_type_name()
        metadata.relief_requested = self.relief_edit.text().strip()
        metadata.principal_arguments = [
            line.strip()
            for line in self.arguments_edit.toPlainText().splitlines()
            if line.strip()
        ]
        return metadata

    def can_continue_to_outline(self) -> bool:
        return self.current_metadata().required_missing() == []

    def set_outline(self, outline: list[OutlineNode]) -> None:
        self.outline = list(outline)
        self.outline_tree.clear()
        for node in self.outline:
            self.outline_tree.addTopLevelItem(self._item_from_node(node))
        self.outline_tree.expandAll()

    def to_dict(self) -> dict:
        return {
            "motion_type_id": self.current_motion_type_id(),
            "motion_type_name": self.current_motion_type_name(),
            "target_files": self.current_target_files(),
            "metadata": self.current_metadata().to_dict(),
            "outline": [node.to_dict() for node in self._outline_from_tree()],
        }

    def from_dict(self, data: dict | None) -> None:
        if not isinstance(data, dict):
            data = {}
        type_id = data.get("motion_type_id", "")
        type_name = data.get("motion_type_name", "")
        combo_idx = self.type_combo.findData(type_id) if type_id and type_id != "generic" else -1
        if combo_idx >= 0:
            self.type_combo.setCurrentIndex(combo_idx)
        else:
            other_idx = self.type_combo.findData(_OTHER)
            if other_idx >= 0:
                self.type_combo.setCurrentIndex(other_idx)
            self.custom_name_edit.setText(type_name)
        self._on_type_changed()

        self.files_list.clear()
        for path in data.get("target_files", []) or []:
            self.files_list.addItem(path)

        self.set_metadata(MotionMetadata.from_dict(data.get("metadata")))
        self.set_outline(
            [
                OutlineNode.from_dict(item)
                for item in data.get("outline", [])
                if isinstance(item, dict)
            ]
        )

    def _on_continue_to_outline(self) -> None:
        if not self.can_continue_to_outline():
            QMessageBox.warning(
                self,
                "Missing required fields",
                "Relief requested and at least one ground are required.",
            )
            return
        self.setCurrentIndex(SETTINGS_PAGE_OUTLINE)

    def _emit_run_requested(self) -> None:
        self.run_requested.emit(self.to_dict())

    def _add_heading(self) -> None:
        item = QTreeWidgetItem(["New heading"])
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(0, Qt.CheckState.Checked)
        item.setData(0, Qt.ItemDataRole.UserRole, "")
        self.outline_tree.addTopLevelItem(item)
        self.outline_tree.editItem(item, 0)

    def _item_from_node(self, node: OutlineNode) -> QTreeWidgetItem:
        item = QTreeWidgetItem([node.text])
        item.setData(0, Qt.ItemDataRole.UserRole, node.id)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(0, Qt.CheckState.Checked if node.selected else Qt.CheckState.Unchecked)
        for child in node.children:
            item.addChild(self._item_from_node(child))
        return item

    def _outline_from_tree(self) -> list[OutlineNode]:
        return [
            self._node_from_item(self.outline_tree.topLevelItem(index))
            for index in range(self.outline_tree.topLevelItemCount())
        ]

    def _node_from_item(self, item: QTreeWidgetItem) -> OutlineNode:
        return OutlineNode(
            id=item.data(0, Qt.ItemDataRole.UserRole) or "",
            text=item.text(0).strip(),
            selected=item.checkState(0) == Qt.CheckState.Checked,
            children=[
                self._node_from_item(item.child(index))
                for index in range(item.childCount())
            ],
        )


def _make_llms():
    """Return (analysis_llm, draft_llm, make_pass_llm) reusing the oppose_motion
    agent config so model selection is shared."""
    from icharlotte_core.llm_config import call_llm

    def analysis_llm(system_prompt, user_prompt):
        return call_llm(
            user_prompt, system_prompt, task_type="general", agent_id="agent_chat"
        ) or ""

    def draft_llm(system_prompt, user_prompt):
        return call_llm(
            user_prompt, system_prompt, task_type="general",
            agent_id="agent_oppose_motion",
        ) or ""

    def make_pass_llm(pass_name):
        def _llm(system_prompt, user_prompt):
            return call_llm(
                user_prompt, system_prompt, task_type="general",
                agent_id="agent_oppose_motion", pass_name=pass_name,
                pass_agent_id="agent_oppose_motion",
            ) or ""
        return _llm

    return analysis_llm, draft_llm, make_pass_llm


class GenerateMotionAnalysisWorker(QThread):
    progress = Signal(str)
    finished_analysis = Signal(bool, object)

    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.settings = dict(settings or {})

    def run(self) -> None:
        try:
            config = get_motion_config(self.settings.get("motion_type_id"))
            name = self.settings.get("motion_type_name") or config.display_name
            user_relief = self.settings.get("user_relief", "")
            user_arguments = list(self.settings.get("user_arguments", []))

            self.progress.emit("Extracting context documents...")
            target_text, warnings = extract_context_bundle(
                self.settings.get("target_files", [])
            )
            for warning in warnings:
                self.progress.emit(f"WARNING: {warning}")

            if not target_text.strip() and not any(a.strip() for a in user_arguments):
                self.finished_analysis.emit(
                    False, "Add at least one document or type at least one argument."
                )
                return

            ai_metadata = MotionMetadata(motion_type=name)
            # Build the analysis LLM up front: it is also used to generate the
            # outline below even when no target documents were supplied.
            analysis_llm, _, _ = _make_llms()
            if target_text.strip():
                self.progress.emit("Proposing additional grounds from documents...")
                ai_metadata = analyze_target(
                    config, target_text, llm_callback=analysis_llm, motion_name=name
                )

            merged = merge_intake_with_analysis(user_relief, user_arguments, ai_metadata, name)
            self.progress.emit("Building a detailed outline for the motion...")
            outline = generate_motion_outline(
                config, merged, context_text="", target_text=target_text,
                llm_callback=analysis_llm,
            )
            self.finished_analysis.emit(
                True,
                {"metadata": merged, "outline": outline, "target_text": target_text},
            )
        except Exception as exc:  # noqa: BLE001
            self.finished_analysis.emit(False, str(exc))


class GenerateMotionWorker(QThread):
    progress = Signal(str)
    finished_result = Signal(bool, object)

    def __init__(self, case_path: str, file_number: str, settings: dict, parent=None):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.settings = dict(settings or {})

    def run(self) -> None:
        try:
            config = get_motion_config(self.settings.get("motion_type_id"))
            self.progress.emit("Extracting context documents...")
            target_text, warnings = extract_context_bundle(
                self.settings.get("target_files", [])
            )
            for warning in warnings:
                self.progress.emit(f"WARNING: {warning}")

            metadata = MotionMetadata.from_dict(self.settings.get("metadata"))
            outline = [
                OutlineNode.from_dict(item)
                for item in self.settings.get("outline", [])
                if isinstance(item, dict)
            ]
            plan = selected_section_plan(outline)

            _, draft_llm, make_pass_llm = _make_llms()

            research_targets = _research_targets(metadata, plan)
            corpus = _make_local_corpus()
            token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()
            retrieved = []
            # Cache opinion text under this task's prompt dir (mirrors oppose).
            repo_root = os.path.dirname(
                os.path.dirname(
                    os.path.dirname(
                        os.path.dirname(os.path.dirname(__file__))
                    )
                )
            )
            opinion_cache = os.path.join(
                repo_root, "Scripts", "prompts", "generate_motion", ".cache", "opinions"
            )
            raw_type = self.settings.get("motion_type_id") or getattr(metadata, "motion_type", "")
            firm_motion_type = raw_type if raw_type not in ("", "generic") else \
                normalize_motion_type(self.settings.get("motion_type_name", "") or getattr(metadata, "motion_type", ""))
            if research_targets and (corpus is not None or token):
                client = corpus
                if client is None:
                    from icharlotte_core.legal_research.sources.courtlistener import (
                        CourtListenerClient,
                    )
                    client = CourtListenerClient(token)
                    self.progress.emit(
                        "Local corpus not built; using CourtListener API "
                        f"({len(research_targets)} points)..."
                    )
                    firm_provider = None
                else:
                    self.progress.emit(
                        f"Researching authorities locally ({len(research_targets)} points)..."
                    )
                    firm_provider = _make_firm_provider(corpus)
                    if firm_provider is not None:
                        self.progress.emit("  Firm brief library active (preferring your prior authorities).")
                retrieved = research_arguments(
                    research_targets,
                    cl_client=client,
                    query_llm=make_pass_llm("research_queries"),
                    rerank_llm=make_pass_llm("rerank_select"),
                    # Keep concurrency low: 4 parallel workers burst the LLM /
                    # CourtListener rate limit on the per-point query-gen +
                    # rerank calls (oppose's hard-won lesson).
                    max_workers=2,
                    on_progress=self.progress.emit,
                    cache_dir=opinion_cache,
                    firm_provider=firm_provider,
                    motion_type=firm_motion_type,
                    side="moving",
                )
                self.progress.emit(f"Retrieved {len(retrieved)} grounded authorities.")
            else:
                self.progress.emit(
                    "WARNING: no grounded research available; drafting from statutes only."
                )

            from icharlotte_core.motion_generation.samples import load_exemplars

            exemplars = load_exemplars(self.settings.get("motion_type_id") or "")
            if exemplars:
                self.progress.emit(f"Using {len(exemplars)} style sample(s) for this motion type.")
            firm_style = _firm_style_exemplars(firm_motion_type, "moving", metadata)
            if firm_style:
                self.progress.emit(f"Using {len(firm_style)} firm-library style sample(s).")
            exemplars = (firm_style + exemplars)[:3]

            self.progress.emit("Drafting motion memorandum...")
            draft = draft_motion(
                config,
                metadata,
                plan,
                target_text,
                "",
                style_exemplars=exemplars,
                retrieved_authorities=retrieved,
                llm_callback=draft_llm,
            )
            if not draft.body_text.strip():
                reason = (draft.rejection_reason or "unknown reason").strip()
                self.finished_result.emit(False, f"Drafting failed: {reason}")
                return

            citations = extract_citations(draft.body_text)
            if citations:
                to_verify, off_pool = pool_membership_check(citations, retrieved)
                self.progress.emit(f"Verifying citations ({len(to_verify)} found)...")
                if corpus is not None:
                    verifier = build_local_opposition_verifier(
                        corpus=corpus, llm_callback=draft_llm, max_workers=4
                    )
                    verified = verifier.verify_all(to_verify, on_progress=self.progress.emit)
                elif token:
                    verifier = build_opposition_verifier(
                        courtlistener_token=token, llm_callback=draft_llm, max_workers=4
                    )
                    verified = verifier.verify_all(to_verify, on_progress=self.progress.emit)
                else:
                    verified = []
                draft.citations = sorted(
                    list(verified) + list(off_pool),
                    key=lambda cv: cv.body_offset if cv.body_offset is not None else 0,
                )
                enrich_with_pool_signals(draft.citations, retrieved)
            else:
                self.progress.emit("WARNING: No citations detected in the draft.")
                draft.citations = []

            preview_dir = os.path.join(
                self.case_path, "NOTES", "AI OUTPUT", ".icharlotte",
                "wizard_previews", "generate_motion",
            )
            preview_path = os.path.join(preview_dir, "Motion Preview.docx")
            caption_path = DiscoveryAssembler.find_caption_page(self.case_path) or ""
            assemble_motion_preview(
                draft=draft, output_path=preview_path, config=config,
                caption_path=caption_path,
            )
            draft.preview_path = preview_path
            self.finished_result.emit(True, draft)
        except Exception as exc:  # noqa: BLE001
            self.finished_result.emit(False, str(exc))


class GenerateMotionOutputPage(CitationReviewOutputPage):
    default_title = "Generated Motion"

    def empty_citations_message(self) -> str:
        return (
            "No citations were detected in this motion. If California "
            "case-law research returned no results, the motion was drafted "
            "without case citations. Review for any statutory support that "
            "may need strengthening."
        )

    def _build_action_buttons(self, row) -> None:
        self._add_save_button(row)
        self.open_btn = QPushButton("Open in Word")
        self.open_btn.clicked.connect(self._open_preview)
        self.open_btn.setEnabled(False)
        row.addWidget(self.open_btn)

    def show_result(self, draft: DraftDocument) -> None:
        super().show_result(draft)
        self._refresh_open_btn()

    def load_output(self, output_path: str) -> None:
        super().load_output(output_path)
        self._refresh_open_btn()

    def _refresh_open_btn(self) -> None:
        path = self.output_path
        self.open_btn.setEnabled(bool(path) and os.path.isfile(path))

    def _open_preview(self) -> None:
        path = self.output_path
        if path and os.path.isfile(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))


class GenerateMotionTaskTab(WizardTaskContainer):
    task_completed = Signal(dict)

    def __init__(self, spec, case_path: str, file_number: str, parent: QWidget | None = None):
        super().__init__(spec, parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._worker = None
        self._last_settings: dict = {}
        self._finishing_worker = None
        self._analysis_worker = None
        self._finishing_analysis_worker = None

        self.settings_page = GenerateMotionSettingsPage(case_path, file_number)
        self.status_page = StatusPage()
        self.output_page = GenerateMotionOutputPage()

        self.addWidget(self.settings_page)
        self.addWidget(self.status_page)
        self.addWidget(self.output_page)
        self.settings_page.analyze_requested.connect(self._start_analysis)
        self.settings_page.run_requested.connect(self._on_run)

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list[str]:
        return list(self.settings_page.current_target_files())

    def _start_analysis(self, intake_settings: dict) -> None:
        if self._analysis_worker is not None or self._finishing_analysis_worker is not None:
            return
        self.status_page.reset()
        self.status_page.on_status("Analyzing context documents...")
        self.status_page.progress_bar.setRange(0, 0)
        self.setCurrentIndex(TASK_PAGE_STATUS)
        worker = GenerateMotionAnalysisWorker(settings=dict(intake_settings or {}), parent=None)
        worker.progress.connect(self.status_page.on_status)
        worker.finished_analysis.connect(self._on_analysis_finished)
        worker.finished.connect(lambda w=worker: self._on_analysis_thread_finished(w))
        worker.finished.connect(worker.deleteLater)
        self._analysis_worker = worker
        worker.start()

    def _on_analysis_finished(self, success: bool, payload: object) -> None:
        if self.sender() is not None and self.sender() is not self._analysis_worker:
            return
        if self.sender() is self._analysis_worker:
            self._finishing_analysis_worker = self._analysis_worker
            self._analysis_worker = None
        elif self._analysis_worker is not None:
            self._analysis_worker = None

        if not success:
            self.status_page.on_status(f"FAILED: {payload}")
            return

        data = payload if isinstance(payload, dict) else {}
        metadata = data.get("metadata")
        if not isinstance(metadata, MotionMetadata):
            metadata = MotionMetadata.from_dict(metadata if isinstance(metadata, dict) else {})
        outline = [
            node if isinstance(node, OutlineNode) else OutlineNode.from_dict(node)
            for node in data.get("outline", []) or []
            if isinstance(node, (OutlineNode, dict))
        ]
        self.settings_page.set_metadata(metadata)
        self.settings_page.set_outline(outline)
        self.settings_page.setCurrentIndex(SETTINGS_PAGE_REVIEW)
        self.setCurrentIndex(TASK_PAGE_SETTINGS)

    def _on_analysis_thread_finished(self, worker) -> None:
        if self._analysis_worker is worker:
            self._analysis_worker = None
        if self._finishing_analysis_worker is worker:
            self._finishing_analysis_worker = None

    def _on_run(self, settings: dict) -> None:
        if self._analysis_worker is not None or self._finishing_analysis_worker is not None:
            self.status_page.on_status("Analysis is still running.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        if self._worker is not None or self._finishing_worker is not None:
            self.status_page.on_status("Motion draft is already running.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        self.status_page.reset()
        self.status_page.on_status("Drafting motion...")
        self.status_page.progress_bar.setRange(0, 0)
        self.setCurrentIndex(TASK_PAGE_STATUS)
        self._last_settings = dict(settings or {})
        worker = GenerateMotionWorker(
            case_path=self._case_path,
            file_number=self._file_number,
            settings=self._last_settings,
            parent=None,
        )
        worker.progress.connect(self.status_page.on_status)
        worker.finished_result.connect(self._on_worker_finished)
        worker.finished.connect(lambda w=worker: self._on_worker_thread_finished(w))
        worker.finished.connect(worker.deleteLater)
        self._worker = worker
        worker.start()

    def _on_worker_finished(self, success: bool, payload: object) -> None:
        from datetime import datetime

        if self.sender() is not None and self.sender() is not self._worker:
            return
        if self.sender() is self._worker:
            self._finishing_worker = self._worker
            self._worker = None
        elif self._worker is not None:
            self._worker = None
        if not success:
            self.status_page.on_status(f"FAILED: {payload}")
            return

        draft = payload if isinstance(payload, DraftDocument) else DraftDocument()
        self.output_page.show_result(draft)
        self.setCurrentIndex(TASK_PAGE_OUTPUT)
        self.task_completed.emit(
            {
                "task_id": self._spec.task_id,
                "title": self._spec.title,
                "files": list(self.settings_page.current_target_files()),
                "settings": self.settings_page.to_dict(),
                "output_path": draft.preview_path,
                "completed_at": datetime.now().isoformat(timespec="seconds"),
            }
        )

    def _on_worker_thread_finished(self, worker) -> None:
        if self._worker is worker:
            self._worker = None
        if self._finishing_worker is worker:
            self._finishing_worker = None

    def closeEvent(self, event) -> None:
        for worker in (
            self._analysis_worker,
            self._finishing_analysis_worker,
            self._worker,
            self._finishing_worker,
        ):
            if worker is not None and worker.isRunning():
                QMessageBox.information(
                    self,
                    "Task running",
                    "The motion draft is still running. Wait for it to finish before closing this tab.",
                )
                event.ignore()
                return
        super().closeEvent(event)


def build_generate_motion_tab(spec, case_path: str, file_number: str, parent: QWidget | None):
    """Open the Generate Motion task directly on its intake settings page."""
    return GenerateMotionTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
