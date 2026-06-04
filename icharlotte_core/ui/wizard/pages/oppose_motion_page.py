"""Custom Wizard task page for drafting oppositions to motions."""

from __future__ import annotations

import os
import re

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
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
from icharlotte_core.opposition.assembler import assemble_opposition_preview
from icharlotte_core.opposition.citation_parser import extract_citations
from icharlotte_core.opposition.drafter import draft_memorandum
from icharlotte_core.opposition.extraction import (
    extract_context_bundle,
    extract_document_text,
    is_supported_context_file,
    is_supported_motion_file,
)
from icharlotte_core.opposition.models import DraftDocument, MotionMetadata, OutlineNode
from icharlotte_core.opposition.motion_analyzer import analyze_motion, generate_outline
from icharlotte_core.opposition.outline import selected_section_plan
from icharlotte_core.opposition.style_examples import (
    StyleExampleRegistry,
    extract_exemplar_text,
)
from icharlotte_core.opposition.verifier import (
    build_local_opposition_verifier,
    build_opposition_verifier,
    enrich_with_pool_signals,
    pool_membership_check,
)
import os as _os_corpus
from icharlotte_core.config import CASELAW_DATA_DIR


def _corpus_paths() -> tuple[str, str]:
    return (_os_corpus.path.join(CASELAW_DATA_DIR, "corpus.db"),
            _os_corpus.path.join(CASELAW_DATA_DIR, "vectors.f16"))


def _corpus_available() -> bool:
    # Require BOTH files: corpus.db is created (empty) at the start of a build,
    # but vectors.f16 is only written at finalize(). Requiring both means an
    # in-progress or partial build safely falls back to the live API instead of
    # using an empty corpus with a missing vector sidecar.
    db, vec = _corpus_paths()
    return _os_corpus.path.exists(db) and _os_corpus.path.exists(vec)


def _corpus_embedder():
    from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder
    return OnnxEmbedder()


def _research_targets(metadata, plan) -> list[str]:
    """Propositions to research, deduped: the union of the principal arguments
    and every selected section-plan leaf.

    The drafter expands the brief into one subsection per section-plan leaf, and
    each subsection makes its own legal proposition (meet-and-confer, discovery
    cutoff, cumulative discovery, ...). Researching only the top-level arguments
    left those sub-points ungrounded — the drafter then emitted "[no case
    authority retrieved for this point]". Researching each leaf gives every
    subsection its own on-point authority. Purely structural sections are
    skipped; the count is capped to bound LLM calls under provider rate limits.
    """
    targets: list[str] = []
    seen: set[str] = set()

    def _add(text: str) -> None:
        t = (text or "").strip()
        if not t:
            return
        key = re.sub(r"\s+", " ", t.lower())
        if key in seen:
            return
        seen.add(key)
        targets.append(t)

    for arg in (getattr(metadata, "principal_arguments", None) or []):
        _add(arg)
    # Structural sections that argue no legal point and need no case authority.
    _skip = ("introduction", "conclusion", "statement of facts",
             "factual background", "preliminary statement", "prayer")
    for item in (plan or []):
        text = (getattr(item, "text", "") or "").strip()
        if not text or any(s in text.lower() for s in _skip):
            continue
        _add(text)
    return targets[:24]


def _make_local_corpus():
    if not _corpus_available():
        return None
    from icharlotte_core.legal_research.local_corpus.corpus import LocalCaseCorpus
    db, vec = _corpus_paths()
    return LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=_corpus_embedder())
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer
from icharlotte_core.ui.context_files_dialog import ContextFilesDialog
from icharlotte_core.word_validator import validate_opposition_docx
from icharlotte_core.ui.wizard.pages.citation_review import (  # noqa: F401  (re-exported for tests/back-compat)
    CitationDetailDialog,
    CitationDetailPanel,
    CitationReviewOutputPage,
    _build_citation_index,
    _citation_body_html,
    _citation_header_html,
    _color_for_verdict,
    _format_inline_html,
    _render_draft_html,
    _run_find_replacement,
    _VERDICT_COLORS,
    _VERDICT_HEADER_COLORS,
    _VERDICT_LABELS,
)


SETTINGS_PAGE_CONFIRM = 0
SETTINGS_PAGE_OUTLINE = 1
TASK_PAGE_SETTINGS = 0
TASK_PAGE_STATUS = 1
TASK_PAGE_OUTPUT = 2


class OpposeMotionSettingsPage(QStackedWidget):
    """Confirmation and editable outline screens for the opposition workflow."""

    run_requested = Signal(dict)

    def __init__(
        self,
        case_root: str,
        file_number: str,
        motion_file: str,
        context_files: list[str],
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.case_root = case_root
        self.file_number = file_number
        self.motion_file = motion_file
        self.context_files = list(context_files)
        self.metadata = MotionMetadata()
        self.outline: list[OutlineNode] = []
        self._build_confirm_page()
        self._build_outline_page()

    def _build_confirm_page(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(10)

        self.motion_label = QLabel()
        layout.addWidget(self.motion_label)
        self._refresh_motion_label()

        self.motion_type_edit = QLineEdit()
        self.motion_type_edit.setPlaceholderText("Motion type")
        layout.addWidget(self.motion_type_edit)

        self.relief_edit = QLineEdit()
        self.relief_edit.setPlaceholderText("Relief requested")
        layout.addWidget(self.relief_edit)

        self.arguments_edit = QPlainTextEdit()
        self.arguments_edit.setPlaceholderText("Principal arguments, one per line")
        layout.addWidget(self.arguments_edit, 1)

        row = QHBoxLayout()
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

        layout.addWidget(QLabel("Opposition Outline"))
        self.outline_tree = QTreeWidget()
        self.outline_tree.setHeaderLabels(["Include / Heading"])
        layout.addWidget(self.outline_tree, 1)

        row = QHBoxLayout()
        self.add_heading_btn = QPushButton("Add Heading")
        self.add_heading_btn.clicked.connect(self._add_heading)
        row.addWidget(self.add_heading_btn)
        row.addStretch()
        self.generate_btn = QPushButton("Generate Draft")
        self.generate_btn.clicked.connect(self._emit_run_requested)
        row.addWidget(self.generate_btn)
        layout.addLayout(row)
        self.addWidget(page)

    def set_metadata(self, metadata: MotionMetadata) -> None:
        self.metadata = metadata
        self.motion_type_edit.setText(metadata.motion_type)
        self.relief_edit.setText(metadata.relief_requested)
        self.arguments_edit.setPlainText("\n".join(metadata.principal_arguments))

    def current_metadata(self) -> MotionMetadata:
        metadata = MotionMetadata.from_dict(self.metadata.to_dict())
        metadata.motion_type = self.motion_type_edit.text().strip()
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
            "motion_file": self.motion_file,
            "context_files": list(self.context_files),
            "metadata": self.current_metadata().to_dict(),
            "outline": [node.to_dict() for node in self._outline_from_tree()],
        }

    def from_dict(self, data: dict | None) -> None:
        if not isinstance(data, dict):
            data = {}
        self.motion_file = data.get("motion_file", self.motion_file)
        self.context_files = list(data.get("context_files", self.context_files))
        self._refresh_motion_label()
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
                "Motion type, relief requested, and principal arguments are required.",
            )
            return
        self.setCurrentIndex(SETTINGS_PAGE_OUTLINE)

    def _refresh_motion_label(self) -> None:
        motion_name = os.path.basename(self.motion_file) or "(no motion selected)"
        self.motion_label.setText(f"Motion: {motion_name}")

    def _emit_run_requested(self) -> None:
        self.run_requested.emit(self.to_dict())

    def _add_heading(self) -> None:
        item = QTreeWidgetItem(["New heading"])
        item.setFlags(
            item.flags()
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsUserCheckable
        )
        item.setCheckState(0, Qt.CheckState.Checked)
        item.setData(0, Qt.ItemDataRole.UserRole, "")
        self.outline_tree.addTopLevelItem(item)
        self.outline_tree.editItem(item, 0)

    def _item_from_node(self, node: OutlineNode) -> QTreeWidgetItem:
        item = QTreeWidgetItem([node.text])
        item.setData(0, Qt.ItemDataRole.UserRole, node.id)
        item.setFlags(
            item.flags()
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsUserCheckable
        )
        item.setCheckState(
            0,
            Qt.CheckState.Checked if node.selected else Qt.CheckState.Unchecked,
        )
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


# ── DEV-ONLY: lightweight worker for the temporary Re-verify button. ──
# Re-runs just the citation extraction + verification step on a body
# that's already in the output page (e.g. after a restart, when the
# previous .docx has been reloaded). Avoids re-running the full draft
# pipeline. Remove this class along with the Re-verify button when no
# longer needed.
class _ReverifyWorker(QThread):
    finished_result = Signal(bool, object)  # (success, list[CitationVerification] | error str)

    def __init__(self, body_text: str, parent=None):
        super().__init__(parent)
        self.body_text = body_text

    def run(self) -> None:
        try:
            from icharlotte_core.llm_config import call_llm

            citations = extract_citations(self.body_text)
            if not citations:
                self.finished_result.emit(True, [])
                return

            token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()

            def llm(system_prompt, user_prompt):
                return call_llm(
                    user_prompt,
                    system_prompt,
                    task_type="general",
                    agent_id="agent_oppose_motion",
                ) or ""

            verifier = build_opposition_verifier(
                courtlistener_token=token,
                llm_callback=llm,
                max_workers=4,
            )
            results = verifier.verify_all(citations)
            self.finished_result.emit(True, results)
        except Exception as exc:
            self.finished_result.emit(False, str(exc))
# ── END DEV-ONLY ──


class OpposeMotionOutputPage(CitationReviewOutputPage):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    default_title = "Opposition Memorandum"

    def empty_citations_message(self) -> str:
        return (
            "No citations were detected in this opposition. If California "
            "case-law research returned no results, the draft was written "
            "without case citations. Review the brief for any factual or "
            "statutory support that may need strengthening."
        )

    def _build_action_buttons(self, row) -> None:
        self._reverify_worker = None  # DEV-only worker handle
        # ── DEV-ONLY: re-verify button. Remove when no longer needed. ──
        self.reverify_btn = QPushButton("Re-verify Citations (DEV)")
        self.reverify_btn.setToolTip(
            "Temporary developer button. Re-extracts citations from the "
            "current body and re-runs the verifier so you can test "
            "parser / verifier changes without re-running the full "
            "draft pipeline."
        )
        self.reverify_btn.setStyleSheet("QPushButton { color: #b06000; }")
        self.reverify_btn.clicked.connect(self._on_reverify_clicked)
        row.addWidget(self.reverify_btn)
        # ── END DEV-ONLY ──
        self._add_save_button(row)

    # ── DEV-ONLY: re-verify handlers. Remove when no longer needed. ──

    def _on_reverify_clicked(self) -> None:
        if self._reverify_worker is not None:
            return  # already running
        body_text = (self.draft.body_text or "").strip()
        if not body_text:
            QMessageBox.warning(
                self,
                "Nothing to verify",
                "No draft body is loaded. Re-open or re-run the task first.",
            )
            return
        self.reverify_btn.setEnabled(False)
        self.reverify_btn.setText("Re-verifying… (DEV)")
        self.summary_banner.setText(
            "<b>Re-verifying citations…</b> (this runs the verifier on the "
            "current body without re-drafting)"
        )
        self.summary_banner.setVisible(True)
        worker = _ReverifyWorker(body_text=body_text, parent=None)
        worker.finished_result.connect(self._on_reverify_finished)
        worker.finished.connect(worker.deleteLater)
        self._reverify_worker = worker
        worker.start()

    def _on_reverify_finished(self, success: bool, payload: object) -> None:
        self._reverify_worker = None
        self.reverify_btn.setEnabled(True)
        self.reverify_btn.setText("Re-verify Citations (DEV)")
        if not success:
            QMessageBox.critical(
                self,
                "Re-verify failed",
                f"Re-verification failed:\n\n{payload}",
            )
            self._refresh_summary_banner()
            return
        citations = payload if isinstance(payload, list) else []
        self.draft.citations = citations
        self.editor.setHtml(_render_draft_html(self.draft))
        self._refresh_summary_banner()
        if citations:
            self.show_citation(0)
        else:
            self.detail_panel.clear(
                "Re-verify found no citations in the current body text."
            )

    # ── END DEV-ONLY ──


class OpposeMotionAnalysisWorker(QThread):
    progress = Signal(str)
    finished_analysis = Signal(bool, object)

    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.settings = dict(settings or {})

    def run(self) -> None:
        try:
            from icharlotte_core.llm_config import call_llm

            self.progress.emit("Extracting motion text...")
            motion_result = extract_document_text(self.settings.get("motion_file", ""))
            if not motion_result.success:
                message = motion_result.error or "Could not read motion."
                self.finished_analysis.emit(False, message)
                return

            self.progress.emit("Extracting context documents...")
            context_text, warnings = extract_context_bundle(
                self.settings.get("context_files", [])
            )
            for warning in warnings:
                self.progress.emit(f"WARNING: {warning}")

            def llm(system_prompt, user_prompt):
                return call_llm(
                    user_prompt,
                    system_prompt,
                    task_type="general",
                    agent_id="agent_chat",
                ) or ""

            self.progress.emit("Analyzing motion...")
            metadata = analyze_motion(motion_result.text, llm_callback=llm)
            missing = metadata.required_missing()
            if missing:
                self.finished_analysis.emit(
                    False,
                    "Could not automatically identify required motion fields: "
                    + ", ".join(missing)
                    + ". Confirm the motion has extractable text.",
                )
                return

            self.progress.emit("Generating opposition outline...")
            outline = generate_outline(
                metadata,
                context_text,
                llm_callback=llm,
            )
            if not outline:
                self.finished_analysis.emit(
                    False,
                    "Could not automatically generate an opposition outline.",
                )
                return

            self.finished_analysis.emit(
                True,
                {
                    "metadata": metadata,
                    "outline": outline,
                    "motion_text": motion_result.text,
                    "context_text": context_text,
                },
            )
        except Exception as exc:
            self.finished_analysis.emit(False, str(exc))


class OpposeMotionWorker(QThread):
    progress = Signal(str)
    finished_result = Signal(bool, object)

    def __init__(self, case_path: str, file_number: str, settings: dict, parent=None):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.settings = dict(settings or {})

    def run(self) -> None:
        try:
            from icharlotte_core.llm_config import call_llm

            self.progress.emit("Extracting motion text...")
            motion_result = extract_document_text(self.settings.get("motion_file", ""))
            if not motion_result.success:
                message = motion_result.error or "Could not read motion."
                self.finished_result.emit(False, message)
                return

            self.progress.emit("Extracting context documents...")
            context_text, warnings = extract_context_bundle(
                self.settings.get("context_files", [])
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

            def make_llm(pass_name):
                def _llm(system_prompt, user_prompt):
                    return call_llm(
                        user_prompt, system_prompt, task_type="general",
                        agent_id="agent_oppose_motion", pass_name=pass_name,
                        pass_agent_id="agent_oppose_motion",
                    ) or ""
                return _llm

            def llm(system_prompt, user_prompt):
                return call_llm(
                    user_prompt,
                    system_prompt,
                    task_type="general",
                    agent_id="agent_oppose_motion",
                ) or ""

            # Load style exemplars matching this motion type. The registry lives
            # under the repo's Scripts/prompts/oppose_motion/ directory; the
            # registry file is created on first save by the workbench so it may
            # be missing here — StyleExampleRegistry.load() handles that.
            self.progress.emit("Loading matching style exemplars...")
            repo_root = os.path.dirname(
                os.path.dirname(
                    os.path.dirname(
                        os.path.dirname(os.path.dirname(__file__))
                    )
                )
            )
            registry_path = os.path.join(
                repo_root,
                "Scripts",
                "prompts",
                "oppose_motion",
                "style_examples.json",
            )
            registry = StyleExampleRegistry.load(registry_path)
            matches = registry.matches_for_motion_type(metadata.motion_type)
            cache_dir = os.path.join(
                os.path.dirname(registry_path), ".cache", "style_examples"
            )
            exemplar_texts: list[str] = []
            for m in matches:
                text = extract_exemplar_text(m.path, cache_dir=cache_dir)
                if text.strip():
                    exemplar_texts.append(text)
            if matches:
                self.progress.emit(f"  Using {len(exemplar_texts)} style exemplar(s).")
            else:
                self.progress.emit("  No matching style exemplars; using default voice.")

            # Retrieval-first grounding: research real California authority for
            # each principal argument before drafting, so the drafter cites only
            # from a verified pool.
            # Prefer the local CA corpus (offline, unlimited). Fall back to the
            # live CourtListener API only if the corpus has not been built yet.
            corpus = _make_local_corpus()
            token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()
            retrieved = []
            opinion_cache = os.path.join(
                os.path.dirname(registry_path), ".cache", "opinions"
            )
            # Research at SUBSECTION granularity, not just the top-level
            # arguments. The drafter expands each principal argument into several
            # subsections, each making its own legal proposition (meet-and-confer,
            # discovery cutoff, cumulative discovery, etc.). Researching only the
            # top-level arguments left those sub-points with no authority, so the
            # drafter emitted "[no case authority retrieved for this point]".
            # Researching every selected section-plan leaf gives each subsection
            # its own on-point authority.
            research_targets = _research_targets(metadata, plan)
            if corpus is not None and research_targets:
                self.progress.emit(
                    f"Researching authorities locally ({len(research_targets)} points)..."
                )
                retrieved = research_arguments(
                    research_targets,
                    cl_client=corpus,
                    query_llm=make_llm("research_queries"),
                    rerank_llm=make_llm("rerank_select"),
                    # Keep concurrency low: the local corpus search is fast, so a
                    # higher worker count just bursts the LLM rate limit (429s)
                    # on the per-point query-gen + rerank calls.
                    max_workers=2,
                    on_progress=self.progress.emit,
                    cache_dir=opinion_cache,
                )
                self.progress.emit(f"Retrieved {len(retrieved)} grounded authorities.")
            elif corpus is None and token and research_targets:
                from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
                self.progress.emit(
                    "Local corpus not built; falling back to CourtListener API "
                    f"({len(research_targets)} points)..."
                )
                retrieved = research_arguments(
                    research_targets,
                    cl_client=CourtListenerClient(token),
                    query_llm=make_llm("research_queries"),
                    rerank_llm=make_llm("rerank_select"),
                    max_workers=2,  # avoid bursting provider rate limits
                    on_progress=self.progress.emit,
                    cache_dir=opinion_cache,
                )
                self.progress.emit(f"Retrieved {len(retrieved)} grounded authorities.")
            else:
                self.progress.emit(
                    "WARNING: no local corpus and no COURTLISTENER_API_TOKEN; "
                    "drafting without grounded research."
                )

            self.progress.emit("Drafting opposition memorandum...")
            draft = draft_memorandum(
                metadata=metadata,
                section_plan=plan,
                motion_text=motion_result.text,
                context_text=context_text,
                style_exemplars=exemplar_texts,
                retrieved_authorities=retrieved,
                llm_callback=llm,
            )
            if not draft.body_text.strip():
                reason = (draft.rejection_reason or "unknown reason").strip()
                self.finished_result.emit(False, f"Drafting failed: {reason}")
                return

            # Parse citations from the drafted body.
            citations = extract_citations(draft.body_text)
            if not citations:
                self.progress.emit(
                    "WARNING: No citations detected in the drafted opposition."
                )
                draft.citations = []
            else:
                to_verify, off_pool = pool_membership_check(citations, retrieved)
                if off_pool:
                    self.progress.emit(
                        f"{len(off_pool)} citation(s) were not in the researched pool "
                        "(flagged NOT_FOUND)."
                    )
                if corpus is None and not token:
                    self.progress.emit(
                        "WARNING: no local corpus and no COURTLISTENER_API_TOKEN; "
                        "case citations cannot be verified."
                    )
                self.progress.emit(f"Verifying citations ({len(to_verify)} found)...")
                if corpus is not None:
                    verifier = build_local_opposition_verifier(
                        corpus=corpus,
                        llm_callback=llm,
                        max_workers=4,
                    )
                else:
                    verifier = build_opposition_verifier(
                        courtlistener_token=token,
                        llm_callback=llm,
                        max_workers=4,
                    )
                verified = verifier.verify_all(to_verify, on_progress=self.progress.emit)
                # Merge verified + off-pool, restored to body order.
                draft.citations = sorted(
                    list(verified) + list(off_pool),
                    key=lambda cv: cv.body_offset if cv.body_offset is not None else 0,
                )
                enrich_with_pool_signals(draft.citations, retrieved)
                verdict_counts: dict[str, int] = {}
                for cv in draft.citations:
                    verdict_counts[cv.verdict] = verdict_counts.get(cv.verdict, 0) + 1
                summary = ", ".join(
                    f"{v.lower()}: {n}" for v, n in sorted(verdict_counts.items())
                )
                self.progress.emit(f"Verification complete ({summary}).")

            preview_dir = os.path.join(
                self.case_path,
                "NOTES",
                "AI OUTPUT",
                ".icharlotte",
                "wizard_previews",
                "oppose_motion",
            )
            preview_path = os.path.join(preview_dir, "Opposition Preview.docx")
            caption_path = DiscoveryAssembler.find_caption_page(self.case_path) or ""
            assemble_opposition_preview(
                draft=draft,
                output_path=preview_path,
                caption_path=caption_path,
            )
            validation = validate_opposition_docx(preview_path)
            if validation.has_errors:
                self.finished_result.emit(
                    False,
                    "Word validation failed for opposition preview.",
                )
                return
            draft.preview_path = preview_path
            self.finished_result.emit(True, draft)
        except Exception as exc:
            self.finished_result.emit(False, str(exc))


class OpposeMotionTaskTab(WizardTaskContainer):
    task_completed = Signal(dict)

    def __init__(
        self,
        spec,
        case_path: str,
        file_number: str,
        motion_file: str,
        context_files: list[str],
        auto_analyze: bool = False,
        parent: QWidget | None = None,
    ):
        super().__init__(spec, parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._files = [motion_file] + list(context_files)
        self._worker = None
        self._last_settings: dict = {}
        self._finishing_worker = None
        self._analysis_worker = None
        self._finishing_analysis_worker = None

        self.settings_page = OpposeMotionSettingsPage(
            case_path,
            file_number,
            motion_file,
            context_files,
        )
        self.status_page = StatusPage()
        self.output_page = OpposeMotionOutputPage()

        self.addWidget(self.settings_page)
        self.addWidget(self.status_page)
        self.addWidget(self.output_page)
        self.settings_page.run_requested.connect(self._on_run)
        if auto_analyze:
            self._start_analysis()

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list[str]:
        return list(self._files)

    def _start_analysis(self) -> None:
        if self._analysis_worker is not None or self._finishing_analysis_worker is not None:
            return
        self.status_page.reset()
        self.status_page.on_status("Analyzing selected motion and context...")
        self.status_page.progress_bar.setRange(0, 0)
        self.setCurrentIndex(TASK_PAGE_STATUS)
        worker = OpposeMotionAnalysisWorker(
            settings={
                "motion_file": self.settings_page.motion_file,
                "context_files": list(self.settings_page.context_files),
            },
            parent=None,
        )
        worker.progress.connect(self.status_page.on_status)
        worker.finished_analysis.connect(self._on_analysis_finished)
        if hasattr(worker, "finished"):
            worker.finished.connect(lambda w=worker: self._on_analysis_thread_finished(w))
        if hasattr(worker, "deleteLater"):
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
        # Skip the metadata-confirmation screen and go straight to the outline.
        # Metadata is still set on the page (line above) so it flows through to
        # the drafter via current_metadata() when the user clicks Generate Draft.
        self.settings_page.setCurrentIndex(SETTINGS_PAGE_OUTLINE)
        self.setCurrentIndex(TASK_PAGE_SETTINGS)

    def _on_analysis_thread_finished(self, worker) -> None:
        if self._analysis_worker is worker:
            self._analysis_worker = None
        if self._finishing_analysis_worker is worker:
            self._finishing_analysis_worker = None

    def _on_run(self, settings: dict) -> None:
        if self._analysis_worker is not None or self._finishing_analysis_worker is not None:
            self.status_page.on_status("Motion analysis is still running.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        if self._worker is not None or self._finishing_worker is not None:
            self.status_page.on_status("Opposition draft is already running.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        self.status_page.reset()
        self.status_page.on_status("Drafting opposition memorandum...")
        self.status_page.progress_bar.setRange(0, 0)
        self.setCurrentIndex(TASK_PAGE_STATUS)
        self._last_settings = dict(settings or {})
        worker = OpposeMotionWorker(
            case_path=self._case_path,
            file_number=self._file_number,
            settings=self._last_settings,
            parent=None,
        )
        worker.progress.connect(self.status_page.on_status)
        worker.finished_result.connect(self._on_worker_finished)
        if hasattr(worker, "finished"):
            worker.finished.connect(lambda w=worker: self._on_worker_thread_finished(w))
        if hasattr(worker, "deleteLater"):
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
                "files": list(self._files),
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
                    "The opposition draft is still running. Wait for it to finish before closing this tab.",
                )
                event.ignore()
                return
        super().closeEvent(event)


def build_oppose_motion_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None,
):
    motion_file, _ = QFileDialog.getOpenFileName(
        parent,
        "Select motion to oppose",
        case_path,
        "Motion files (*.pdf *.docx)",
    )
    if not motion_file:
        return None
    if not is_supported_motion_file(motion_file):
        QMessageBox.warning(
            parent,
            "Unsupported motion file",
            "Select a PDF or DOCX motion.",
        )
        return None

    context_files = ContextFilesDialog.get_files(
        parent,
        title="Select context document(s)",
        start_dir=os.path.dirname(motion_file) or case_path,
        file_filter="Context files (*.pdf *.docx *.txt *.msg);;All files (*.*)",
    )
    context_files = [
        path for path in (context_files or []) if is_supported_context_file(path)
    ]
    return OpposeMotionTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        motion_file=motion_file,
        context_files=list(context_files),
        auto_analyze=True,
        parent=parent,
    )
