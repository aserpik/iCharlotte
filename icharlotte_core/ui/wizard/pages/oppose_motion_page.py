"""Custom Wizard task page for drafting oppositions to motions."""

from __future__ import annotations

import html
import os
import re
import shutil

from PySide6.QtCore import QThread, QUrl, Qt, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QStackedWidget,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.discovery.assembler import DiscoveryAssembler
from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
from icharlotte_core.opposition.assembler import assemble_opposition_preview
from icharlotte_core.opposition.authority import research_opposition_authorities
from icharlotte_core.opposition.citation_verifier import verify_citations
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
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.word_validator import validate_opposition_docx


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


class OpposeMotionOutputPage(QWidget):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.draft = DraftDocument()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        layout = QHBoxLayout()
        layout.setSpacing(12)
        self.editor = QTextBrowser()
        self.editor.setOpenLinks(False)
        self.editor.setOpenExternalLinks(False)
        self.editor.anchorClicked.connect(self._on_anchor_clicked)
        layout.addWidget(self.editor, 2)

        self.source_drawer = QPlainTextEdit()
        self.source_drawer.setReadOnly(True)
        layout.addWidget(self.source_drawer, 1)
        outer.addLayout(layout, 1)

        row = QHBoxLayout()
        row.addStretch()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_as)
        row.addWidget(self.save_btn)
        outer.addLayout(row)

    def show_result(self, draft: DraftDocument) -> None:
        self.draft = draft
        self.editor.setHtml(_render_draft_html(draft))
        if draft.citations:
            self.show_citation(0)
        else:
            self.source_drawer.setPlainText(
                "No citations were detected in this opposition.\n\n"
                "If California case-law research returned no results, the draft "
                "was written without case citations. Review the brief for any "
                "factual or statutory support that may need strengthening."
            )

    def _on_anchor_clicked(self, url: QUrl) -> None:
        scheme = url.scheme()
        if scheme == "citation":
            try:
                index = int(url.path().lstrip("/") or url.host() or "0")
            except (TypeError, ValueError):
                return
            self.show_citation(index)
            self.open_citation_dialog(index)
            return
        QDesktopServices.openUrl(url)

    def open_citation_dialog(self, index: int) -> None:
        if index < 0 or index >= len(self.draft.citations):
            return
        dialog = CitationDetailDialog(self.draft.citations[index], parent=self)
        dialog.exec()

    @property
    def output_path(self) -> str:
        return self.draft.preview_path

    def load_output(self, output_path: str) -> None:
        body_text = ""
        if output_path and output_path.lower().endswith(".docx") and os.path.isfile(output_path):
            try:
                from docx import Document

                doc = Document(output_path)
                body_text = "\n".join(p.text for p in doc.paragraphs if p.text)
            except Exception:
                body_text = f"(Could not render generated document: {output_path})"
        self.show_result(
            DraftDocument(
                title=os.path.splitext(os.path.basename(output_path or ""))[0] or "Opposition",
                body_text=body_text,
                preview_path=output_path or "",
            )
        )

    @staticmethod
    def default_save_dir(preview_path: str) -> str:
        marker = os.path.join(".icharlotte", "wizard_previews")
        before_marker, _, _after_marker = preview_path.partition(marker)
        if before_marker:
            return os.path.dirname(os.path.normpath(before_marker))
        return os.path.dirname(preview_path)

    def show_citation(self, index: int) -> None:
        if index < 0 or index >= len(self.draft.citations):
            return
        citation = self.draft.citations[index]
        self.source_drawer.setPlainText(
            "\n".join(
                [
                    citation.citation_text,
                    f"Normalized: {citation.normalized_citation}",
                    f"Status: {citation.status}",
                    f"Case: {citation.case_name}",
                    f"Court: {citation.court}",
                    f"Date: {citation.date}",
                    f"Opinion: {citation.opinion_url}",
                    "",
                    "Supporting passage:",
                    citation.supporting_passage or "(support not confirmed)",
                    "",
                    citation.warning,
                ]
            ).strip()
        )

    def save_as(self) -> None:
        if not self.draft.preview_path:
            QMessageBox.warning(
                self,
                "No preview",
                "No generated opposition preview is available.",
            )
            return
        suggested = os.path.join(
            self.default_save_dir(self.draft.preview_path),
            f"{self.draft.title or 'Opposition Memorandum'}.docx",
        )
        target, _ = QFileDialog.getSaveFileName(
            self,
            "Save Opposition Memorandum",
            suggested,
            "Word Documents (*.docx);;All files (*.*)",
        )
        if not target:
            return
        if not target.lower().endswith(".docx"):
            target += ".docx"
        if os.path.abspath(target) == os.path.abspath(self.draft.preview_path):
            QMessageBox.warning(
                self,
                "Choose another location",
                "Select a location outside the internal preview file.",
            )
            return
        try:
            shutil.copyfile(self.draft.preview_path, target)
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", f"Could not save file:\n{exc}")
            return
        QMessageBox.information(self, "Saved", f"Saved:\n{target}")


_HORIZONTAL_RULE_RE = re.compile(r"^[\*\-_]{3,}\s*$")
_MD_HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
_MD_ITALIC_RE = re.compile(r"\*([^\*\n]+?)\*")


def _render_draft_html(draft: DraftDocument) -> str:
    """Render the draft body into HTML with clickable citation anchors."""
    citation_spans = _build_citation_index(draft)
    body_html_lines: list[str] = []

    for raw_line in (draft.body_text or "").splitlines():
        stripped = raw_line.strip()
        if not stripped:
            body_html_lines.append("<p>&nbsp;</p>")
            continue
        if _HORIZONTAL_RULE_RE.match(stripped):
            continue
        heading_match = _MD_HEADING_RE.match(stripped)
        if heading_match:
            level = min(max(len(heading_match.group(1)), 2), 4)
            content = _format_inline_html(heading_match.group(2), citation_spans)
            body_html_lines.append(f"<h{level}>{content}</h{level}>")
            continue
        content = _format_inline_html(stripped, citation_spans)
        body_html_lines.append(f"<p>{content}</p>")

    title = html.escape(draft.title or "Opposition Memorandum")
    body = "\n".join(body_html_lines)
    return (
        "<html><body style=\"font-family:'Times New Roman',serif; font-size:13pt;\">"
        f"<h1 style=\"text-align:center;\">{title}</h1>{body}</body></html>"
    )


def _build_citation_index(draft: DraftDocument) -> list[tuple[str, int]]:
    """Return [(citation_text, draft_citation_index), ...] sorted by length desc.

    Sorting by length avoids partial-match collisions (e.g. "62 Cal. 4th 1081"
    must be wrapped before a shorter substring matches inside it).
    """
    spans: list[tuple[str, int]] = []
    for index, citation in enumerate(draft.citations or []):
        text = (citation.citation_text or "").strip()
        if text:
            spans.append((text, index))
    spans.sort(key=lambda pair: len(pair[0]), reverse=True)
    return spans


def _format_inline_html(line: str, citation_spans: list[tuple[str, int]]) -> str:
    """Escape, italicize *case names*, and wrap citation texts as clickable anchors."""
    italicized = _MD_ITALIC_RE.sub(
        lambda match: f"\x00ITA{html.escape(match.group(1))}\x00ITAEND",
        line,
    )
    escaped = html.escape(italicized)
    escaped = escaped.replace("\x00ITA", "<i>").replace("\x00ITAEND", "</i>")
    for citation_text, index in citation_spans:
        if not citation_text:
            continue
        pattern = re.escape(html.escape(citation_text))
        anchor = (
            f"<a href=\"citation:{index}\" "
            f"style=\"color:#1a5dbf; text-decoration:underline;\">"
            f"{html.escape(citation_text)}</a>"
        )
        escaped = re.sub(pattern, anchor, escaped, count=0)
    return escaped


class CitationDetailDialog(QDialog):
    """Modal dialog showing a single citation's verification details."""

    def __init__(self, citation, parent: QWidget | None = None):
        super().__init__(parent)
        self.citation = citation
        self.setWindowTitle(citation.case_name or citation.citation_text or "Citation")
        self.resize(720, 540)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        header = QLabel(self._header_html(citation))
        header.setTextFormat(Qt.TextFormat.RichText)
        header.setWordWrap(True)
        header.setOpenExternalLinks(False)
        layout.addWidget(header)

        passage_label = QLabel("Supporting passage from opinion:")
        passage_label.setStyleSheet("font-weight: 600; margin-top: 4px;")
        layout.addWidget(passage_label)

        self.passage_view = QTextBrowser()
        self.passage_view.setOpenLinks(False)
        self.passage_view.setHtml(self._passage_html(citation))
        layout.addWidget(self.passage_view, 1)

        if citation.warning:
            warning_label = QLabel(html.escape(citation.warning))
            warning_label.setStyleSheet("color: #b85c00;")
            warning_label.setWordWrap(True)
            layout.addWidget(warning_label)

        buttons = QDialogButtonBox()
        self.open_btn = QPushButton("Open in CourtListener")
        self.open_btn.setEnabled(bool(citation.opinion_url))
        self.open_btn.clicked.connect(self._open_opinion_url)
        buttons.addButton(self.open_btn, QDialogButtonBox.ButtonRole.ActionRole)
        close_btn = buttons.addButton(QDialogButtonBox.StandardButton.Close)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(buttons)

    @staticmethod
    def _header_html(citation) -> str:
        rows: list[tuple[str, str]] = []
        if citation.case_name:
            rows.append(("Case", citation.case_name))
        if citation.citation_text:
            rows.append(("Citation", citation.citation_text))
        if citation.court:
            rows.append(("Court", citation.court))
        if citation.date:
            rows.append(("Date", citation.date))
        if citation.status:
            rows.append(("Status", citation.status.replace("_", " ")))
        cells = "".join(
            f"<tr><td style=\"padding:2px 8px; color:#555;\">{html.escape(label)}</td>"
            f"<td style=\"padding:2px 0;\">{html.escape(value)}</td></tr>"
            for label, value in rows
        )
        return f"<table style=\"font-size:11pt;\">{cells}</table>"

    @staticmethod
    def _passage_html(citation) -> str:
        if citation.supporting_passage:
            passage = html.escape(citation.supporting_passage)
            return (
                "<div style=\"font-family:'Times New Roman',serif; font-size:12pt;\">"
                f"<p style=\"background:#fff7c2; padding:6px;\">{passage}</p></div>"
            )
        if citation.status in ("throttled", "exists_support_unconfirmed"):
            note = (
                "CourtListener did not return the supporting opinion text "
                "(rate-limited or text unavailable). The citation exists, but "
                "automatic support extraction did not complete. Click "
                "<i>Open in CourtListener</i> to read the opinion."
            )
        elif citation.status == "not_found":
            note = (
                "CourtListener could not locate this citation. The case may be "
                "unpublished, mis-cited, or absent from the database. Verify "
                "the citation manually before relying on it."
            )
        elif citation.status == "invalid":
            note = (
                "CourtListener rejected the citation as malformed. Re-check "
                "the citation format before relying on it."
            )
        else:
            note = (
                "No supporting passage was extracted for this citation. Open "
                "the opinion in CourtListener to confirm the proposition."
            )
        return (
            "<div style=\"font-family:'Times New Roman',serif; color:#444;\">"
            f"<p>{note}</p></div>"
        )

    def _open_opinion_url(self) -> None:
        if self.citation.opinion_url:
            QDesktopServices.openUrl(QUrl(self.citation.opinion_url))


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

            def llm(system_prompt, user_prompt):
                return call_llm(
                    user_prompt,
                    system_prompt,
                    task_type="general",
                    agent_id="agent_chat",
                ) or ""

            token = os.environ.get("COURTLISTENER_API_TOKEN", "").strip()
            if not token:
                self.finished_result.emit(
                    False,
                    (
                        "CourtListener API token missing; cannot research and cite "
                        "California case law for the opposition."
                    ),
                )
                return

            self.progress.emit("Researching California authorities...")
            client = CourtListenerClient(token)
            authority_packet = research_opposition_authorities(
                metadata=metadata,
                section_plan=plan,
                motion_text=motion_result.text,
                context_text=context_text,
                courtlistener=client,
                llm_callback=llm,
                status_callback=self.progress.emit,
            )
            for warning in getattr(authority_packet, "warnings", []) or []:
                self.progress.emit(f"WARNING: {warning}")
            if not getattr(authority_packet, "cases", []):
                self.progress.emit(
                    "WARNING: No California case law was found for the selected "
                    "opposition issues. Drafting an opposition without case-law "
                    "citations; statutes and facts from the moving papers may "
                    "still be cited."
                )

            self.progress.emit("Drafting memorandum with researched authorities...")
            draft = draft_memorandum(
                metadata=metadata,
                section_plan=plan,
                motion_text=motion_result.text,
                context_text=context_text,
                authority_block=authority_packet.authority_block,
                llm_callback=llm,
            )
            if not draft.body_text.strip():
                reason = (draft.rejection_reason or "unknown reason").strip()
                self.finished_result.emit(
                    False,
                    f"Drafting failed: {reason}",
                )
                return

            if authority_packet.cases:
                self.progress.emit("Verifying citations...")
                draft.citations = verify_citations(
                    draft.body_text,
                    citation_propositions=authority_packet.citation_propositions,
                    courtlistener=client,
                )
                if not draft.citations:
                    self.progress.emit(
                        "WARNING: No citations were detected in the drafted "
                        "opposition. Review the draft for unsupported assertions."
                    )
            else:
                draft.citations = []

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


class OpposeMotionTaskTab(QStackedWidget):
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
        super().__init__(parent)
        self._spec = spec
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
        self.settings_page.setCurrentIndex(SETTINGS_PAGE_CONFIRM)
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

    context_files, _ = QFileDialog.getOpenFileNames(
        parent,
        "Select context document(s)",
        os.path.dirname(motion_file) or case_path,
        "Context files (*.pdf *.docx *.txt *.msg);;All files (*.*)",
    )
    context_files = [
        path for path in (context_files or []) if is_supported_context_file(path)
    ]
    return OpposeMotionTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        motion_file=motion_file,
        context_files=list(context_files or []),
        auto_analyze=True,
        parent=parent,
    )
