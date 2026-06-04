"""Wizard task: defense-side Mediation Brief generation.

Drives the existing ``MediationBriefGenerator`` (unchanged) through a
Settings -> Status -> Output flow. Generation only; section refinement and
deposition-quote insertion are provided by the Word AI Assistant (Win+V) on
the generated document.
"""
from __future__ import annotations

import logging
import os
import re
import shutil

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.mediation_brief import (
    GENERATION_ORDER,
    SECTION_HEADINGS,
    MediationBriefGenerator,
)
from icharlotte_core.ui.context_files_dialog import ContextFilesDialog
from icharlotte_core.ui.wizard import theme
from icharlotte_core.ui.wizard.docx_io import load_docx_as_html
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer

logger = logging.getLogger(__name__)

TASK_PAGE_SETTINGS = 0
TASK_PAGE_STATUS = 1
TASK_PAGE_OUTPUT = 2

DEFAULT_BRIEF_FILENAME = "Defendant's Confidential Mediation Brief.docx"
MEDIATION_SUBDIR = ("NOTES", "AI OUTPUT", "MEDIATION")


# --------------------------- document extraction ---------------------------

def _read_pdf(path: str) -> str:
    import fitz

    parts = []
    doc = fitz.open(path)
    try:
        for page in doc:
            parts.append(page.get_text())
    finally:
        doc.close()
    return "\n".join(parts)


def _read_doc_via_word(path: str) -> str:
    """Read a legacy .doc via read-only Word COM. Never Quit/Visible the user's
    Word (global safety rule); only close the document we open."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = None
        try:
            doc = word.Documents.Open(
                FileName=os.path.abspath(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            return doc.Content.Text or ""
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
    finally:
        pythoncom.CoUninitialize()


def _read_msg(path: str) -> str:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        item = namespace.OpenSharedItem(os.path.abspath(path))
        try:
            subject = item.Subject or ""
            sender = item.SenderName or ""
            body = item.Body or ""
            if not body.strip() and item.HTMLBody:
                body = re.sub(r"<[^>]+>", "", item.HTMLBody).strip()
            return f"From: {sender}\nSubject: {subject}\n\n{body}"
        finally:
            try:
                item.Close(0)
            except Exception:
                pass
    finally:
        pythoncom.CoUninitialize()


def _read_documents(paths) -> tuple[str, list[str]]:
    """Extract text from the given files into a single string + warning list.

    Table-aware for .docx (uses ``extract_docx_text`` — legal chronologies and
    depo summaries live inside tables; a ``doc.paragraphs`` loop loses them).
    Mirrors the Chat-tab reader so the Wizard feeds the engine identical input.
    """
    content_parts: list[str] = []
    warnings: list[str] = []
    for path in paths or []:
        if not path or not os.path.isfile(path):
            warnings.append(f"File not found: {path}")
            continue
        ext = os.path.splitext(path)[1].lower()
        name = os.path.basename(path)
        try:
            if ext == ".txt":
                with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                    text = fh.read()
            elif ext == ".docx":
                from icharlotte_core.document_processor import extract_docx_text

                text = extract_docx_text(path)
            elif ext == ".doc":
                text = _read_doc_via_word(path)
            elif ext == ".pdf":
                text = _read_pdf(path)
            elif ext == ".msg":
                text = _read_msg(path)
            else:
                warnings.append(f"Unsupported file type skipped: {name}")
                continue
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"Could not read {name}: {exc}")
            continue
        if text and text.strip():
            content_parts.append(f"--- FILE: {name} ---\n{text}")
        else:
            warnings.append(f"No text extracted from {name}")
    return "\n\n".join(content_parts), warnings


# ------------------------------- assembly ---------------------------------

def _save_brief(generator, caption_path: str, dest_dir: str, filename: str) -> str:
    """Assemble the brief into ``dest_dir/filename``. If the destination is
    locked (open in Word), fall back to a counter-suffixed name. Returns the
    path actually written. Never closes the user's Word (global safety rule)."""
    os.makedirs(dest_dir, exist_ok=True)
    base, ext = os.path.splitext(filename)
    candidates = [filename] + [f"{base} ({i}){ext}" for i in range(2, 21)]
    last_error: Exception | None = None
    for candidate in candidates:
        dest = os.path.join(dest_dir, candidate)
        try:
            generator.assemble_document(caption_path, dest)
            return dest
        except PermissionError as exc:  # destination locked — try next name
            last_error = exc
            continue
    raise last_error or PermissionError("Could not save the mediation brief.")


# -------------------------------- worker ----------------------------------

class MediationBriefWizardWorker(QThread):
    progress = Signal(str)
    finished_result = Signal(bool, object)  # (success, output_path | error_msg)

    def __init__(self, case_path: str, file_number: str, settings: dict, parent=None):
        super().__init__(parent)
        self.case_path = case_path or ""
        self.file_number = file_number or ""
        self.settings = dict(settings or {})
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def _emit_cancelled(self) -> bool:
        if self._cancelled:
            self.finished_result.emit(False, "Generation cancelled.")
            return True
        return False

    def run(self) -> None:
        try:
            files = list(self.settings.get("files", []))
            caption_path = self.settings.get("caption_path", "")

            self.progress.emit("Reading source documents…")
            content, warnings = _read_documents(files)
            for warning in warnings:
                self.progress.emit(f"WARNING: {warning}")
            if not content.strip():
                self.finished_result.emit(
                    False, "Could not read any text from the selected documents."
                )
                return
            if not caption_path or not os.path.isfile(caption_path):
                self.finished_result.emit(
                    False, "No caption template selected (or it no longer exists)."
                )
                return
            if self._emit_cancelled():
                return

            generator = MediationBriefGenerator()
            generator.caption_template_path = caption_path
            generator.document_content = content

            self.progress.emit("Loading style reference…")
            generator.get_style_excerpts()

            if self._emit_cancelled():
                return
            self.progress.emit("Analyzing documents…")
            generator.run_planning_pass()

            total = len(GENERATION_ORDER)
            for index, section_name in enumerate(GENERATION_ORDER, start=1):
                if self._emit_cancelled():
                    return
                _, title = SECTION_HEADINGS.get(section_name, ("", section_name))
                label = title or section_name
                self.progress.emit(f"Generating {label} ({index} of {total})…")
                result = generator.generate_section(section_name)
                if result:
                    generator.sections[section_name] = result
            generator.is_active = True

            if self._emit_cancelled():
                return
            self.progress.emit("Assembling document…")
            dest_dir = os.path.join(self.case_path, *MEDIATION_SUBDIR)
            filename = self.settings.get("suggested_filename") or DEFAULT_BRIEF_FILENAME
            saved_path = _save_brief(generator, caption_path, dest_dir, filename)
            self.progress.emit(f"Saved to {saved_path}")
            self.finished_result.emit(True, saved_path)
        except Exception as exc:  # noqa: BLE001
            logger.exception("MediationBriefWizardWorker failed")
            self.finished_result.emit(False, str(exc))


# ----------------------------- settings page ------------------------------

class MediationBriefSettingsPage(QWidget):
    """Source documents + caption template + Generate."""

    run_requested = Signal(dict)

    def __init__(self, case_path: str, file_number: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.case_path = case_path or ""
        self.file_number = file_number or ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.helper_text(
            "Generate a defense-side mediation brief from the selected case "
            "documents. The brief is built section by section and saved to "
            "NOTES/AI OUTPUT/MEDIATION. Refine sections or add deposition quotes "
            "afterward in Word (Win+V → Mediation Brief)."
        ))

        files_row = QHBoxLayout()
        files_row.addWidget(theme.section_header("Source documents"))
        files_row.addStretch()
        self.add_files_btn = theme.secondary_button("Add Files…")
        self.add_files_btn.clicked.connect(self._on_add_files)
        files_row.addWidget(self.add_files_btn)
        self.remove_files_btn = theme.secondary_button("Remove")
        self.remove_files_btn.clicked.connect(self._on_remove_files)
        files_row.addWidget(self.remove_files_btn)
        layout.addLayout(files_row)

        self.files_list = QListWidget()
        self.files_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        layout.addWidget(self.files_list, 1)

        layout.addWidget(theme.section_header("Caption template"))
        caption_row = QHBoxLayout()
        self.caption_edit = QLineEdit()
        self.caption_edit.setPlaceholderText("Path to the case caption .docx…")
        self.caption_edit.textChanged.connect(self._update_caption_hint)
        caption_row.addWidget(self.caption_edit, 1)
        self.browse_caption_btn = theme.secondary_button("Browse…")
        self.browse_caption_btn.clicked.connect(self._on_browse_caption)
        caption_row.addWidget(self.browse_caption_btn)
        layout.addLayout(caption_row)
        self.caption_hint = theme.error_text("")
        layout.addWidget(self.caption_hint)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.generate_btn = theme.primary_button("Generate Mediation Brief")
        self.generate_btn.clicked.connect(self._on_generate)
        btn_row.addWidget(self.generate_btn)
        layout.addLayout(btn_row)

        self._autodetect_caption()

    # ---- helpers ----

    def _autodetect_caption(self) -> None:
        if self.case_path:
            try:
                found = MediationBriefGenerator().find_caption_template(self.case_path)
            except Exception:  # noqa: BLE001
                found = None
            if found:
                self.caption_edit.setText(found)
        self._update_caption_hint()

    def _update_caption_hint(self, *_args) -> None:
        path = self.caption_edit.text().strip()
        if not path:
            self.caption_hint.setText(
                "No caption template found — click Browse… to select one."
            )
        elif not os.path.isfile(path):
            self.caption_hint.setText("Caption template not found at that path.")
        else:
            self.caption_hint.setText("")

    def current_files(self) -> list[str]:
        return [self.files_list.item(i).text() for i in range(self.files_list.count())]

    def _on_add_files(self) -> None:
        picked = ContextFilesDialog.get_files(
            self,
            title="Select source document(s) for the mediation brief",
            start_dir=self.case_path or "",
            file_filter="Documents (*.pdf *.docx *.doc *.txt *.msg);;All files (*.*)",
        )
        existing = set(self.current_files())
        for path in picked or []:
            if path not in existing:
                self.files_list.addItem(path)
                existing.add(path)

    def _on_remove_files(self) -> None:
        for item in self.files_list.selectedItems():
            self.files_list.takeItem(self.files_list.row(item))

    def _on_browse_caption(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select caption template", self.case_path or "",
            "Word Documents (*.docx)",
        )
        if path:
            self.caption_edit.setText(path)

    def to_dict(self) -> dict:
        return {
            "files": self.current_files(),
            "caption_path": self.caption_edit.text().strip(),
        }

    def from_dict(self, data: dict | None) -> None:
        data = data or {}
        self.files_list.clear()
        for path in data.get("files", []) or []:
            self.files_list.addItem(path)
        caption = data.get("caption_path", "")
        if caption:
            self.caption_edit.setText(caption)
        self._update_caption_hint()

    def _on_generate(self) -> None:
        files = self.current_files()
        caption_path = self.caption_edit.text().strip()
        if not files:
            QMessageBox.warning(self, "No documents", "Add at least one source document.")
            return
        if not caption_path or not os.path.isfile(caption_path):
            QMessageBox.warning(
                self, "No caption template", "Select a caption template (.docx)."
            )
            return
        self.run_requested.emit({
            "files": files,
            "caption_path": caption_path,
            "save_default_dir": os.path.join(self.case_path, *MEDIATION_SUBDIR),
            "suggested_filename": DEFAULT_BRIEF_FILENAME,
        })


# ------------------------------ output page -------------------------------

class MediationBriefOutputPage(QWidget):
    """Read-only preview + Open in Word + Save a Copy As…."""

    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._output_path = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        layout.setSpacing(theme.SPACE_MD)

        self.saved_banner = QLabel("")
        self.saved_banner.setWordWrap(True)
        self.saved_banner.setStyleSheet(
            f"background-color: {theme.SUCCESS_BG}; color: {theme.SUCCESS};"
            f" border-radius: {theme.RADIUS_SM}px; padding: 8px 12px;"
            f" font-weight: 600;"
        )
        self.saved_banner.setVisible(False)
        layout.addWidget(self.saved_banner)

        self.preview = QTextBrowser()
        layout.addWidget(self.preview, 1)

        btn_row = QHBoxLayout()
        self.open_btn = theme.secondary_button("Open in Word")
        self.open_btn.clicked.connect(self._on_open)
        self.open_btn.setEnabled(False)
        btn_row.addWidget(self.open_btn)
        self.save_as_btn = theme.secondary_button("Save a Copy As…")
        self.save_as_btn.clicked.connect(self._on_save_as)
        self.save_as_btn.setEnabled(False)
        btn_row.addWidget(self.save_as_btn)
        btn_row.addStretch()
        self.rerun_btn = theme.secondary_button("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_btn = theme.secondary_button("Edit Settings & Re-run")
        self.edit_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_btn)
        layout.addLayout(btn_row)

    @property
    def output_path(self) -> str:
        return self._output_path

    def show_result(self, path: str) -> None:
        self._output_path = path or ""
        ok = bool(self._output_path) and os.path.isfile(self._output_path)
        self.open_btn.setEnabled(ok)
        self.save_as_btn.setEnabled(ok)
        if ok:
            self.saved_banner.setText(f"Saved to: {self._output_path}")
            self.saved_banner.setVisible(True)
            try:
                self.preview.setHtml(load_docx_as_html(self._output_path))
            except Exception as exc:  # noqa: BLE001
                self.preview.setPlainText(f"(Could not render preview: {exc})")
        else:
            self.saved_banner.setVisible(False)
            self.preview.setPlainText("(No output produced.)")

    def load_output(self, path: str) -> None:
        self.show_result(path)

    def _on_open(self) -> None:
        if self._output_path and os.path.isfile(self._output_path):
            try:
                os.startfile(self._output_path)  # Windows
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "Open failed", f"Could not open in Word:\n{exc}")

    def _on_save_as(self) -> None:
        if not self._output_path or not os.path.isfile(self._output_path):
            return
        default = os.path.join(
            os.path.dirname(self._output_path), os.path.basename(self._output_path)
        )
        target, _ = QFileDialog.getSaveFileName(
            self, "Save a copy of the mediation brief", default,
            "Word Documents (*.docx);;All files (*.*)",
        )
        if not target:
            return
        if not target.lower().endswith(".docx"):
            target += ".docx"
        try:
            shutil.copyfile(self._output_path, target)
            QMessageBox.information(self, "Saved", f"Saved a copy to:\n{target}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Save failed", f"Could not save a copy:\n{exc}")


# --------------------------------- tab ------------------------------------

class MediationBriefTaskTab(WizardTaskContainer):
    """Settings → Status → Output container for the Mediation Brief task."""

    task_completed = Signal(dict)

    def __init__(self, spec, case_path: str, file_number: str, parent: QWidget | None = None):
        super().__init__(spec, parent=parent)
        self._case_path = case_path
        self._file_number = file_number
        self._worker = None
        self._finishing_worker = None
        self._last_settings: dict = {}

        self.settings_page = MediationBriefSettingsPage(case_path, file_number)
        self.status_page = StatusPage()
        self.output_page = MediationBriefOutputPage()

        self.addWidget(self.settings_page)   # TASK_PAGE_SETTINGS
        self.addWidget(self.status_page)     # TASK_PAGE_STATUS
        self.addWidget(self.output_page)     # TASK_PAGE_OUTPUT

        self.settings_page.run_requested.connect(self._on_run)
        self.status_page.cancel_requested.connect(self._on_cancel)
        self.output_page.rerun_requested.connect(self._on_rerun)
        self.output_page.edit_settings_requested.connect(
            lambda: self.setCurrentIndex(TASK_PAGE_SETTINGS)
        )

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list[str]:
        return list(self.settings_page.current_files())

    def _on_run(self, settings: dict) -> None:
        if self._worker is not None or self._finishing_worker is not None:
            self.status_page.on_status("A mediation brief is already generating.")
            self.setCurrentIndex(TASK_PAGE_STATUS)
            return
        self._last_settings = dict(settings or {})
        self.status_page.reset()
        self.status_page.on_status("Starting mediation brief…")
        self.status_page.progress_bar.setRange(0, 0)  # indeterminate
        self.setCurrentIndex(TASK_PAGE_STATUS)

        worker = MediationBriefWizardWorker(
            self._case_path, self._file_number, self._last_settings, parent=None
        )
        worker.progress.connect(self.status_page.on_status)
        worker.finished_result.connect(self._on_worker_finished)
        worker.finished.connect(lambda w=worker: self._on_worker_thread_finished(w))
        worker.finished.connect(worker.deleteLater)
        self._worker = worker
        worker.start()

    def _on_cancel(self) -> None:
        if self._worker is not None:
            self.status_page.on_status("Cancelling… (finishing the current section)")
            try:
                self._worker.cancel()
            except Exception:  # noqa: BLE001
                pass
        else:
            self.setCurrentIndex(TASK_PAGE_SETTINGS)

    def _on_rerun(self) -> None:
        self._on_run(self._last_settings)

    def _on_worker_finished(self, success: bool, payload: object) -> None:
        from datetime import datetime

        if self.sender() is not None and self.sender() is not self._worker:
            return
        if self.sender() is self._worker:
            self._finishing_worker = self._worker
            self._worker = None

        if not success:
            self.status_page.on_status(f"FAILED: {payload}")
            # Re-enable the status button so the user can go back (routes through
            # cancel_requested → _on_cancel, which returns to Settings when no
            # worker is running). No disconnect/reconnect — keeps cancel wiring
            # intact across re-runs.
            self.status_page.cancel_btn.setEnabled(True)
            self.status_page.cancel_btn.setText("Back to Settings")
            return

        path = payload if isinstance(payload, str) else ""
        self.output_page.show_result(path)
        self.setCurrentIndex(TASK_PAGE_OUTPUT)
        self.task_completed.emit({
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": list(self.settings_page.current_files()),
            "settings": self.settings_page.to_dict(),
            "output_path": path,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        })

    def _on_worker_thread_finished(self, worker) -> None:
        if self._worker is worker:
            self._worker = None
        if self._finishing_worker is worker:
            self._finishing_worker = None

    def closeEvent(self, event) -> None:
        for worker in (self._worker, self._finishing_worker):
            if worker is not None and worker.isRunning():
                QMessageBox.information(
                    self, "Task running",
                    "The mediation brief is still generating. Wait for it to "
                    "finish before closing this tab.",
                )
                event.ignore()
                return
        super().closeEvent(event)


def build_mediation_brief_tab(spec, case_path: str, file_number: str, parent: QWidget | None):
    """Open the Mediation Brief task on its Settings page."""
    return MediationBriefTaskTab(
        spec=spec, case_path=case_path, file_number=file_number, parent=parent
    )
