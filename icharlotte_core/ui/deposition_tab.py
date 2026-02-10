"""
Depositions Tab — Split-panel UI for deposition testimony extraction.

Left panel: Transcript viewer with page/line numbers
Right panel: Extraction results with Q/A pairs and citations
Bottom: Prompt input bar
Top: Toolbar (load, transcript selector, export, highlight checkbox, deponent info)

Supports multiple transcripts per case with drag-and-drop loading.
Loaded transcripts persist within a case across sessions.
"""

import os
import json
import logging

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QLabel,
    QPushButton, QTextEdit, QPlainTextEdit, QCheckBox,
    QFileDialog, QMessageBox, QProgressBar, QApplication,
    QToolBar, QSizePolicy, QComboBox
)
from PySide6.QtCore import Qt, Signal, QThread, QTimer, QSettings
from PySide6.QtGui import (
    QTextCursor, QFont, QTextCharFormat, QColor,
    QDragEnterEvent, QDropEvent, QKeyEvent
)

from icharlotte_core.config import GEMINI_DATA_DIR

logger = logging.getLogger(__name__)


# =============================================================================
# Background Workers
# =============================================================================

class ParseWorker(QThread):
    """Background thread for parsing a transcript."""
    finished = Signal(object)   # TranscriptIndex or None
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, pdf_path: str, force_reparse: bool = False):
        super().__init__()
        self.pdf_path = pdf_path
        self.force_reparse = force_reparse

    def run(self):
        try:
            from icharlotte_core.deposition.transcript_parser import TranscriptParser
            self.progress.emit("Parsing transcript...")
            parser = TranscriptParser()
            index = parser.parse(self.pdf_path, force_reparse=self.force_reparse)
            self.finished.emit(index)
        except Exception as e:
            logger.exception("Parse error")
            self.error.emit(str(e))


class SelectWorker(QThread):
    """Background thread for LLM testimony selection."""
    finished = Signal(object)   # ExtractionResult or None
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, index, prompt: str):
        super().__init__()
        self.index = index
        self.prompt = prompt

    def run(self):
        try:
            from icharlotte_core.deposition.testimony_selector import TestimonySelector
            self.progress.emit("Selecting relevant testimony...")
            selector = TestimonySelector()
            result = selector.select(self.index, self.prompt)
            self.finished.emit(result)
        except Exception as e:
            logger.exception("Selection error")
            self.error.emit(str(e))


# =============================================================================
# Depositions Tab
# =============================================================================

class DepositionTab(QWidget):
    """
    Split-panel deposition testimony extraction tab.

    Layout:
    ┌──────────────────────────────────────────────────────┐
    │  [Load] [▼ Transcript selector]  [Export] ☐ Highlight│
    ├──────────────────────┬───────────────────────────────┤
    │  Transcript Viewer   │  Extraction Results            │
    ├──────────────────────┴───────────────────────────────┤
    │  [Prompt input]                             [Extract] │
    └──────────────────────────────────────────────────────┘

    Supports drag-and-drop of PDF files to load transcripts.
    Multiple transcripts can be loaded and switched via the combo box.
    Loaded transcripts persist within a case across sessions.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_number = None
        self.setAcceptDrops(True)

        # Multi-transcript state
        # Each entry: {"path": str, "index": TranscriptIndex|None, "extractions": list}
        self._transcripts = []
        self._active_idx = -1

        # Workers
        self._parse_worker = None
        self._select_worker = None
        # Path of the transcript being parsed (to add to list on completion)
        self._pending_parse_path = None

        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.setSpacing(4)

        # --- Top Toolbar ---
        toolbar = QHBoxLayout()

        self.load_btn = QPushButton("Load Transcript")
        self.load_btn.clicked.connect(self.on_load_transcript)
        toolbar.addWidget(self.load_btn)

        self.transcript_combo = QComboBox()
        self.transcript_combo.setMinimumWidth(200)
        self.transcript_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.transcript_combo.setPlaceholderText("No transcripts loaded")
        self.transcript_combo.currentIndexChanged.connect(self._on_transcript_switched)
        toolbar.addWidget(self.transcript_combo)

        self.remove_btn = QPushButton("Remove")
        self.remove_btn.setToolTip("Remove the currently selected transcript")
        self.remove_btn.clicked.connect(self._on_remove_transcript)
        self.remove_btn.setEnabled(False)
        toolbar.addWidget(self.remove_btn)

        self.export_btn = QPushButton("Export to Word")
        self.export_btn.clicked.connect(self.on_export_word)
        self.export_btn.setEnabled(False)
        toolbar.addWidget(self.export_btn)

        self.highlight_cb = QCheckBox("Highlight text")
        self.highlight_cb.setToolTip(
            "When enabled, creates a highlighted PDF copy with extracted testimony marked in yellow"
        )
        toolbar.addWidget(self.highlight_cb)

        toolbar.addStretch()

        self.deponent_label = QLabel("")
        self.deponent_label.setStyleSheet("font-weight: bold; color: #444;")
        toolbar.addWidget(self.deponent_label)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #888;")
        toolbar.addWidget(self.status_label)

        main_layout.addLayout(toolbar)

        # --- Split Panel ---
        self.splitter = QSplitter(Qt.Horizontal)

        # Left: Transcript viewer
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_header = QLabel("Transcript")
        left_header.setStyleSheet("font-weight: bold; padding: 2px;")
        left_layout.addWidget(left_header)

        self.transcript_viewer = QTextEdit()
        self.transcript_viewer.setReadOnly(True)
        self.transcript_viewer.setFont(QFont("Consolas", 9))
        self.transcript_viewer.setPlaceholderText(
            "Load a deposition transcript PDF to begin.\n\n"
            "Click 'Load Transcript' or drag-and-drop a PDF file."
        )
        left_layout.addWidget(self.transcript_viewer)

        self.splitter.addWidget(left_container)

        # Right: Extraction results
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)

        right_header = QLabel("Extraction Results")
        right_header.setStyleSheet("font-weight: bold; padding: 2px;")
        right_layout.addWidget(right_header)

        self.results_viewer = QTextEdit()
        self.results_viewer.setReadOnly(True)
        self.results_viewer.setFont(QFont("Times New Roman", 11))
        self.results_viewer.setPlaceholderText(
            "Enter a prompt below and click 'Extract' to find relevant testimony.\n\n"
            "Examples:\n"
            "  - \"testimony about plaintiff's injuries prior to the accident\"\n"
            "  - \"employment history and job duties\"\n"
            "  - \"prior medical treatment for back and neck\""
        )
        right_layout.addWidget(self.results_viewer)

        self.splitter.addWidget(right_container)

        # Set initial splitter proportions (40/60)
        self.splitter.setSizes([400, 600])

        # --- Progress Bar ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(4)

        # --- Top container (content splitter + progress bar) ---
        top_container = QWidget()
        top_layout = QVBoxLayout(top_container)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        top_layout.addWidget(self.splitter, stretch=1)
        top_layout.addWidget(self.progress_bar)

        # --- Prompt Input Bar (bottom of vertical splitter) ---
        prompt_container = QWidget()
        prompt_layout = QHBoxLayout(prompt_container)
        prompt_layout.setContentsMargins(0, 0, 0, 0)

        self.prompt_input = QPlainTextEdit()
        self.prompt_input.setPlaceholderText(
            "Enter extraction prompt (e.g., 'testimony about prior accidents')  \u2014  Ctrl+Enter to submit"
        )
        self.prompt_input.setEnabled(False)
        self.prompt_input.setMinimumHeight(36)
        prompt_layout.addWidget(self.prompt_input, stretch=1)

        self.extract_btn = QPushButton("Extract")
        self.extract_btn.clicked.connect(self.on_extract)
        self.extract_btn.setEnabled(False)
        self.extract_btn.setMinimumWidth(100)
        prompt_layout.addWidget(self.extract_btn)

        # --- Vertical splitter (content | prompt) for resizable prompt ---
        self.v_splitter = QSplitter(Qt.Vertical)
        self.v_splitter.addWidget(top_container)
        self.v_splitter.addWidget(prompt_container)
        self.v_splitter.setSizes([700, 60])
        self.v_splitter.setStretchFactor(0, 1)  # Content stretches
        self.v_splitter.setStretchFactor(1, 0)  # Prompt stays fixed unless dragged

        main_layout.addWidget(self.v_splitter, stretch=1)

        # --- Persist layout sizes ---
        self.splitter.splitterMoved.connect(self._on_splitter_moved)
        self.v_splitter.splitterMoved.connect(self._on_v_splitter_moved)
        self._load_layout_settings()

    # =========================================================================
    # Keyboard & Layout Persistence
    # =========================================================================

    def keyPressEvent(self, event: QKeyEvent):
        """Ctrl+Enter submits extraction when prompt is focused."""
        if (self.prompt_input.hasFocus()
                and event.key() in (Qt.Key_Return, Qt.Key_Enter)
                and event.modifiers() & Qt.ControlModifier):
            self.on_extract()
            return
        super().keyPressEvent(event)

    def _on_splitter_moved(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue("deposition_tab/splitter_sizes", self.splitter.sizes())

    def _on_v_splitter_moved(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        settings.setValue("deposition_tab/v_splitter_sizes", self.v_splitter.sizes())

    def _load_layout_settings(self):
        settings = QSettings("iCharlotte", "iCharlotte")
        sizes = settings.value("deposition_tab/splitter_sizes")
        if sizes:
            self.splitter.setSizes([int(s) for s in sizes])
        v_sizes = settings.value("deposition_tab/v_splitter_sizes")
        if v_sizes:
            self.v_splitter.setSizes([int(s) for s in v_sizes])

    # =========================================================================
    # Drag and Drop
    # =========================================================================

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            # Accept if any URL is a PDF file
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith('.pdf'):
                    event.accept()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        pdf_paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith('.pdf'):
                pdf_paths.append(path)

        event.accept()

        # Defer processing to avoid blocking Explorer
        if pdf_paths:
            QTimer.singleShot(0, lambda: self._process_dropped_files(pdf_paths))

    def _process_dropped_files(self, pdf_paths: list):
        """Load dropped PDF files as transcripts."""
        for path in pdf_paths:
            # Skip if already loaded
            if any(t["path"] == path for t in self._transcripts):
                logger.info(f"Transcript already loaded: {path}")
                continue
            self._load_transcript(path)
            # Only load one at a time (parser is async); rest will need
            # manual loading. For bulk drops, queue them.
            break

        # Queue remaining paths if any
        if len(pdf_paths) > 1:
            remaining = pdf_paths[1:]
            # Filter out already-loaded and the one we just started
            remaining = [p for p in remaining
                         if not any(t["path"] == p for t in self._transcripts)
                         and p != pdf_paths[0]]
            if remaining:
                self._drop_queue = remaining

    # =========================================================================
    # Persistence
    # =========================================================================

    def _state_path(self) -> str:
        """Path to the persistence JSON for the current case."""
        os.makedirs(GEMINI_DATA_DIR, exist_ok=True)
        return os.path.join(GEMINI_DATA_DIR, f"{self.file_number}_depo.json")

    def _save_state(self):
        """Save transcript list and active index to disk."""
        if not self.file_number:
            return
        data = {
            "version": "1.0",
            "transcripts": [
                {"pdf_path": t["path"], "label": self._label_for(t)}
                for t in self._transcripts
            ],
            "active_index": self._active_idx
        }
        try:
            with open(self._state_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            logger.error(f"Failed to save deposition state: {e}")

    def _load_state(self):
        """Load transcript list from disk and restore cached indexes."""
        path = self._state_path()
        if not os.path.exists(path):
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            logger.error(f"Failed to load deposition state: {e}")
            return

        saved_active = data.get("active_index", 0)

        for entry in data.get("transcripts", []):
            pdf_path = entry.get("pdf_path", "")
            if not pdf_path or not os.path.exists(pdf_path):
                logger.warning(f"Skipping missing transcript: {pdf_path}")
                continue

            # Try to load cached index (fast — reads .depo_index.json)
            index = self._try_load_cached_index(pdf_path)
            self._transcripts.append({
                "path": pdf_path,
                "index": index,
                "extractions": []
            })
            label = entry.get("label", os.path.basename(pdf_path))
            self.transcript_combo.addItem(label)

        # Activate saved transcript
        if self._transcripts:
            active = min(saved_active, len(self._transcripts) - 1)
            active = max(0, active)
            self.transcript_combo.setCurrentIndex(active)

    def _try_load_cached_index(self, pdf_path: str):
        """Try to load a cached TranscriptIndex without full parsing."""
        try:
            from icharlotte_core.deposition.transcript_parser import TranscriptParser
            parser = TranscriptParser()
            # parse() checks cache first — returns instantly if cached
            return parser.parse(pdf_path)
        except Exception as e:
            logger.warning(f"Could not load cached index for {pdf_path}: {e}")
            return None

    @staticmethod
    def _label_for(transcript: dict) -> str:
        """Generate a combo box label for a transcript."""
        index = transcript.get("index")
        if index and index.deponent and index.deponent.last_name:
            dep = index.deponent
            parts = [dep.last_name]
            if dep.deposition_date:
                parts.append(dep.deposition_date)
            return " — ".join(parts)
        return os.path.basename(transcript["path"])

    # =========================================================================
    # Case Switching
    # =========================================================================

    def load_case(self, file_number: str):
        """Called when the active case changes."""
        # Save current case state
        if self.file_number and self._transcripts:
            self._save_state()

        self.file_number = file_number

        # Clear all state
        self._transcripts.clear()
        self._active_idx = -1
        self.transcript_combo.blockSignals(True)
        self.transcript_combo.clear()
        self.transcript_combo.blockSignals(False)
        self.transcript_viewer.clear()
        self.results_viewer.clear()
        self.deponent_label.setText("")
        self.status_label.setText("")
        self.prompt_input.setEnabled(False)
        self.extract_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.remove_btn.setEnabled(False)

        # Load persisted state for new case
        if file_number:
            self._load_state()

    # =========================================================================
    # Transcript Switching
    # =========================================================================

    def _on_transcript_switched(self, idx: int):
        """Handle combo box selection change."""
        if idx < 0 or idx >= len(self._transcripts):
            return

        self._active_idx = idx
        entry = self._transcripts[idx]
        index = entry["index"]

        # Clear viewers
        self.transcript_viewer.clear()
        self.results_viewer.clear()

        if index:
            self._show_transcript(index)
            # Re-display any extractions for this transcript
            for _, result in entry["extractions"]:
                self._append_results(index, result)
        else:
            # Index not loaded yet — trigger parse
            self._load_transcript(entry["path"])
            return

        self.remove_btn.setEnabled(True)
        self._save_state()

    def _show_transcript(self, index):
        """Populate viewers for a loaded transcript index."""
        # Deponent label
        dep = index.deponent
        label_parts = []
        if dep.full_name:
            label_parts.append(dep.full_name)
        if dep.deposition_date:
            label_parts.append(dep.deposition_date)
        self.deponent_label.setText(" | ".join(label_parts))

        # Populate transcript viewer
        self._populate_transcript_viewer(index)

        # Enable prompt input
        self.prompt_input.setEnabled(True)
        self.extract_btn.setEnabled(True)

        self.status_label.setText(
            f"{len(index.exchanges)} Q/A exchanges | "
            f"{index.total_transcript_pages} pages | "
            f"{'Condensed' if index.is_condensed else 'Full-size'}"
        )

        # Enable export if any extractions exist
        has_extractions = any(t["extractions"] for t in self._transcripts)
        self.export_btn.setEnabled(has_extractions)
        self.remove_btn.setEnabled(True)

    # =========================================================================
    # Actions
    # =========================================================================

    def on_load_transcript(self):
        """Open file dialog to select a transcript PDF."""
        start_dir = ""
        if self.file_number:
            main_win = self.window()
            if main_win and hasattr(main_win, 'case_root_path'):
                case_root = main_win.case_root_path
                if case_root:
                    transcript_dir = os.path.join(case_root, "DISCOVERY", "TRANSCRIPTS")
                    if os.path.isdir(transcript_dir):
                        start_dir = transcript_dir
                    else:
                        start_dir = case_root

        path, _ = QFileDialog.getOpenFileName(
            self, "Select Deposition Transcript", start_dir,
            "PDF Files (*.pdf);;All Files (*)"
        )
        if not path:
            return

        # Skip if already loaded
        if any(t["path"] == path for t in self._transcripts):
            # Just switch to it
            for i, t in enumerate(self._transcripts):
                if t["path"] == path:
                    self.transcript_combo.setCurrentIndex(i)
                    return
            return

        self._load_transcript(path)

    def _load_transcript(self, pdf_path: str):
        """Start parsing a transcript."""
        if self._parse_worker and self._parse_worker.isRunning():
            QMessageBox.warning(self, "Busy", "A transcript is already being parsed.")
            return

        self.status_label.setText("Parsing...")
        self.progress_bar.setVisible(True)
        self.load_btn.setEnabled(False)

        self._pending_parse_path = pdf_path

        self._parse_worker = ParseWorker(pdf_path)
        self._parse_worker.finished.connect(self._on_parse_finished)
        self._parse_worker.error.connect(self._on_parse_error)
        self._parse_worker.start()

    def _on_parse_finished(self, index):
        """Handle successful transcript parsing."""
        self.progress_bar.setVisible(False)
        self.load_btn.setEnabled(True)

        pdf_path = self._pending_parse_path
        self._pending_parse_path = None

        # Check if this transcript is already in the list (restored from
        # persistence but index was None)
        existing_idx = None
        for i, t in enumerate(self._transcripts):
            if t["path"] == pdf_path:
                existing_idx = i
                break

        if existing_idx is not None:
            # Update existing entry with parsed index
            self._transcripts[existing_idx]["index"] = index
            label = self._label_for(self._transcripts[existing_idx])
            self.transcript_combo.setItemText(existing_idx, label)
            self._active_idx = existing_idx
        else:
            # Add new transcript
            entry = {"path": pdf_path, "index": index, "extractions": []}
            self._transcripts.append(entry)
            label = self._label_for(entry)
            self.transcript_combo.blockSignals(True)
            self.transcript_combo.addItem(label)
            new_idx = len(self._transcripts) - 1
            self.transcript_combo.setCurrentIndex(new_idx)
            self.transcript_combo.blockSignals(False)
            self._active_idx = new_idx

        # Show the transcript
        self._show_transcript(index)
        self.prompt_input.setFocus()

        # Persist
        self._save_state()

        # Process queued drops if any
        if hasattr(self, '_drop_queue') and self._drop_queue:
            next_path = self._drop_queue.pop(0)
            if not any(t["path"] == next_path for t in self._transcripts):
                QTimer.singleShot(100, lambda: self._load_transcript(next_path))

    def _on_parse_error(self, error_msg: str):
        """Handle transcript parsing failure."""
        self.progress_bar.setVisible(False)
        self.load_btn.setEnabled(True)
        self._pending_parse_path = None
        self.status_label.setText("Parse failed")
        QMessageBox.critical(self, "Parse Error", f"Failed to parse transcript:\n\n{error_msg}")

    def _on_remove_transcript(self):
        """Remove the currently selected transcript."""
        if self._active_idx < 0 or self._active_idx >= len(self._transcripts):
            return

        self._transcripts.pop(self._active_idx)
        self.transcript_combo.blockSignals(True)
        self.transcript_combo.removeItem(self._active_idx)
        self.transcript_combo.blockSignals(False)

        if self._transcripts:
            new_idx = min(self._active_idx, len(self._transcripts) - 1)
            self.transcript_combo.setCurrentIndex(new_idx)
            self._on_transcript_switched(new_idx)
        else:
            self._active_idx = -1
            self.transcript_viewer.clear()
            self.results_viewer.clear()
            self.deponent_label.setText("")
            self.status_label.setText("")
            self.prompt_input.setEnabled(False)
            self.extract_btn.setEnabled(False)
            self.export_btn.setEnabled(False)
            self.remove_btn.setEnabled(False)

        self._save_state()

    def _populate_transcript_viewer(self, index):
        """Fill the left panel with parsed transcript text."""
        self.transcript_viewer.clear()

        if index.raw_text:
            self.transcript_viewer.setPlainText(index.raw_text)
        else:
            # Build from exchanges
            lines = []
            for ex in index.exchanges:
                lines.append(f"p.{ex.page_start}:{ex.line_start}")
                lines.append(f"  Q.  {ex.question}")
                lines.append(f"  A.  {ex.answer}")
                lines.append("")
            self.transcript_viewer.setPlainText("\n".join(lines))

    def on_extract(self):
        """Start LLM extraction with the current prompt."""
        prompt = self.prompt_input.toPlainText().strip()
        if not prompt:
            return

        if self._active_idx < 0:
            QMessageBox.warning(self, "No Transcript", "Please load a transcript first.")
            return

        active = self._transcripts[self._active_idx]
        index = active.get("index")
        if not index:
            QMessageBox.warning(self, "No Transcript", "Transcript is still loading.")
            return

        if self._select_worker and self._select_worker.isRunning():
            QMessageBox.warning(self, "Busy", "An extraction is already in progress.")
            return

        self.status_label.setText("Extracting...")
        self.progress_bar.setVisible(True)
        self.extract_btn.setEnabled(False)
        self.prompt_input.setEnabled(False)

        self._select_worker = SelectWorker(index, prompt)
        self._select_worker.finished.connect(self._on_select_finished)
        self._select_worker.error.connect(self._on_select_error)
        self._select_worker.start()

    def _on_select_finished(self, result):
        """Handle successful testimony selection."""
        self.progress_bar.setVisible(False)
        self.extract_btn.setEnabled(True)
        self.prompt_input.setEnabled(True)

        if not result or not result.selected_ids:
            self.status_label.setText("No relevant testimony found")
            QMessageBox.information(
                self, "No Results",
                "The LLM did not find any testimony relevant to your prompt.\n"
                "Try broadening your search or rephrasing."
            )
            return

        # Store extraction on the active transcript
        if 0 <= self._active_idx < len(self._transcripts):
            active = self._transcripts[self._active_idx]
            index = active["index"]
            active["extractions"].append((index, result))
            self.export_btn.setEnabled(True)

            # Display results
            self._append_results(index, result)

            # Highlight transcript viewer
            self._highlight_viewer(index, result)

            # PDF highlighting if enabled
            if self.highlight_cb.isChecked():
                try:
                    from icharlotte_core.deposition.testimony_formatter import TestimonyFormatter
                    formatter = TestimonyFormatter()
                    highlight_path = formatter.highlight_pdf(index, result)
                    if highlight_path:
                        self.status_label.setText(
                            f"Extracted {len(result.selected_ids)} exchanges | "
                            f"Highlighted: {os.path.basename(highlight_path)}"
                        )
                        return
                except Exception as e:
                    logger.exception("PDF highlighting failed")

            self.status_label.setText(
                f"Extracted {len(result.selected_ids)} exchanges in "
                f"{len(result.groups)} groups"
            )

    def _on_select_error(self, error_msg: str):
        """Handle extraction failure."""
        self.progress_bar.setVisible(False)
        self.extract_btn.setEnabled(True)
        self.prompt_input.setEnabled(True)
        self.status_label.setText("Extraction failed")
        QMessageBox.critical(self, "Extraction Error", f"LLM extraction failed:\n\n{error_msg}")

    def _append_results(self, index, result):
        """Append extraction results to the right panel."""
        cursor = self.results_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)

        # Separator if not first extraction in viewer
        if self.results_viewer.toPlainText().strip():
            cursor.insertText("\n" + "=" * 60 + "\n\n")

        # Prompt header
        header_fmt = QTextCharFormat()
        header_fmt.setFontWeight(QFont.Bold)
        header_fmt.setFontPointSize(12)
        cursor.insertText(f"Extraction: {result.prompt}\n", header_fmt)

        # Normal formatting
        normal_fmt = QTextCharFormat()
        normal_fmt.setFontPointSize(11)
        normal_fmt.setFontFamily("Times New Roman")

        marker_fmt = QTextCharFormat()
        marker_fmt.setFontPointSize(11)
        marker_fmt.setFontFamily("Times New Roman")
        marker_fmt.setFontWeight(QFont.Bold)

        cite_fmt = QTextCharFormat()
        cite_fmt.setFontPointSize(10)
        cite_fmt.setFontFamily("Times New Roman")
        cite_fmt.setForeground(QColor("#555"))

        exchange_map = {ex.id: ex for ex in index.exchanges}

        for group_idx, group_ids in enumerate(result.groups):
            group_exchanges = [exchange_map[eid] for eid in group_ids if eid in exchange_map]
            if not group_exchanges:
                continue

            if group_idx > 0:
                cursor.insertText("\n")

            for ex in group_exchanges:
                cursor.insertText("Q.\t", marker_fmt)
                cursor.insertText(f"{ex.question}\n", normal_fmt)
                cursor.insertText("A.\t", marker_fmt)
                cursor.insertText(f"{ex.answer}\n", normal_fmt)

            # Citation
            ranges = [ex.citation_range() for ex in group_exchanges]
            range_str = "; ".join(ranges)
            last_name = index.deponent.last_name or "___"
            citation = f"\n(Exh. __ ({last_name} Depo. Trns.) at p. {range_str}.)\n\n"
            cursor.insertText(citation, cite_fmt)

        self.results_viewer.setTextCursor(cursor)
        self.results_viewer.ensureCursorVisible()

    def _highlight_viewer(self, index, result):
        """Highlight extracted exchanges in the transcript viewer."""
        if not index:
            return

        exchange_map = {ex.id: ex for ex in index.exchanges}
        cursor = self.transcript_viewer.textCursor()

        highlight_fmt = QTextCharFormat()
        highlight_fmt.setBackground(QColor("#FFFF99"))

        for eid in result.selected_ids:
            ex = exchange_map.get(eid)
            if not ex:
                continue

            # Search for the Q text in the viewer
            for search_text in [ex.question[:80], ex.answer[:80]]:
                cursor = self.transcript_viewer.document().find(search_text)
                if not cursor.isNull():
                    cursor.mergeCharFormat(highlight_fmt)

    def on_export_word(self):
        """Export all extractions to a Word document."""
        # Gather extractions from all transcripts
        all_extractions = []
        for t in self._transcripts:
            all_extractions.extend(t["extractions"])

        if not all_extractions:
            return

        from icharlotte_core.deposition.testimony_formatter import TestimonyFormatter

        formatter = TestimonyFormatter()
        for index, result in all_extractions:
            formatter.add_extraction(index, result)

        # Determine default path
        first_index = all_extractions[0][0]
        default_dir = os.path.dirname(first_index.source_pdf)
        deponent = first_index.deponent.last_name or "Unknown"
        default_name = f"[Extracted] {deponent} Depo Trns.docx"
        default_path = os.path.join(default_dir, default_name)

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Extraction Document", default_path,
            "Word Documents (*.docx);;All Files (*)"
        )
        if not path:
            return

        try:
            output_dir = os.path.dirname(path)
            filename = os.path.basename(path)
            saved = formatter.save_word(output_dir, filename)
            self.status_label.setText(f"Saved: {os.path.basename(saved)}")
            QMessageBox.information(self, "Saved", f"Extraction document saved to:\n{saved}")
        except Exception as e:
            logger.exception("Export error")
            QMessageBox.critical(self, "Export Error", f"Failed to save document:\n\n{e}")
