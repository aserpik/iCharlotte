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
    QPushButton, QTextEdit, QPlainTextEdit,
    QFileDialog, QMessageBox, QProgressBar, QApplication,
    QToolBar, QSizePolicy, QComboBox
)
from PySide6.QtCore import Qt, Signal, QThread, QTimer, QSettings
from PySide6.QtGui import (
    QTextCursor, QFont, QTextCharFormat, QTextBlockFormat,
    QColor, QDragEnterEvent, QDropEvent, QKeyEvent
)

from icharlotte_core.ui.pdf_viewer_widget import PdfViewerWidget

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
    progress = Signal(str)      # Phase description
    chunk_progress = Signal(int, int)  # (current_chunk, total_chunks)

    def __init__(self, index, prompt: str):
        super().__init__()
        self.index = index
        self.prompt = prompt

    def run(self):
        try:
            from icharlotte_core.deposition.testimony_selector import TestimonySelector
            selector = TestimonySelector()

            # Calculate chunks to report progress
            exchanges = self.index.exchanges
            max_per_chunk = 500
            import math
            num_chunks = max(1, math.ceil(len(exchanges) / max_per_chunk))

            self.progress.emit(f"Analyzing {len(exchanges)} exchanges...")
            self.chunk_progress.emit(0, num_chunks)

            result = selector.select(
                self.index, self.prompt,
                on_chunk_done=lambda cur, total: self.chunk_progress.emit(cur, total)
            )
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

        # Map cursor positions in results viewer to transcript page numbers
        # List of (start_pos, end_pos, transcript_page) tuples
        self._result_page_anchors = []
        # Map cursor positions to exchange IDs for deletion
        # List of (start_pos, end_pos, [exchange_ids], transcript_idx, extraction_idx)
        self._result_exchange_anchors = []

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

        self.reparse_btn = QPushButton("Reparse")
        self.reparse_btn.setToolTip("Re-extract text from the transcript PDF (clears cache)")
        self.reparse_btn.clicked.connect(self._on_reparse)
        self.reparse_btn.setEnabled(False)
        toolbar.addWidget(self.reparse_btn)

        self.export_btn = QPushButton("Export to Word")
        self.export_btn.clicked.connect(self.on_export_word)
        self.export_btn.setEnabled(False)
        toolbar.addWidget(self.export_btn)

        self.clear_results_btn = QPushButton("Clear Results")
        self.clear_results_btn.setToolTip("Clear all extraction results")
        self.clear_results_btn.clicked.connect(self._on_clear_results)
        self.clear_results_btn.setEnabled(False)
        toolbar.addWidget(self.clear_results_btn)

        # Highlight color picker
        self._highlight_colors = [
            ("#FFCC00", "Yellow"),
            ("#FF6666", "Red"),
            ("#66CC66", "Green"),
            ("#6699FF", "Blue"),
            ("#FF9933", "Orange"),
            ("#FF99CC", "Pink"),
        ]
        self.color_btn = QPushButton()
        self.color_btn.setFixedSize(28, 28)
        self.color_btn.setToolTip("Highlight color for next extraction")
        self.color_btn.clicked.connect(self._on_pick_color)
        toolbar.addWidget(self.color_btn)
        settings = QSettings("iCharlotte", "iCharlotte")
        self._current_color = settings.value(
            "deposition_tab/highlight_color", "#FFCC00"
        )
        self._update_color_btn()

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

        # Left: PDF transcript viewer
        self.transcript_viewer = PdfViewerWidget()
        self.transcript_viewer.setAcceptDrops(False)
        self.transcript_viewer.web_view.setAcceptDrops(False)
        self.transcript_viewer.annotationDeleteRequested.connect(
            self._on_annotation_delete_requested
        )
        self.splitter.addWidget(self.transcript_viewer)

        # Right: Extraction results
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)

        right_header = QLabel("Extraction Results")
        right_header.setStyleSheet("font-weight: bold; padding: 2px;")
        right_layout.addWidget(right_header)

        self.results_viewer = QTextEdit()
        self.results_viewer.setReadOnly(False)
        self.results_viewer.setFont(QFont("Times New Roman", 11))
        self.results_viewer.setTabStopDistance(40)  # Tab aligns with left margin indent
        self.results_viewer.setContextMenuPolicy(Qt.CustomContextMenu)
        self.results_viewer.customContextMenuRequested.connect(
            self._on_results_context_menu
        )
        self.results_viewer.setPlaceholderText(
            "Enter a prompt below and click 'Extract' to find relevant testimony.\n\n"
            "Examples:\n"
            "  - \"testimony about plaintiff's injuries prior to the accident\"\n"
            "  - \"employment history and job duties\"\n"
            "  - \"prior medical treatment for back and neck\""
        )
        self.results_viewer.cursorPositionChanged.connect(self._on_results_clicked)
        right_layout.addWidget(self.results_viewer)

        self.splitter.addWidget(right_container)

        # Set initial splitter proportions (40/60)
        self.splitter.setSizes([400, 600])

        # --- Progress / Status Bar (between panes and prompt) ---
        progress_container = QWidget()
        progress_layout = QHBoxLayout(progress_container)
        progress_layout.setContentsMargins(0, 2, 0, 2)
        progress_layout.setSpacing(6)

        self.progress_label = QLabel("")
        self.progress_label.setStyleSheet("font-size: 11px; color: #555;")
        self.progress_label.setVisible(False)
        progress_layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(12)
        self.progress_bar.setMinimumWidth(150)
        self.progress_bar.setTextVisible(False)
        progress_layout.addWidget(self.progress_bar)

        progress_layout.addStretch()

        # --- Top container (content splitter + progress bar) ---
        top_container = QWidget()
        top_layout = QVBoxLayout(top_container)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        top_layout.addWidget(self.splitter, stretch=1)
        top_layout.addWidget(progress_container)

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
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith('.pdf'):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragMoveEvent(self, event):
        """Accept drag move so the drop indicator stays active everywhere."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        pdf_paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith('.pdf'):
                pdf_paths.append(path)

        event.acceptProposedAction()

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
        self._result_page_anchors.clear()
        self._result_exchange_anchors.clear()

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
        self.reparse_btn.setEnabled(True)
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
        self.reparse_btn.setEnabled(True)

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
        self._show_progress("Parsing transcript...")
        self.load_btn.setEnabled(False)

        self._pending_parse_path = pdf_path

        self._parse_worker = ParseWorker(pdf_path)
        self._parse_worker.finished.connect(self._on_parse_finished)
        self._parse_worker.error.connect(self._on_parse_error)
        self._parse_worker.start()

    def _on_parse_finished(self, index):
        """Handle successful transcript parsing."""
        self._hide_progress()
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
        self._hide_progress()
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

    def _on_reparse(self):
        """Force reparse of the current transcript (clears cached index)."""
        if self._active_idx < 0 or self._active_idx >= len(self._transcripts):
            return
        active = self._transcripts[self._active_idx]
        pdf_path = active["path"]

        # Clear existing extractions and results
        active["extractions"].clear()
        active["index"] = None
        self.results_viewer.clear()
        self._result_page_anchors.clear()
        self._result_exchange_anchors.clear()

        # Delete cached index file if it exists
        from icharlotte_core.deposition.models import TranscriptIndex
        cache_path = TranscriptIndex.cache_path_for(pdf_path)
        if os.path.exists(cache_path):
            os.remove(cache_path)
            logger.info(f"Deleted cached index: {cache_path}")

        # Reload with force_reparse
        self.status_label.setText("Reparsing...")
        self._show_progress("Reparsing transcript...")
        self.load_btn.setEnabled(False)
        self._pending_parse_path = pdf_path

        self._parse_worker = ParseWorker(pdf_path, force_reparse=True)
        self._parse_worker.finished.connect(self._on_parse_finished)
        self._parse_worker.error.connect(self._on_parse_error)
        self._parse_worker.start()

    def _populate_transcript_viewer(self, index):
        """Load the transcript PDF into the left panel viewer."""
        if index and index.source_pdf and os.path.isfile(index.source_pdf):
            self.transcript_viewer.load_pdf(index.source_pdf)

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
        self._show_progress("Sending to LLM...", 0, 100)
        self.extract_btn.setEnabled(False)
        self.prompt_input.setEnabled(False)

        self._select_worker = SelectWorker(index, prompt)
        self._select_worker.finished.connect(self._on_select_finished)
        self._select_worker.error.connect(self._on_select_error)
        self._select_worker.progress.connect(self._on_extract_progress)
        self._select_worker.chunk_progress.connect(self._on_chunk_progress)
        self._select_worker.start()

    def _on_select_finished(self, result):
        """Handle successful testimony selection."""
        self.extract_btn.setEnabled(True)
        self.prompt_input.setEnabled(True)

        if not result or not result.selected_ids:
            self._hide_progress()
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
            result.highlight_color = self._current_color
            active["extractions"].append((index, result))
            self.export_btn.setEnabled(True)
            self.clear_results_btn.setEnabled(True)

            # Phase 2: Format results (70-85%)
            num_selected = len(result.selected_ids)
            self._show_progress(
                f"Formatting {num_selected} exchanges...", 72, 100
            )
            QApplication.processEvents()
            self._append_results(index, result)
            self._show_progress("Formatting complete", 85, 100)
            QApplication.processEvents()

            # Phase 3: PDF highlighting (85-100%)
            try:
                from icharlotte_core.deposition.testimony_formatter import TestimonyFormatter
                formatter = TestimonyFormatter()
                ai_dir = self._get_ai_output_dir()
                self._show_progress(
                    f"Highlighting {num_selected} exchanges in PDF...", 86, 100
                )
                QApplication.processEvents()
                highlight_path = formatter.highlight_pdf(
                    index, result, output_dir=ai_dir,
                    on_progress=lambda cur, total: self._on_highlight_progress(cur, total),
                    color=result.highlight_color,
                )
                if highlight_path:
                    self._show_progress("Loading highlighted PDF...", 97, 100)
                    QApplication.processEvents()
                    self._load_highlighted_pdf(highlight_path)
            except Exception as e:
                logger.exception("PDF highlighting failed")

            self._hide_progress()
            self.status_label.setText(
                f"Extracted {num_selected} exchanges in "
                f"{len(result.groups)} groups"
            )

    def _on_select_error(self, error_msg: str):
        """Handle extraction failure."""
        self._hide_progress()
        self.extract_btn.setEnabled(True)
        self.prompt_input.setEnabled(True)
        self.status_label.setText("Extraction failed")
        QMessageBox.critical(self, "Extraction Error", f"LLM extraction failed:\n\n{error_msg}")

    def _show_progress(self, message: str, value: int = -1, maximum: int = 100):
        """Show progress bar with label. value=-1 for indeterminate."""
        self.progress_label.setText(message)
        self.progress_label.setVisible(True)
        self.progress_bar.setVisible(True)
        if value < 0:
            self.progress_bar.setRange(0, 0)  # Indeterminate animation
        else:
            self.progress_bar.setRange(0, maximum)
            self.progress_bar.setValue(value)

    def _hide_progress(self):
        """Hide progress bar and label."""
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)

    def _on_extract_progress(self, message: str):
        """Update progress label during extraction."""
        self.progress_label.setText(message)

    def _on_highlight_progress(self, current: int, total: int):
        """Update progress bar during PDF highlighting (85-97% of total bar)."""
        pct = 85 + int((current / max(total, 1)) * 12)
        self._show_progress(
            f"Highlighting exchange {current}/{total}...", pct, 100
        )
        QApplication.processEvents()

    def _on_chunk_progress(self, current: int, total: int):
        """Update progress bar with chunk completion (0-70% of total bar)."""
        # LLM selection uses 0-70% of the progress bar
        pct = int((current / max(total, 1)) * 70)
        if total > 1:
            self._show_progress(f"LLM analyzing chunk {current}/{total}...", pct, 100)
        else:
            self._show_progress("LLM analyzing testimony...", pct, 100)

    def _update_color_btn(self):
        """Update the color button's background to reflect current color."""
        self.color_btn.setStyleSheet(
            f"background-color: {self._current_color}; "
            "border: 1px solid #888; border-radius: 4px;"
        )

    def _on_pick_color(self):
        """Show a popup menu with preset highlight colors."""
        from PySide6.QtWidgets import QMenu
        from PySide6.QtGui import QPixmap, QIcon
        menu = QMenu(self)
        for hex_color, name in self._highlight_colors:
            pixmap = QPixmap(16, 16)
            pixmap.fill(QColor(hex_color))
            action = menu.addAction(QIcon(pixmap), name)
            action.setData(hex_color)
        chosen = menu.exec(
            self.color_btn.mapToGlobal(self.color_btn.rect().bottomLeft())
        )
        if chosen:
            self._current_color = chosen.data()
            self._update_color_btn()
            settings = QSettings("iCharlotte", "iCharlotte")
            settings.setValue("deposition_tab/highlight_color", self._current_color)

    def _on_clear_results(self):
        """Clear all extraction results."""
        self.results_viewer.clear()
        self._result_page_anchors.clear()
        self._result_exchange_anchors.clear()
        self.clear_results_btn.setEnabled(False)
        # Clear stored extractions on all transcripts
        for t in self._transcripts:
            t["extractions"].clear()
        self.export_btn.setEnabled(False)
        self.status_label.setText("Results cleared")

    def _on_results_clicked(self):
        """Navigate PDF viewer to the page corresponding to clicked results text."""
        cursor = self.results_viewer.textCursor()
        pos = cursor.position()
        for start, end, page_num in self._result_page_anchors:
            if start <= pos <= end:
                self.transcript_viewer.go_to_page(page_num)
                break

    def _on_results_context_menu(self, pos):
        """Show context menu with option to delete a highlight group."""
        cursor = self.results_viewer.cursorForPosition(pos)
        click_pos = cursor.position()

        target = None
        for anchor in self._result_exchange_anchors:
            start, end, ex_ids, transcript_idx, extraction_idx = anchor
            if start <= click_pos <= end:
                target = anchor
                break

        from PySide6.QtWidgets import QMenu
        menu = self.results_viewer.createStandardContextMenu()
        if target:
            menu.addSeparator()
            delete_action = menu.addAction("Delete this highlight group")
        else:
            delete_action = None

        chosen = menu.exec(self.results_viewer.mapToGlobal(pos))
        if chosen and chosen == delete_action:
            self._delete_highlight_group(target)

    def _delete_highlight_group(self, anchor):
        """Delete a highlight group from results viewer and PDF annotations."""
        start, end, ex_ids, transcript_idx, extraction_idx = anchor

        # 1. Remove annotations from [H.AI] PDF
        active = self._transcripts[transcript_idx]
        index = active["index"]
        ai_dir = self._get_ai_output_dir()
        source_name = os.path.basename(index.source_pdf)
        highlight_path = os.path.join(ai_dir, f"[H.AI] {source_name}")

        if os.path.isfile(highlight_path):
            import fitz
            import tempfile
            doc = fitz.open(highlight_path)
            ex_id_set = {str(eid) for eid in ex_ids}
            for page in doc:
                annots_to_delete = []
                for annot in (page.annots() or []):
                    content = annot.info.get("content", "")
                    if content.startswith("ex:") and content[3:] in ex_id_set:
                        annots_to_delete.append(annot)
                for annot in annots_to_delete:
                    page.delete_annot(annot)
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pdf", dir=ai_dir)
            os.close(tmp_fd)
            doc.save(tmp_path, garbage=3)
            doc.close()
            os.replace(tmp_path, highlight_path)
            self._load_highlighted_pdf(highlight_path)

        # 2. Remove text from results viewer
        cursor = self.results_viewer.textCursor()
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.KeepAnchor)
        cursor.removeSelectedText()

        # 3. Adjust anchor positions
        removed_length = end - start
        self._result_exchange_anchors.remove(anchor)

        # Remove matching page anchor and adjust remaining positions
        self._result_page_anchors = [
            (s - removed_length if s > start else s,
             e - removed_length if e > start else e, p)
            for s, e, p in self._result_page_anchors
            if not (s >= start and e <= end)
        ]
        self._result_exchange_anchors = [
            (s - removed_length if s > start else s,
             e - removed_length if e > start else e, ids, ti, ei)
            for s, e, ids, ti, ei in self._result_exchange_anchors
        ]

        # 4. Remove exchange IDs from the extraction result
        if 0 <= extraction_idx < len(active["extractions"]):
            _, ext_result = active["extractions"][extraction_idx]
            ex_id_set_int = set(ex_ids)
            ext_result.selected_ids = [
                eid for eid in ext_result.selected_ids if eid not in ex_id_set_int
            ]
            ext_result.groups = [
                g for g in ext_result.groups
                if not all(eid in ex_id_set_int for eid in g)
            ]

        self.status_label.setText(f"Deleted {len(ex_ids)} exchanges")

    def _on_annotation_delete_requested(self, ex_id_str: str):
        """Handle Delete key on a selected highlight in the PDF viewer."""
        try:
            ex_id = int(ex_id_str)
        except ValueError:
            return

        # Find the anchor that contains this exchange ID
        target = None
        for anchor in self._result_exchange_anchors:
            start, end, ex_ids, transcript_idx, extraction_idx = anchor
            if ex_id in ex_ids:
                target = anchor
                break

        if target:
            self._delete_highlight_group(target)

    def _append_results(self, index, result):
        """Append extraction results to the right panel."""
        cursor = self.results_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)

        # Separator if not first extraction in viewer
        if self.results_viewer.toPlainText().strip():
            cursor.insertText("\n" + "=" * 60 + "\n\n")

        # Color indicator + prompt header
        color = getattr(result, "highlight_color", "#FFCC00")
        color_fmt = QTextCharFormat()
        color_fmt.setBackground(QColor(color))
        color_fmt.setFontPointSize(12)
        color_fmt.setFontWeight(QFont.Bold)
        cursor.insertText(" \u2588 ", color_fmt)

        header_fmt = QTextCharFormat()
        header_fmt.setFontWeight(QFont.Bold)
        header_fmt.setFontPointSize(12)
        cursor.insertText(f" Extraction: {result.prompt}\n", header_fmt)

        # Normal formatting
        normal_fmt = QTextCharFormat()
        normal_fmt.setFontPointSize(11)
        normal_fmt.setFontFamily("Times New Roman")

        marker_fmt = QTextCharFormat()
        marker_fmt.setFontPointSize(11)
        marker_fmt.setFontFamily("Times New Roman")
        marker_fmt.setFontWeight(QFont.Bold)

        # Block format for Q/A paragraphs — hanging indent so continuation lines align
        qa_block_fmt = QTextBlockFormat()
        qa_block_fmt.setLeftMargin(40)      # Text body indent
        qa_block_fmt.setTextIndent(-40)     # Marker hangs left

        cite_fmt = QTextCharFormat()
        cite_fmt.setFontPointSize(10)
        cite_fmt.setFontFamily("Times New Roman")
        cite_fmt.setForeground(QColor("#555"))

        exchange_map = {ex.id: ex for ex in index.exchanges}

        # Determine extraction index for anchor tracking
        active = self._transcripts[self._active_idx]
        extraction_idx = len(active["extractions"]) - 1

        for group_idx, group_ids in enumerate(result.groups):
            group_exchanges = [exchange_map[eid] for eid in group_ids if eid in exchange_map]
            if not group_exchanges:
                continue

            if group_idx > 0:
                cursor.insertText("\n")

            # Record position anchor for click-to-navigate
            group_start_pos = cursor.position()
            first_page = group_exchanges[0].page_start

            for ex in group_exchanges:
                cursor.setBlockFormat(qa_block_fmt)
                cursor.insertText("Q.\t", marker_fmt)
                cursor.insertText(f"{ex.question}\n", normal_fmt)
                cursor.setBlockFormat(qa_block_fmt)
                cursor.insertText("A.\t", marker_fmt)
                cursor.insertText(f"{ex.answer}\n", normal_fmt)

            # Reset block format for citation
            cursor.setBlockFormat(QTextBlockFormat())

            # Citation — merge consecutive exchanges into single range
            first = group_exchanges[0]
            last = group_exchanges[-1]
            if first.page_start == last.page_end:
                range_str = f"{first.page_start}:{first.line_start}-{last.line_end}"
            else:
                range_str = f"{first.page_start}:{first.line_start}-{last.page_end}:{last.line_end}"
            raw_name = index.deponent.last_name or "___"
            last_name = raw_name.title() if raw_name.isupper() else raw_name
            citation = f"\n(Exh. __ ({last_name} Depo. Trns.) at p. {range_str}.)\n\n"
            cursor.insertText(citation, cite_fmt)

            group_end_pos = cursor.position()
            group_ex_ids = [ex.id for ex in group_exchanges]
            self._result_page_anchors.append((group_start_pos, group_end_pos, first_page))
            self._result_exchange_anchors.append((
                group_start_pos, group_end_pos, group_ex_ids,
                self._active_idx, extraction_idx
            ))

        self.results_viewer.setTextCursor(cursor)
        self.results_viewer.ensureCursorVisible()

    def _load_highlighted_pdf(self, highlight_path: str):
        """Reload the viewer with the highlighted PDF copy."""
        if not os.path.isfile(highlight_path):
            return
        current_page = self.transcript_viewer.get_current_page()
        self.transcript_viewer.load_pdf(highlight_path)
        if current_page > 0:
            page = current_page
            QTimer.singleShot(2000, lambda: self.transcript_viewer.go_to_page(page))

    def _get_ai_output_dir(self) -> str:
        """Get the NOTES/AI OUTPUT directory for the current case.

        Never returns the transcript's own directory — always resolves to
        a NOTES/AI OUTPUT folder so we don't pollute the transcript folder.
        """
        # Try via file_number first
        if self.file_number:
            from icharlotte_core.utils import get_case_path
            case_path = get_case_path(self.file_number)
            if case_path:
                ai_dir = os.path.join(case_path, "NOTES", "AI OUTPUT")
                os.makedirs(ai_dir, exist_ok=True)
                return ai_dir

        # Fallback: derive case root from transcript path by walking up
        # to find a folder that contains a NOTES subfolder
        if self._transcripts and self._active_idx >= 0:
            transcript_path = self._transcripts[self._active_idx]["path"]
            current = os.path.dirname(transcript_path)
            while current and current != os.path.dirname(current):
                notes_dir = os.path.join(current, "NOTES")
                if os.path.isdir(notes_dir):
                    ai_dir = os.path.join(notes_dir, "AI OUTPUT")
                    os.makedirs(ai_dir, exist_ok=True)
                    return ai_dir
                current = os.path.dirname(current)

            # Last resort: create NOTES/AI OUTPUT next to transcript
            ai_dir = os.path.join(
                os.path.dirname(transcript_path), "NOTES", "AI OUTPUT"
            )
            os.makedirs(ai_dir, exist_ok=True)
            return ai_dir

        return os.getcwd()

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

        # Determine default path — save in NOTES/AI OUTPUT within case folder
        first_index = all_extractions[0][0]
        default_dir = self._get_ai_output_dir()
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
