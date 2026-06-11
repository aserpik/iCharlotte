"""Workbench tab for managing the firm-brief Sample Library roots."""

from __future__ import annotations

import json
import os
from typing import List, Optional

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from . import theme

# Default path for the roots config JSON (relative to cwd at import time).
_DEFAULT_ROOTS_CONFIG = os.path.join(
    os.path.dirname(__file__),
    "..",
    "..",
    "Scripts",
    "prompts",
    "firm_briefs",
    "roots.json",
)


class _ReindexWorker(QThread):
    """Background worker that calls ingest_root for each root folder."""

    progress = Signal(str)
    finished = Signal(bool, str)  # (success, message)

    def __init__(self, roots: List[str], ocr_only: bool = False, parent=None):
        super().__init__(parent)
        self._roots = list(roots)
        self._ocr_only = ocr_only

    def run(self):
        try:
            from icharlotte_core.firm_briefs.embedding import get_embedder
            from icharlotte_core.firm_briefs.index import FirmBriefIndex
            from icharlotte_core.firm_briefs.ingest import ingest_root
            from icharlotte_core import config as _cfg

            index = FirmBriefIndex(_cfg.FIRM_BRIEFS_DATA_DIR)
            embedder = get_embedder()

            total = len(self._roots)
            for i, root in enumerate(self._roots):
                self.progress.emit(f"Indexing root {i + 1}/{total}: {root}")
                ingest_root(
                    root,
                    index,
                    embedder,
                    on_progress=lambda msg: self.progress.emit(msg),
                )
            self.finished.emit(True, f"Re-indexed {total} root(s) successfully.")
        except Exception as exc:
            self.finished.emit(False, f"Re-index failed: {exc}")


class SampleLibraryTab(QWidget):
    """Workbench tab to manage firm-brief library roots and trigger re-indexing."""

    def __init__(
        self,
        *,
        roots_config_path: str = _DEFAULT_ROOTS_CONFIG,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._roots_config_path = roots_config_path
        self._roots: List[str] = self._load_roots()
        self._worker: Optional[_ReindexWorker] = None

        self._build_ui()

    # ------------------------------------------------------------------
    # Programmatic API (event-loop-free — safe to call from tests)
    # ------------------------------------------------------------------

    def roots(self) -> List[str]:
        """Return the current list of roots in insertion order."""
        return list(self._roots)

    def add_root_programmatic(self, path: str) -> None:
        """Add *path* to the roots list and persist immediately."""
        if path and path not in self._roots:
            self._roots.append(path)
            self._save_roots()
            self._refresh_list()

    def remove_root_programmatic(self, path: str) -> bool:
        """Remove *path* from the roots list and persist. Returns True if found."""
        if path in self._roots:
            self._roots.remove(path)
            self._save_roots()
            self._refresh_list()
            return True
        return False

    def stats_text(self, index) -> str:
        """Return a human-readable stats string.

        ``index`` is either None (not yet built) or an object with a
        ``stats() -> dict`` method returning at least ``briefs`` and
        ``citations`` keys.
        """
        if index is None:
            return "Index not built."
        try:
            s = index.stats()
            briefs = s.get("briefs", "?")
            citations = s.get("citations", "?")
            lines = [f"{briefs} briefs, {citations} citations"]
            # Append any per-type breakdown if present
            by_type = s.get("by_type")
            if by_type:
                lines.append(
                    "  " + "; ".join(f"{k}: {v}" for k, v in sorted(by_type.items()))
                )
            return "\n".join(lines)
        except Exception as exc:
            return f"Stats unavailable: {exc}"

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _load_roots(self) -> List[str]:
        if os.path.isfile(self._roots_config_path):
            try:
                with open(self._roots_config_path, encoding="utf-8") as fh:
                    data = json.load(fh)
                return [str(r) for r in (data.get("roots") or [])]
            except Exception:
                pass
        # Seed from config when the file is absent.
        try:
            from icharlotte_core import config as _cfg
            return list(_cfg.FIRM_BRIEFS_ROOTS)
        except Exception:
            return []

    def _save_roots(self) -> None:
        dir_path = os.path.dirname(self._roots_config_path)
        if dir_path:
            os.makedirs(dir_path, exist_ok=True)
        try:
            with open(self._roots_config_path, "w", encoding="utf-8") as fh:
                json.dump({"roots": self._roots}, fh, indent=2)
        except Exception:
            pass

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        # --- Header label ---
        header = QLabel("Firm Brief Sample Library — Root Folders")
        header.setStyleSheet("font-weight: bold; font-size: 13px;")
        layout.addWidget(header)

        desc = QLabel(
            "Add the top-level folders that contain your firm's brief PDFs.\n"
            "Subfolders are scanned recursively when you click Re-index."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #555;")
        layout.addWidget(desc)

        # --- Roots list ---
        self._list_widget = QListWidget()
        self._refresh_list()
        layout.addWidget(self._list_widget, 1)

        # --- Add / Remove buttons ---
        btn_row = QHBoxLayout()
        self._add_btn = QPushButton("Add Folder…")
        self._add_btn.clicked.connect(self._on_add_clicked)
        btn_row.addWidget(self._add_btn)

        self._remove_btn = QPushButton("Remove Selected")
        self._remove_btn.clicked.connect(self._on_remove_clicked)
        btn_row.addWidget(self._remove_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        # --- Re-index section ---
        reindex_row = QHBoxLayout()

        self._ocr_check = QCheckBox("OCR image-only PDFs")
        self._ocr_check.setToolTip(
            "Run OCR on image-only PDF pages during ingestion (slower; stub — not yet wired)."
        )
        reindex_row.addWidget(self._ocr_check)

        reindex_row.addStretch()

        self._reindex_btn = theme.primary_button("Re-index")
        self._reindex_btn.clicked.connect(self._on_reindex_clicked)
        reindex_row.addWidget(self._reindex_btn)
        layout.addLayout(reindex_row)

        # --- Progress / status log ---
        self._status_label = QLabel("")
        self._status_label.setWordWrap(True)
        self._status_label.setStyleSheet("color: #333; font-style: italic;")
        layout.addWidget(self._status_label)

    def _refresh_list(self) -> None:
        self._list_widget.clear()
        for root in self._roots:
            self._list_widget.addItem(QListWidgetItem(root))

    # ------------------------------------------------------------------
    # Interactive handlers
    # ------------------------------------------------------------------

    def _on_add_clicked(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select Brief Library Root", "")
        if folder:
            self.add_root_programmatic(folder)

    def _on_remove_clicked(self) -> None:
        row = self._list_widget.currentRow()
        if row < 0 or row >= len(self._roots):
            return
        path = self._roots[row]
        confirm = QMessageBox.question(
            self,
            "Remove root",
            f"Remove '{path}' from the library roots?\n(No files are deleted.)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            self.remove_root_programmatic(path)

    def _on_reindex_clicked(self) -> None:
        if self._worker and self._worker.isRunning():
            QMessageBox.information(
                self, "Re-index", "Re-indexing is already in progress."
            )
            return
        if not self._roots:
            QMessageBox.warning(self, "Re-index", "No library roots configured.")
            return

        self._reindex_btn.setEnabled(False)
        self._status_label.setText("Starting re-index…")

        self._worker = _ReindexWorker(
            self._roots,
            ocr_only=self._ocr_check.isChecked(),
            parent=self,
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_reindex_done)
        self._worker.start()

    def _on_progress(self, msg: str) -> None:
        self._status_label.setText(msg)

    def _on_reindex_done(self, success: bool, message: str) -> None:
        self._reindex_btn.setEnabled(True)
        self._status_label.setText(message)
