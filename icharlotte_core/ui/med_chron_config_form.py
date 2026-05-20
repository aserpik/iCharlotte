"""Configuration form for the multi-analysis Med-Cron task.

Reads the Phase 1 session JSON, presents:
- a checkbox per curated analysis (Rewrite Chronology pre-checked)
- a "Custom analyses" panel with add/remove rows
- a "narrative missing" warning banner when applicable

On commit_user_config(), validates selection and writes user_config back
to the session, flipping phase to ``ready_to_run``.

Custom analyses persist globally: rows that the user fills in are saved
to ``custom_analyses_store`` on every commit and pre-populated when the
form is next opened. Each row has an "Include in this run" checkbox so
the user can skip a saved analysis once without deleting it from the
global store.
"""

from pathlib import Path
from typing import Optional

from PySide6.QtCore import QSettings, Qt
from PySide6.QtWidgets import (
    QCheckBox, QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit,
    QPushButton, QScrollArea, QSplitter, QVBoxLayout, QWidget,
)

from icharlotte_core.med_chron import custom_analyses_store, session_manager


# ---- Context-document helpers ----

SUPPORTED_CONTEXT_EXTS = (".pdf", ".docx", ".txt")
MAX_CONTEXT_CHARS = 120_000  # mirrors the Phase 2 cap; only used by callers that want it


def sniff_text_layer(path: str) -> tuple[bool, str]:
    """Return (has_text, reason).

    Cheap, attach-time check: does this file appear to have extractable
    text without needing OCR? Used by the UI to warn the user.

    - .txt: any non-whitespace in the first 4 KB.
    - .docx: any non-empty paragraph in the first ~50 paragraphs.
    - .pdf: at least 200 chars of text extractable from pages 0..2.
    - any error / unknown extension: (False, "...").
    """
    import os

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".txt":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                sample = f.read(4096)
            return (bool(sample.strip()), "")
        if ext == ".docx":
            from docx import Document
            doc = Document(path)
            text = "".join(p.text for p in doc.paragraphs[:50])
            return (len(text.strip()) > 0, "")
        if ext == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(path)
            n = min(3, len(reader.pages))
            text = ""
            for i in range(n):
                try:
                    text += reader.pages[i].extract_text() or ""
                except Exception:
                    pass
            return (len(text.strip()) > 200, "")
        return (False, f"unsupported extension {ext}")
    except FileNotFoundError:
        return (False, "file not found")
    except Exception as e:
        return (False, f"could not read file: {e}")


_ANALYSES_SPLITTER_KEY = "med_chron_form/analyses_splitter_sizes"


class CustomAnalysisRow(QWidget):
    """One row in the custom-analyses list: include-checkbox + label + instruction + remove btn."""

    def __init__(self, parent: QWidget, on_remove):
        super().__init__(parent)
        self._on_remove = on_remove

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        top = QHBoxLayout()
        self.include_cb = QCheckBox()
        self.include_cb.setChecked(True)
        self.include_cb.setToolTip("Include this analysis in the current run")
        top.addWidget(self.include_cb)
        top.addWidget(QLabel("Label:"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("Short name (e.g. 'Left-knee mentions')")
        top.addWidget(self.label_edit, 1)
        self.remove_btn = QPushButton("−")
        self.remove_btn.setFixedSize(24, 24)
        self.remove_btn.setStyleSheet("QPushButton { color: #c62828; font-weight: bold; }")
        self.remove_btn.setToolTip("Remove this analysis (also deletes from global save)")
        self.remove_btn.clicked.connect(self._handle_remove)
        top.addWidget(self.remove_btn)
        layout.addLayout(top)

        layout.addWidget(QLabel("Request:"))
        self.instruction_edit = QPlainTextEdit()
        self.instruction_edit.setPlaceholderText("Describe the analysis…")
        self.instruction_edit.setFixedHeight(60)
        layout.addWidget(self.instruction_edit)

        self.setStyleSheet(
            "CustomAnalysisRow { border: 1px solid #ddd; border-radius: 4px; }"
        )

    def _handle_remove(self):
        self._on_remove(self)

    def label(self) -> str:
        return self.label_edit.text().strip()

    def instruction(self) -> str:
        return self.instruction_edit.toPlainText().strip()

    def is_included(self) -> bool:
        return self.include_cb.isChecked()

    def is_empty(self) -> bool:
        return not self.label() and not self.instruction()


class MedChronConfigForm(QWidget):
    """Pickable list of catalog analyses + custom analysis rows."""

    def __init__(self, session_path, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.session_path = Path(session_path)
        self._session = session_manager.read_session(self.session_path)
        self._settings = QSettings("iCharlotte", "iCharlotte")

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        header = QLabel(
            f"<b>Analyses to run on {self._session.get('provider_name', '')}</b>"
        )
        root.addWidget(header)

        # Narrative-missing warning banner
        self.narrative_missing_banner = QLabel(
            "⚠ Narrative text not found in this document — "
            "Rewrite Chronology will be skipped."
        )
        self.narrative_missing_banner.setStyleSheet(
            "background-color: #FFF3CD; color: #856404; "
            "padding: 6px; border-radius: 4px;"
        )
        self.narrative_missing_banner.setVisible(
            bool(self._session.get("narrative_missing"))
        )
        root.addWidget(self.narrative_missing_banner)

        # Vertical splitter: Curated (top) ↔ Custom (bottom).
        self._analyses_splitter = QSplitter(Qt.Orientation.Vertical)
        self._analyses_splitter.setChildrenCollapsible(False)
        self._analyses_splitter.setHandleWidth(6)
        self._analyses_splitter.setStyleSheet(
            "QSplitter::handle { background: #e0e0e0; }"
            " QSplitter::handle:hover { background: #b0b0b0; }"
            " QSplitter::handle:pressed { background: #909090; }"
        )

        # --- Curated catalog pane ---
        curated_section = QWidget()
        curated_layout = QVBoxLayout(curated_section)
        curated_layout.setContentsMargins(0, 0, 0, 0)
        curated_layout.setSpacing(4)
        curated_layout.addWidget(QLabel("<b>Curated analyses:</b>"))

        curated_inner = QWidget()
        curated_inner_layout = QVBoxLayout(curated_inner)
        curated_inner_layout.setContentsMargins(0, 0, 0, 0)
        curated_inner_layout.setSpacing(0)

        self.catalog_checkboxes: dict[str, QCheckBox] = {}
        for entry in self._session.get("catalog", []):
            row = QWidget()
            row_layout = QVBoxLayout(row)
            row_layout.setContentsMargins(16, 2, 0, 2)
            row_layout.setSpacing(0)
            cb = QCheckBox(entry.get("title", entry.get("id", "")))
            cb.setChecked(bool(entry.get("default_selected")))
            self.catalog_checkboxes[entry["id"]] = cb
            row_layout.addWidget(cb)
            desc = QLabel(entry.get("description", ""))
            desc.setStyleSheet("color: #666; font-size: 11px; padding-left: 22px;")
            desc.setWordWrap(True)
            row_layout.addWidget(desc)
            curated_inner_layout.addWidget(row)
        curated_inner_layout.addStretch()

        curated_scroll = QScrollArea()
        curated_scroll.setWidget(curated_inner)
        curated_scroll.setWidgetResizable(True)
        curated_scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        curated_layout.addWidget(curated_scroll, 1)
        self._analyses_splitter.addWidget(curated_section)

        # --- Custom analyses pane ---
        custom_section = QWidget()
        custom_layout = QVBoxLayout(custom_section)
        custom_layout.setContentsMargins(0, 0, 0, 0)
        custom_layout.setSpacing(4)

        custom_header = QHBoxLayout()
        custom_header.addWidget(QLabel("<b>Custom analyses:</b>"))
        custom_header.addStretch()
        custom_layout.addLayout(custom_header)

        # Scrollable container for custom rows.
        self._custom_container = QWidget()
        self._custom_container_layout = QVBoxLayout(self._custom_container)
        self._custom_container_layout.setContentsMargins(0, 0, 0, 0)
        self._custom_container_layout.setSpacing(4)
        self._custom_container_layout.addStretch()  # pin rows to top

        custom_scroll = QScrollArea()
        custom_scroll.setWidget(self._custom_container)
        custom_scroll.setWidgetResizable(True)
        custom_scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        custom_layout.addWidget(custom_scroll, 1)

        self.custom_rows: list[CustomAnalysisRow] = []

        # Pre-populate previously-saved custom analyses so the user does
        # not have to re-type recurring requests.
        for saved in custom_analyses_store.load():
            row = self.add_custom_row()
            row.label_edit.setText(saved.get("label", ""))
            row.instruction_edit.setPlainText(saved.get("instruction", ""))

        add_btn = QPushButton("+ Add custom analysis")
        add_btn.clicked.connect(self.add_custom_row)
        custom_layout.addWidget(add_btn)

        self._analyses_splitter.addWidget(custom_section)
        self._analyses_splitter.setStretchFactor(0, 1)
        self._analyses_splitter.setStretchFactor(1, 1)
        self._restore_analyses_splitter_sizes()
        self._analyses_splitter.splitterMoved.connect(
            self._save_analyses_splitter_sizes
        )
        root.addWidget(self._analyses_splitter, 1)

        # Inline validation error label.
        self._error_label = QLabel("")
        self._error_label.setStyleSheet("color: #c62828; font-style: italic;")
        self._error_label.setVisible(False)
        root.addWidget(self._error_label)

    def _restore_analyses_splitter_sizes(self) -> None:
        saved = self._settings.value(_ANALYSES_SPLITTER_KEY)
        if saved:
            try:
                sizes = [int(x) for x in saved]
                if len(sizes) == 2 and all(s > 0 for s in sizes):
                    self._analyses_splitter.setSizes(sizes)
                    return
            except (TypeError, ValueError):
                pass
        self._analyses_splitter.setSizes([240, 200])

    def _save_analyses_splitter_sizes(self, *_args) -> None:
        self._settings.setValue(
            _ANALYSES_SPLITTER_KEY, self._analyses_splitter.sizes()
        )

    def add_custom_row(self) -> "CustomAnalysisRow":
        row = CustomAnalysisRow(self._custom_container, self._remove_custom_row)
        # Insert above the stretch at the end.
        idx = self._custom_container_layout.count() - 1
        self._custom_container_layout.insertWidget(idx, row)
        self.custom_rows.append(row)
        return row

    def _remove_custom_row(self, row: "CustomAnalysisRow") -> None:
        if row in self.custom_rows:
            self.custom_rows.remove(row)
        self._custom_container_layout.removeWidget(row)
        row.setParent(None)
        row.deleteLater()

    def _selected_catalog_ids(self) -> list[str]:
        return [cid for cid, cb in self.catalog_checkboxes.items() if cb.isChecked()]

    def _validated_custom_rows(self) -> tuple[list[dict], list[dict], str]:
        """Return ``(all_valid_rows, included_rows, error_msg)``.

        - ``all_valid_rows`` — every non-empty, fully-filled row. This is
          what gets persisted to the global store so the user keeps them
          for next time even if the include-checkbox is unchecked.
        - ``included_rows`` — subset of ``all_valid_rows`` whose include
          checkbox is checked. This is what actually runs in Phase 2.
        - ``error_msg`` — non-empty if a row is partially filled.
        """
        all_valid: list[dict] = []
        included: list[dict] = []
        for r in self.custom_rows:
            if r.is_empty():
                continue
            lbl, instr = r.label(), r.instruction()
            if not lbl or not instr:
                return [], [], (
                    "Custom analyses need both a label and an instruction. "
                    "Fill in (or remove) the partially-completed row."
                )
            entry = {"label": lbl, "instruction": instr}
            all_valid.append(entry)
            if r.is_included():
                included.append(entry)
        return all_valid, included, ""

    def commit_user_config(self) -> bool:
        """Validate, persist saved analyses globally, and write user_config.

        Returns True on success.
        """
        self._error_label.setVisible(False)

        selected = self._selected_catalog_ids()
        all_valid_custom, included_custom, err = self._validated_custom_rows()
        if err:
            self._error_label.setText(err)
            self._error_label.setVisible(True)
            return False
        if not selected and not included_custom:
            self._error_label.setText(
                "Select at least one analysis, or add a custom analysis."
            )
            self._error_label.setVisible(True)
            return False

        # Auto-save the full list of valid custom rows (including unchecked
        # ones) to the global store so they reappear next time.
        try:
            custom_analyses_store.save(all_valid_custom)
        except OSError:
            # Persisting globally is best-effort; the run can still proceed.
            pass

        session_manager.update_user_config(
            self.session_path,
            {
                "selected_catalog_ids": selected,
                "custom_analyses": included_custom,
            },
        )
        return True
