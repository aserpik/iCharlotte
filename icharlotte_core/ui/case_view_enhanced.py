"""
Enhanced Case View Tab Components

This module contains improved UI components for the Case View Tab including:
- Enhanced Agent Buttons with running indicator, last docket date, and settings
- Enhanced File Tree with processing status, tags, and page count columns
- File Preview Pane
- Advanced Filtering
- Output Browser with compare, search, and export functionality
- Processing Log
"""

import os
import re
import json
import shutil
import datetime
from functools import partial

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QSplitter, QTreeWidget, QTreeWidgetItem, QHeaderView, QMenu,
    QLineEdit, QComboBox, QCheckBox, QTextEdit, QScrollArea,
    QAbstractItemView, QMessageBox, QDialog, QFormLayout,
    QDialogButtonBox, QFileDialog, QListWidget, QListWidgetItem,
    QTabWidget, QProgressBar, QApplication, QToolButton, QInputDialog,
    QSizePolicy, QGroupBox, QSpinBox, QDateEdit, QStackedWidget
)
from PySide6.QtCore import Qt, Signal, QTimer, QDate, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QAction, QColor, QFont, QMovie, QIcon, QPainter, QPen

from ..utils import log_event
from ..config import GEMINI_DATA_DIR, SCRIPTS_DIR


# =============================================================================
# Processing Log Database
# =============================================================================

class ProcessingLogDB:
    """Manages persistent processing logs for each case."""

    def __init__(self, file_number):
        self.file_number = file_number
        self.log_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}_processing_log.json")
        self.logs = self._load()

    def _load(self):
        if os.path.exists(self.log_path):
            try:
                with open(self.log_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return []
        return []

    def _save(self):
        try:
            os.makedirs(os.path.dirname(self.log_path), exist_ok=True)
            with open(self.log_path, 'w', encoding='utf-8') as f:
                json.dump(self.logs, f, indent=2)
        except Exception as e:
            log_event(f"Error saving processing log: {e}", "error")

    def add_entry(self, file_path, task_type, status, output_path=None, error_message=None, duration_sec=0):
        entry = {
            "timestamp": datetime.datetime.now().isoformat(),
            "file_path": file_path,
            "file_name": os.path.basename(file_path),
            "task_type": task_type,
            "status": status,  # "success", "failed", "in_progress"
            "output_path": output_path,
            "error_message": error_message,
            "duration_sec": duration_sec
        }
        self.logs.insert(0, entry)  # Newest first
        self._save()
        return entry

    def get_file_processing_status(self, file_path):
        """Get all processing done on a specific file."""
        return [log for log in self.logs if log.get("file_path") == file_path]

    def get_all_logs(self):
        return self.logs

    def clear_logs(self):
        self.logs = []
        self._save()


# =============================================================================
# File Tags Database
# =============================================================================

class FileTagsDB:
    """Manages file tags for each case."""

    def __init__(self, file_number):
        self.file_number = file_number
        self.tags_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}_file_tags.json")
        self.tags = self._load()

    def _load(self):
        if os.path.exists(self.tags_path):
            try:
                with open(self.tags_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def _save(self):
        try:
            os.makedirs(os.path.dirname(self.tags_path), exist_ok=True)
            with open(self.tags_path, 'w', encoding='utf-8') as f:
                json.dump(self.tags, f, indent=2)
        except Exception as e:
            log_event(f"Error saving file tags: {e}", "error")

    def get_tags(self, file_path):
        return self.tags.get(file_path, [])

    def set_tags(self, file_path, tags):
        self.tags[file_path] = tags
        self._save()

    def add_tag(self, file_path, tag):
        if file_path not in self.tags:
            self.tags[file_path] = []
        if tag not in self.tags[file_path]:
            self.tags[file_path].append(tag)
            self._save()

    def remove_tag(self, file_path, tag):
        if file_path in self.tags and tag in self.tags[file_path]:
            self.tags[file_path].remove(tag)
            self._save()

    def get_all_tags(self):
        """Get unique list of all tags used."""
        all_tags = set()
        for tags in self.tags.values():
            all_tags.update(tags)
        return sorted(list(all_tags))


# =============================================================================
# Agent Settings Database
# =============================================================================

class AgentSettingsDB:
    """Manages per-agent configuration settings."""

    DEFAULT_SETTINGS = {
        "docket.py": {
            "auto_run_days": 30,
            "save_pdf": True,
            "headless": True
        },
        "complaint.py": {
            "extract_parties": True,
            "headless": True
        },
        "report.py": {
            "include_discovery": True,
            "include_depositions": True,
            "max_length": 5000
        },
        "summarize.py": {
            "max_pages": 0,  # 0 = no limit
            "output_format": "markdown"
        },
        "ocr.py": {
            "language": "eng",
            "deskew": True
        },
        "extract_timeline.py": {
            "include_medical": True,
            "include_legal": True,
            "date_format": "MM/DD/YYYY"
        },
        "detect_contradictions.py": {
            "sensitivity": "medium",
            "include_depositions": True,
            "include_discovery": True
        }
    }

    def __init__(self):
        self.settings_path = os.path.join(GEMINI_DATA_DIR, "..", "config", "agent_settings.json")
        self.settings = self._load()

    def _load(self):
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    # Merge with defaults
                    merged = dict(self.DEFAULT_SETTINGS)
                    for key, val in loaded.items():
                        if key in merged:
                            merged[key].update(val)
                        else:
                            merged[key] = val
                    return merged
            except:
                return dict(self.DEFAULT_SETTINGS)
        return dict(self.DEFAULT_SETTINGS)

    def _save(self):
        try:
            os.makedirs(os.path.dirname(self.settings_path), exist_ok=True)
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            log_event(f"Error saving agent settings: {e}", "error")

    def get_settings(self, script_name):
        return self.settings.get(script_name, {})

    def update_settings(self, script_name, settings):
        if script_name not in self.settings:
            self.settings[script_name] = {}
        self.settings[script_name].update(settings)
        self._save()


# =============================================================================
# Spinning Indicator Widget
# =============================================================================

class SpinningIndicator(QLabel):
    """A simple spinning indicator for showing agent is running."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._rotate)
        self.setFixedSize(16, 16)
        self.setStyleSheet("background: transparent;")
        self._running = False

    def _rotate(self):
        self.angle = (self.angle + 30) % 360
        self.update()

    def paintEvent(self, event):
        if not self._running:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        pen = QPen(QColor("#2196F3"))
        pen.setWidth(2)
        painter.setPen(pen)

        center = self.rect().center()
        painter.translate(center)
        painter.rotate(self.angle)

        # Draw spinning arc
        from PySide6.QtCore import QRectF
        rect = QRectF(-6, -6, 12, 12)
        painter.drawArc(rect, 0 * 16, 270 * 16)

        painter.end()

    def start(self):
        self._running = True
        self.timer.start(50)
        self.show()

    def stop(self):
        self._running = False
        self.timer.stop()
        self.hide()


# =============================================================================
# Enhanced Agent Button Widget
# =============================================================================

class EnhancedAgentButton(QFrame):
    """Enhanced agent button with running indicator, status info, and settings."""

    clicked = Signal()
    settings_clicked = Signal()

    def __init__(self, name, script_name, parent=None):
        super().__init__(parent)
        self.name = name
        self.script_name = script_name
        self.is_running = False

        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            EnhancedAgentButton {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            EnhancedAgentButton:hover {
                background-color: #e8e8e8;
                border-color: #bbb;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 4, 4)
        layout.setSpacing(4)

        # Spinning indicator (hidden by default)
        self.spinner = SpinningIndicator()
        self.spinner.hide()
        layout.addWidget(self.spinner)

        # Main content area
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        # Agent name button
        self.name_btn = QPushButton(name)
        self.name_btn.setStyleSheet("""
            QPushButton {
                border: none;
                background: transparent;
                text-align: left;
                font-size: 11px;
                font-weight: bold;
                padding: 2px;
            }
            QPushButton:hover {
                color: #1976D2;
            }
        """)
        self.name_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.name_btn.clicked.connect(self.clicked.emit)
        content_layout.addWidget(self.name_btn)

        # Status label (e.g., "Last run: 2 days ago")
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: gray; font-size: 9px; padding-left: 2px;")
        content_layout.addWidget(self.status_label)

        layout.addLayout(content_layout)
        layout.addStretch()

        # Settings button
        self.settings_btn = QToolButton()
        self.settings_btn.setText("⚙")
        self.settings_btn.setStyleSheet("""
            QToolButton {
                border: none;
                background: transparent;
                font-size: 12px;
                color: #888;
            }
            QToolButton:hover {
                color: #333;
                background: #ddd;
                border-radius: 2px;
            }
        """)
        self.settings_btn.setFixedSize(20, 20)
        self.settings_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.settings_btn.clicked.connect(self.settings_clicked.emit)
        layout.addWidget(self.settings_btn)

        self.setFixedHeight(45)

    def set_running(self, running):
        self.is_running = running
        if running:
            self.spinner.start()
            self.name_btn.setEnabled(False)
            self.setStyleSheet("""
                EnhancedAgentButton {
                    background-color: #e3f2fd;
                    border: 1px solid #2196F3;
                    border-radius: 4px;
                }
            """)
        else:
            self.spinner.stop()
            self.name_btn.setEnabled(True)
            self.setStyleSheet("""
                EnhancedAgentButton {
                    background-color: #f5f5f5;
                    border: 1px solid #ddd;
                    border-radius: 4px;
                }
                EnhancedAgentButton:hover {
                    background-color: #e8e8e8;
                    border-color: #bbb;
                }
            """)

    def set_status(self, status_text):
        self.status_label.setText(status_text)

    def set_last_run(self, last_run_date):
        if not last_run_date:
            self.status_label.setText("")
            return

        try:
            if isinstance(last_run_date, str):
                last_date = datetime.datetime.strptime(last_run_date, "%Y-%m-%d")
            else:
                last_date = last_run_date

            days_ago = (datetime.datetime.now() - last_date).days

            if days_ago == 0:
                self.status_label.setText("Last: Today")
            elif days_ago == 1:
                self.status_label.setText("Last: Yesterday")
            else:
                self.status_label.setText(f"Last: {days_ago} days ago")
        except:
            self.status_label.setText(f"Last: {last_run_date}")


# =============================================================================
# Agent Settings Dialog
# =============================================================================

class AgentSettingsDialog(QDialog):
    """Dialog for configuring agent-specific settings."""

    def __init__(self, script_name, settings_db, parent=None):
        super().__init__(parent)
        self.script_name = script_name
        self.settings_db = settings_db
        self.setWindowTitle(f"Settings: {script_name}")
        self.setMinimumWidth(400)

        layout = QVBoxLayout(self)

        # Settings form
        self.form_layout = QFormLayout()
        self.widgets = {}

        settings = self.settings_db.get_settings(script_name)

        for key, value in settings.items():
            widget = self._create_widget(key, value)
            if widget:
                self.form_layout.addRow(self._format_label(key), widget)
                self.widgets[key] = widget

        layout.addLayout(self.form_layout)

        # Buttons
        btn_layout = QHBoxLayout()

        reset_btn = QPushButton("Reset to Defaults")
        reset_btn.clicked.connect(self._reset_defaults)
        btn_layout.addWidget(reset_btn)

        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        save_btn.clicked.connect(self._save)
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)

    def _format_label(self, key):
        """Convert key_name to Key Name."""
        return key.replace("_", " ").title() + ":"

    def _create_widget(self, key, value):
        """Create appropriate widget based on value type."""
        if isinstance(value, bool):
            widget = QCheckBox()
            widget.setChecked(value)
            return widget
        elif isinstance(value, int):
            widget = QSpinBox()
            widget.setRange(0, 99999)
            widget.setValue(value)
            return widget
        elif isinstance(value, str):
            if key in ["sensitivity"]:
                widget = QComboBox()
                widget.addItems(["low", "medium", "high"])
                widget.setCurrentText(value)
                return widget
            elif key in ["output_format"]:
                widget = QComboBox()
                widget.addItems(["markdown", "plain", "html"])
                widget.setCurrentText(value)
                return widget
            else:
                widget = QLineEdit()
                widget.setText(value)
                return widget
        return None

    def _get_widget_value(self, widget):
        """Get value from widget."""
        if isinstance(widget, QCheckBox):
            return widget.isChecked()
        elif isinstance(widget, QSpinBox):
            return widget.value()
        elif isinstance(widget, QComboBox):
            return widget.currentText()
        elif isinstance(widget, QLineEdit):
            return widget.text()
        return None

    def _reset_defaults(self):
        defaults = AgentSettingsDB.DEFAULT_SETTINGS.get(self.script_name, {})
        for key, widget in self.widgets.items():
            if key in defaults:
                value = defaults[key]
                if isinstance(widget, QCheckBox):
                    widget.setChecked(value)
                elif isinstance(widget, QSpinBox):
                    widget.setValue(value)
                elif isinstance(widget, QComboBox):
                    widget.setCurrentText(value)
                elif isinstance(widget, QLineEdit):
                    widget.setText(value)

    def _save(self):
        settings = {}
        for key, widget in self.widgets.items():
            settings[key] = self._get_widget_value(widget)

        self.settings_db.update_settings(self.script_name, settings)
        self.accept()


# =============================================================================
# Advanced Filter Widget
# =============================================================================

class AdvancedFilterWidget(QFrame):
    """Advanced filtering options for the file tree."""

    filter_changed = Signal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet("background-color: #fafafa; border: 1px solid #ddd;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Header with collapse button
        header_layout = QHBoxLayout()
        header_label = QLabel("<b>Advanced Filters</b>")
        header_layout.addWidget(header_label)
        header_layout.addStretch()

        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setFixedHeight(22)
        self.clear_btn.clicked.connect(self._clear_filters)
        header_layout.addWidget(self.clear_btn)

        layout.addLayout(header_layout)

        # Filter options grid
        options_layout = QHBoxLayout()

        # File Type Filter
        type_group = QGroupBox("File Type")
        type_layout = QVBoxLayout(type_group)
        self.type_pdf = QCheckBox("PDF")
        self.type_doc = QCheckBox("Word (doc/docx)")
        self.type_img = QCheckBox("Images")
        self.type_other = QCheckBox("Other")

        for cb in [self.type_pdf, self.type_doc, self.type_img, self.type_other]:
            cb.setChecked(True)
            cb.stateChanged.connect(self._emit_filter)
            type_layout.addWidget(cb)

        options_layout.addWidget(type_group)

        # Processing Status Filter
        status_group = QGroupBox("Processing Status")
        status_layout = QVBoxLayout(status_group)
        self.status_unprocessed = QCheckBox("Unprocessed")
        self.status_ocr = QCheckBox("OCR'd")
        self.status_summarized = QCheckBox("Summarized")
        self.status_all = QCheckBox("All")
        self.status_all.setChecked(True)

        self.status_all.stateChanged.connect(self._on_all_status_changed)
        for cb in [self.status_unprocessed, self.status_ocr, self.status_summarized]:
            cb.stateChanged.connect(self._emit_filter)
            status_layout.addWidget(cb)
        status_layout.addWidget(self.status_all)

        options_layout.addWidget(status_group)

        # Date Filter
        date_group = QGroupBox("Date Modified")
        date_layout = QVBoxLayout(date_group)

        self.date_filter = QComboBox()
        self.date_filter.addItems([
            "Any time",
            "Today",
            "Last 7 days",
            "Last 30 days",
            "Last 90 days",
            "Custom range..."
        ])
        self.date_filter.currentIndexChanged.connect(self._on_date_filter_changed)
        date_layout.addWidget(self.date_filter)

        # Custom date range (hidden by default)
        self.custom_date_widget = QWidget()
        custom_layout = QHBoxLayout(self.custom_date_widget)
        custom_layout.setContentsMargins(0, 0, 0, 0)
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        custom_layout.addWidget(QLabel("From:"))
        custom_layout.addWidget(self.date_from)
        custom_layout.addWidget(QLabel("To:"))
        custom_layout.addWidget(self.date_to)
        self.custom_date_widget.hide()
        date_layout.addWidget(self.custom_date_widget)

        options_layout.addWidget(date_group)

        # Tags Filter
        tags_group = QGroupBox("Tags")
        tags_layout = QVBoxLayout(tags_group)
        self.tags_combo = QComboBox()
        self.tags_combo.addItem("All Tags")
        self.tags_combo.currentIndexChanged.connect(self._emit_filter)
        tags_layout.addWidget(self.tags_combo)

        options_layout.addWidget(tags_group)

        layout.addLayout(options_layout)

        # Quick filter buttons
        quick_layout = QHBoxLayout()
        quick_layout.addWidget(QLabel("Quick:"))

        self.quick_unprocessed = QPushButton("Unprocessed Only")
        self.quick_unprocessed.setCheckable(True)
        self.quick_unprocessed.clicked.connect(self._quick_unprocessed)
        quick_layout.addWidget(self.quick_unprocessed)

        self.quick_recent = QPushButton("Recent (7 days)")
        self.quick_recent.setCheckable(True)
        self.quick_recent.clicked.connect(self._quick_recent)
        quick_layout.addWidget(self.quick_recent)

        quick_layout.addStretch()
        layout.addLayout(quick_layout)

    def _emit_filter(self):
        filters = self.get_filters()
        self.filter_changed.emit(filters)

    def _on_all_status_changed(self, state):
        if state == Qt.CheckState.Checked.value:
            self.status_unprocessed.setChecked(False)
            self.status_ocr.setChecked(False)
            self.status_summarized.setChecked(False)
        self._emit_filter()

    def _on_date_filter_changed(self, index):
        if self.date_filter.currentText() == "Custom range...":
            self.custom_date_widget.show()
        else:
            self.custom_date_widget.hide()
        self._emit_filter()

    def _quick_unprocessed(self, checked):
        if checked:
            self.quick_recent.setChecked(False)
            self.status_all.setChecked(False)
            self.status_unprocessed.setChecked(True)
        else:
            self.status_all.setChecked(True)
        self._emit_filter()

    def _quick_recent(self, checked):
        if checked:
            self.quick_unprocessed.setChecked(False)
            self.date_filter.setCurrentText("Last 7 days")
        else:
            self.date_filter.setCurrentText("Any time")
        self._emit_filter()

    def _clear_filters(self):
        # Reset all filters to default
        self.type_pdf.setChecked(True)
        self.type_doc.setChecked(True)
        self.type_img.setChecked(True)
        self.type_other.setChecked(True)
        self.status_all.setChecked(True)
        self.date_filter.setCurrentIndex(0)
        self.tags_combo.setCurrentIndex(0)
        self.quick_unprocessed.setChecked(False)
        self.quick_recent.setChecked(False)
        self._emit_filter()

    def set_available_tags(self, tags):
        """Update the tags dropdown with available tags."""
        current = self.tags_combo.currentText()
        self.tags_combo.clear()
        self.tags_combo.addItem("All Tags")
        for tag in tags:
            self.tags_combo.addItem(tag)

        # Restore selection if still valid
        idx = self.tags_combo.findText(current)
        if idx >= 0:
            self.tags_combo.setCurrentIndex(idx)

    def get_filters(self):
        """Get current filter settings as a dictionary."""
        filters = {
            "file_types": [],
            "processing_status": [],
            "date_filter": self.date_filter.currentText(),
            "date_from": self.date_from.date().toString("yyyy-MM-dd") if self.custom_date_widget.isVisible() else None,
            "date_to": self.date_to.date().toString("yyyy-MM-dd") if self.custom_date_widget.isVisible() else None,
            "tag": self.tags_combo.currentText() if self.tags_combo.currentText() != "All Tags" else None
        }

        if self.type_pdf.isChecked():
            filters["file_types"].append("pdf")
        if self.type_doc.isChecked():
            filters["file_types"].extend(["doc", "docx"])
        if self.type_img.isChecked():
            filters["file_types"].extend(["jpg", "jpeg", "png", "gif", "bmp", "tiff"])
        if self.type_other.isChecked():
            filters["file_types"].append("other")

        if self.status_all.isChecked():
            filters["processing_status"] = ["all"]
        else:
            if self.status_unprocessed.isChecked():
                filters["processing_status"].append("unprocessed")
            if self.status_ocr.isChecked():
                filters["processing_status"].append("ocr")
            if self.status_summarized.isChecked():
                filters["processing_status"].append("summarized")

        return filters


# =============================================================================
# File Preview Widget
# =============================================================================

class FilePreviewWidget(QFrame):
    """Split view preview pane for files."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.current_file = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)

        # Header
        header_layout = QHBoxLayout()
        self.title_label = QLabel("Preview")
        self.title_label.setStyleSheet("font-weight: bold;")
        header_layout.addWidget(self.title_label)
        header_layout.addStretch()

        self.open_btn = QPushButton("Open")
        self.open_btn.setFixedSize(60, 24)
        self.open_btn.clicked.connect(self._open_file)
        self.open_btn.setEnabled(False)
        header_layout.addWidget(self.open_btn)

        self.close_btn = QPushButton("×")
        self.close_btn.setFixedSize(24, 24)
        self.close_btn.setStyleSheet("font-weight: bold;")
        self.close_btn.clicked.connect(self.hide)
        header_layout.addWidget(self.close_btn)

        layout.addLayout(header_layout)

        # File info
        self.info_label = QLabel()
        self.info_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(self.info_label)

        # Preview content area
        self.preview_stack = QStackedWidget()

        # Text preview
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.text_preview.setStyleSheet("font-family: Consolas; font-size: 10px;")
        self.preview_stack.addWidget(self.text_preview)

        # Image preview
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setStyleSheet("background-color: #f0f0f0;")
        scroll = QScrollArea()
        scroll.setWidget(self.image_label)
        scroll.setWidgetResizable(True)
        self.preview_stack.addWidget(scroll)

        # No preview placeholder
        self.no_preview_label = QLabel("No preview available.\nDouble-click to open file.")
        self.no_preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.no_preview_label.setStyleSheet("color: gray;")
        self.preview_stack.addWidget(self.no_preview_label)

        layout.addWidget(self.preview_stack)

        self.preview_stack.setCurrentWidget(self.no_preview_label)

    def show_file(self, file_path):
        """Display preview for the given file."""
        self.current_file = file_path
        self.open_btn.setEnabled(True)

        filename = os.path.basename(file_path)
        self.title_label.setText(f"Preview: {filename[:30]}...")

        # File info
        try:
            stat = os.stat(file_path)
            size = stat.st_size
            size_str = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"
            mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%m/%d/%Y %H:%M")
            self.info_label.setText(f"Size: {size_str} | Modified: {mtime}")
        except:
            self.info_label.setText("")

        ext = os.path.splitext(file_path)[1].lower()

        # Handle different file types
        if ext in ['.txt', '.md', '.json', '.xml', '.html', '.css', '.js', '.py']:
            self._show_text_preview(file_path)
        elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
            self._show_image_preview(file_path)
        elif ext == '.pdf':
            self._show_pdf_info(file_path)
        else:
            self.preview_stack.setCurrentWidget(self.no_preview_label)

    def _show_text_preview(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(50000)  # Limit to first 50KB
            self.text_preview.setPlainText(content)
            self.preview_stack.setCurrentWidget(self.text_preview)
        except Exception as e:
            self.no_preview_label.setText(f"Error reading file:\n{str(e)}")
            self.preview_stack.setCurrentWidget(self.no_preview_label)

    def _show_image_preview(self, file_path):
        from PySide6.QtGui import QPixmap
        pixmap = QPixmap(file_path)
        if not pixmap.isNull():
            # Scale to fit
            scaled = pixmap.scaled(
                400, 400,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            self.image_label.setPixmap(scaled)
            self.preview_stack.setCurrentIndex(1)
        else:
            self.preview_stack.setCurrentWidget(self.no_preview_label)

    def _show_pdf_info(self, file_path):
        """Show PDF info and first page text if available."""
        info_text = f"PDF File: {os.path.basename(file_path)}\n\n"

        # Try to get page count
        try:
            page_count = self._get_pdf_page_count(file_path)
            if page_count:
                info_text += f"Pages: {page_count}\n\n"
        except:
            pass

        # Check for existing OCR text
        text_path = file_path.replace('.pdf', '_ocr.txt').replace('.PDF', '_ocr.txt')
        if os.path.exists(text_path):
            try:
                with open(text_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read(10000)
                info_text += "--- OCR Text Preview ---\n\n"
                info_text += content
            except:
                pass
        else:
            info_text += "(No OCR text available - Run OCR to enable text preview)"

        self.text_preview.setPlainText(info_text)
        self.preview_stack.setCurrentWidget(self.text_preview)

    def _get_pdf_page_count(self, file_path):
        """Get PDF page count without external dependencies."""
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
            # Simple regex to find page count in PDF
            matches = re.findall(rb'/Type\s*/Page[^s]', content)
            return len(matches) if matches else None
        except:
            return None

    def _open_file(self):
        if self.current_file and os.path.exists(self.current_file):
            try:
                os.startfile(self.current_file)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def clear(self):
        self.current_file = None
        self.title_label.setText("Preview")
        self.info_label.setText("")
        self.open_btn.setEnabled(False)
        self.preview_stack.setCurrentWidget(self.no_preview_label)


# =============================================================================
# Output Browser Widget
# =============================================================================

class OutputBrowserWidget(QDialog):
    """Dialog for browsing, comparing, searching, and exporting outputs."""

    def __init__(self, case_path, file_number, parent=None):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.setWindowTitle(f"Output Browser - {file_number}")
        self.resize(1000, 700)

        layout = QVBoxLayout(self)

        # Tabs for different views
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Browse tab
        self.browse_tab = self._create_browse_tab()
        self.tabs.addTab(self.browse_tab, "Browse Outputs")

        # Compare tab
        self.compare_tab = self._create_compare_tab()
        self.tabs.addTab(self.compare_tab, "Compare")

        # Search tab
        self.search_tab = self._create_search_tab()
        self.tabs.addTab(self.search_tab, "Search")

        # Export tab
        self.export_tab = self._create_export_tab()
        self.tabs.addTab(self.export_tab, "Export")

        # Load outputs
        self._load_outputs()

    def _create_browse_tab(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)

        # Left: Output list
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        left_layout.addWidget(QLabel("<b>Generated Outputs</b>"))

        self.output_list = QListWidget()
        self.output_list.itemSelectionChanged.connect(self._on_output_selected)
        left_layout.addWidget(self.output_list)

        # Filter by type
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Type:"))
        self.type_filter = QComboBox()
        self.type_filter.addItems(["All", "Summaries", "Reports", "Timelines", "OCR Text"])
        self.type_filter.currentIndexChanged.connect(self._filter_outputs)
        filter_layout.addWidget(self.type_filter)
        left_layout.addLayout(filter_layout)

        layout.addWidget(left_widget, 1)

        # Right: Preview
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        right_layout.addWidget(QLabel("<b>Preview</b>"))

        self.output_preview = QTextEdit()
        self.output_preview.setReadOnly(True)
        self.output_preview.setStyleSheet("font-family: Consolas;")
        right_layout.addWidget(self.output_preview)

        # Open button
        self.open_output_btn = QPushButton("Open in External Editor")
        self.open_output_btn.clicked.connect(self._open_selected_output)
        right_layout.addWidget(self.open_output_btn)

        layout.addWidget(right_widget, 2)

        return widget

    def _create_compare_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Selection
        select_layout = QHBoxLayout()
        select_layout.addWidget(QLabel("Left:"))
        self.compare_left = QComboBox()
        select_layout.addWidget(self.compare_left)
        select_layout.addWidget(QLabel("Right:"))
        self.compare_right = QComboBox()
        select_layout.addWidget(self.compare_right)

        compare_btn = QPushButton("Compare")
        compare_btn.clicked.connect(self._do_compare)
        select_layout.addWidget(compare_btn)

        layout.addLayout(select_layout)

        # Comparison view
        compare_split = QSplitter(Qt.Orientation.Horizontal)

        self.compare_left_text = QTextEdit()
        self.compare_left_text.setReadOnly(True)
        compare_split.addWidget(self.compare_left_text)

        self.compare_right_text = QTextEdit()
        self.compare_right_text.setReadOnly(True)
        compare_split.addWidget(self.compare_right_text)

        layout.addWidget(compare_split)

        return widget

    def _create_search_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Search input
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search within outputs...")
        self.search_input.returnPressed.connect(self._do_search)
        search_layout.addWidget(self.search_input)

        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self._do_search)
        search_layout.addWidget(search_btn)

        layout.addLayout(search_layout)

        # Results
        self.search_results = QListWidget()
        self.search_results.itemDoubleClicked.connect(self._open_search_result)
        layout.addWidget(self.search_results)

        return widget

    def _create_export_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        layout.addWidget(QLabel("<b>Export Options</b>"))

        # Checkboxes for what to include
        self.export_summaries = QCheckBox("Include Summaries")
        self.export_summaries.setChecked(True)
        layout.addWidget(self.export_summaries)

        self.export_reports = QCheckBox("Include Reports")
        self.export_reports.setChecked(True)
        layout.addWidget(self.export_reports)

        self.export_timelines = QCheckBox("Include Timelines")
        self.export_timelines.setChecked(True)
        layout.addWidget(self.export_timelines)

        layout.addSpacing(20)

        # Format selection
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("Export Format:"))
        self.export_format = QComboBox()
        self.export_format.addItems(["Markdown (.md)", "Plain Text (.txt)", "Word Document (.docx)"])
        format_layout.addWidget(self.export_format)
        layout.addLayout(format_layout)

        layout.addStretch()

        # Export button
        export_btn = QPushButton("Export All Selected")
        export_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px;")
        export_btn.clicked.connect(self._do_export)
        layout.addWidget(export_btn)

        return widget

    def _load_outputs(self):
        """Load all generated outputs from AI OUTPUT folder."""
        self.outputs = []

        ai_output_dir = os.path.join(self.case_path, "NOTES", "AI OUTPUT")
        if not os.path.exists(ai_output_dir):
            return

        for filename in os.listdir(ai_output_dir):
            file_path = os.path.join(ai_output_dir, filename)
            if os.path.isfile(file_path):
                # Determine type
                lower = filename.lower()
                if 'summary' in lower or lower.endswith('.md'):
                    output_type = "Summaries"
                elif 'report' in lower:
                    output_type = "Reports"
                elif 'timeline' in lower:
                    output_type = "Timelines"
                elif '_ocr' in lower or lower.endswith('.txt'):
                    output_type = "OCR Text"
                else:
                    output_type = "Other"

                mtime = os.path.getmtime(file_path)
                self.outputs.append({
                    "name": filename,
                    "path": file_path,
                    "type": output_type,
                    "mtime": mtime
                })

        # Sort by modification time (newest first)
        self.outputs.sort(key=lambda x: x["mtime"], reverse=True)

        self._populate_output_list()
        self._populate_compare_combos()

    def _populate_output_list(self, type_filter=None):
        self.output_list.clear()

        for output in self.outputs:
            if type_filter and type_filter != "All" and output["type"] != type_filter:
                continue

            item = QListWidgetItem(f"[{output['type'][:3].upper()}] {output['name']}")
            item.setData(Qt.ItemDataRole.UserRole, output["path"])
            self.output_list.addItem(item)

    def _populate_compare_combos(self):
        self.compare_left.clear()
        self.compare_right.clear()

        for output in self.outputs:
            self.compare_left.addItem(output["name"], output["path"])
            self.compare_right.addItem(output["name"], output["path"])

    def _filter_outputs(self):
        type_filter = self.type_filter.currentText()
        self._populate_output_list(type_filter)

    def _on_output_selected(self):
        items = self.output_list.selectedItems()
        if not items:
            return

        path = items[0].data(Qt.ItemDataRole.UserRole)
        try:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            self.output_preview.setPlainText(content)
        except Exception as e:
            self.output_preview.setPlainText(f"Error reading file: {e}")

    def _open_selected_output(self):
        items = self.output_list.selectedItems()
        if items:
            path = items[0].data(Qt.ItemDataRole.UserRole)
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def _do_compare(self):
        left_path = self.compare_left.currentData()
        right_path = self.compare_right.currentData()

        if not left_path or not right_path:
            return

        try:
            with open(left_path, 'r', encoding='utf-8', errors='ignore') as f:
                left_content = f.read()
            with open(right_path, 'r', encoding='utf-8', errors='ignore') as f:
                right_content = f.read()

            self.compare_left_text.setPlainText(left_content)
            self.compare_right_text.setPlainText(right_content)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error reading files: {e}")

    def _do_search(self):
        query = self.search_input.text().strip().lower()
        if not query:
            return

        self.search_results.clear()

        for output in self.outputs:
            try:
                with open(output["path"], 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()

                if query in content.lower():
                    # Find context around match
                    idx = content.lower().find(query)
                    start = max(0, idx - 50)
                    end = min(len(content), idx + len(query) + 50)
                    context = "..." + content[start:end].replace('\n', ' ') + "..."

                    item = QListWidgetItem(f"{output['name']}: {context}")
                    item.setData(Qt.ItemDataRole.UserRole, output["path"])
                    self.search_results.addItem(item)
            except:
                pass

    def _open_search_result(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        try:
            os.startfile(path)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def _do_export(self):
        # Get export format
        format_text = self.export_format.currentText()
        if "Markdown" in format_text:
            ext = ".md"
        elif "Word" in format_text:
            ext = ".docx"
        else:
            ext = ".txt"

        # Choose save location
        default_name = f"{self.file_number}_combined_outputs{ext}"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Outputs",
            os.path.join(self.case_path, default_name),
            f"*{ext}"
        )

        if not save_path:
            return

        # Collect content
        combined = f"# Combined Outputs for Case {self.file_number}\n\n"
        combined += f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
        combined += "---\n\n"

        for output in self.outputs:
            include = False
            if self.export_summaries.isChecked() and output["type"] == "Summaries":
                include = True
            elif self.export_reports.isChecked() and output["type"] == "Reports":
                include = True
            elif self.export_timelines.isChecked() and output["type"] == "Timelines":
                include = True

            if include:
                try:
                    with open(output["path"], 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()

                    combined += f"## {output['name']}\n\n"
                    combined += content
                    combined += "\n\n---\n\n"
                except:
                    pass

        # Save
        try:
            if ext == ".docx":
                # For Word, we'd need python-docx - for now just save as txt
                save_path = save_path.replace(".docx", ".txt")

            with open(save_path, 'w', encoding='utf-8') as f:
                f.write(combined)

            QMessageBox.information(self, "Success", f"Exported to:\n{save_path}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Export failed: {e}")


# =============================================================================
# Processing Log Widget
# =============================================================================

class ProcessingLogWidget(QDialog):
    """Dialog for viewing processing history and errors."""

    def __init__(self, file_number, parent=None):
        super().__init__(parent)
        self.file_number = file_number
        self.log_db = ProcessingLogDB(file_number)

        self.setWindowTitle(f"Processing Log - {file_number}")
        self.resize(800, 500)

        layout = QVBoxLayout(self)

        # Filter bar
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Filter:"))

        self.status_filter = QComboBox()
        self.status_filter.addItems(["All", "Success", "Failed", "In Progress"])
        self.status_filter.currentIndexChanged.connect(self._apply_filter)
        filter_layout.addWidget(self.status_filter)

        self.task_filter = QComboBox()
        self.task_filter.addItems(["All Tasks", "OCR", "Summarize", "Timeline", "Other"])
        self.task_filter.currentIndexChanged.connect(self._apply_filter)
        filter_layout.addWidget(self.task_filter)

        filter_layout.addStretch()

        clear_btn = QPushButton("Clear Log")
        clear_btn.setStyleSheet("color: red;")
        clear_btn.clicked.connect(self._clear_log)
        filter_layout.addWidget(clear_btn)

        layout.addLayout(filter_layout)

        # Log list
        self.log_list = QListWidget()
        self.log_list.itemSelectionChanged.connect(self._show_details)
        layout.addWidget(self.log_list)

        # Details panel
        details_group = QGroupBox("Details")
        details_layout = QVBoxLayout(details_group)

        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setMaximumHeight(150)
        details_layout.addWidget(self.details_text)

        layout.addWidget(details_group)

        self._populate_log()

    def _populate_log(self, status_filter=None, task_filter=None):
        self.log_list.clear()

        for entry in self.log_db.get_all_logs():
            # Apply filters
            if status_filter and status_filter != "All":
                if entry["status"].lower() != status_filter.lower():
                    continue

            if task_filter and task_filter != "All Tasks":
                if task_filter.lower() not in entry["task_type"].lower():
                    continue

            # Format entry
            timestamp = entry["timestamp"][:16].replace("T", " ")
            status_icon = "✓" if entry["status"] == "success" else "✗" if entry["status"] == "failed" else "⋯"
            text = f"[{timestamp}] {status_icon} {entry['task_type']} - {entry['file_name']}"

            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, entry)

            # Color by status
            if entry["status"] == "success":
                item.setForeground(QColor("#4CAF50"))
            elif entry["status"] == "failed":
                item.setForeground(QColor("#f44336"))

            self.log_list.addItem(item)

    def _apply_filter(self):
        status = self.status_filter.currentText()
        task = self.task_filter.currentText()
        self._populate_log(status, task)

    def _show_details(self):
        items = self.log_list.selectedItems()
        if not items:
            return

        entry = items[0].data(Qt.ItemDataRole.UserRole)

        details = f"File: {entry['file_path']}\n"
        details += f"Task: {entry['task_type']}\n"
        details += f"Status: {entry['status']}\n"
        details += f"Time: {entry['timestamp']}\n"

        if entry.get("duration_sec"):
            details += f"Duration: {entry['duration_sec']:.1f} seconds\n"

        if entry.get("output_path"):
            details += f"Output: {entry['output_path']}\n"

        if entry.get("error_message"):
            details += f"\n--- Error Details ---\n{entry['error_message']}\n"

            # Add suggested fixes for common errors
            error = entry["error_message"].lower()
            if "permission" in error:
                details += "\n💡 Suggested Fix: Close the file if it's open in another application."
            elif "memory" in error:
                details += "\n💡 Suggested Fix: Try processing smaller files or close other applications."
            elif "timeout" in error:
                details += "\n💡 Suggested Fix: The file may be too large. Try splitting it first."

        self.details_text.setPlainText(details)

    def _clear_log(self):
        confirm = QMessageBox.question(
            self,
            "Confirm Clear",
            "Are you sure you want to clear the processing log?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            self.log_db.clear_logs()
            self._populate_log()


# =============================================================================
# Enhanced File Tree Widget
# =============================================================================

class EnhancedFileTreeWidget(QTreeWidget):
    """Enhanced file tree with additional columns and features."""

    item_moved = Signal(str, str)  # old_path, new_folder_path
    rename_requested = Signal(str, str)  # old_path, new_name
    folder_created = Signal(str)  # new_folder_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.tags_db = None
        self.processing_log = None

        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_context_menu)

        # Enable editing for rename
        self.setEditTriggers(QAbstractItemView.EditTrigger.EditKeyPressed)
        self.itemChanged.connect(self._on_item_renamed)

        self._editing_item = None
        self._original_name = None

    def set_databases(self, file_number):
        """Set up databases for tags and processing log."""
        self.tags_db = FileTagsDB(file_number)
        self.processing_log = ProcessingLogDB(file_number)

    def open_context_menu(self, position):
        item = self.itemAt(position)

        menu = QMenu()

        if item:
            path = item.data(0, Qt.ItemDataRole.UserRole)
            item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)

            if path:
                open_action = QAction("Open", self)
                open_action.triggered.connect(lambda: self._open_file(path))
                menu.addAction(open_action)

                reveal_action = QAction("Reveal in Explorer", self)
                reveal_action.triggered.connect(lambda: self._reveal_in_explorer(path))
                menu.addAction(reveal_action)

                copy_path_action = QAction("Copy Path", self)
                copy_path_action.triggered.connect(lambda: QApplication.clipboard().setText(path))
                menu.addAction(copy_path_action)

                menu.addSeparator()

                # Rename
                rename_action = QAction("Rename", self)
                rename_action.triggered.connect(lambda: self._start_rename(item))
                menu.addAction(rename_action)

                # Tags submenu
                if item_type == "file":
                    tags_menu = QMenu("Tags", self)

                    current_tags = self.tags_db.get_tags(path) if self.tags_db else []

                    # Add tag action
                    add_tag_action = QAction("Add Tag...", self)
                    add_tag_action.triggered.connect(lambda: self._add_tag(path))
                    tags_menu.addAction(add_tag_action)

                    # Remove tag actions
                    if current_tags:
                        tags_menu.addSeparator()
                        for tag in current_tags:
                            remove_action = QAction(f"Remove: {tag}", self)
                            remove_action.triggered.connect(partial(self._remove_tag, path, tag))
                            tags_menu.addAction(remove_action)

                    menu.addMenu(tags_menu)

                menu.addSeparator()

                # Create folder (for directories)
                if item_type == "dir":
                    create_folder_action = QAction("Create Subfolder", self)
                    create_folder_action.triggered.connect(lambda: self._create_folder(path))
                    menu.addAction(create_folder_action)
        else:
            # Right-clicked on empty space - allow creating folder at root
            pass

        menu.exec(self.viewport().mapToGlobal(position))

    def _open_file(self, path):
        try:
            os.startfile(path)
        except Exception as e:
            log_event(f"Error opening file: {e}", "error")

    def _reveal_in_explorer(self, path):
        try:
            import subprocess
            subprocess.run(['explorer', '/select,', path])
        except Exception as e:
            log_event(f"Error revealing file: {e}", "error")

    def _start_rename(self, item):
        """Start inline rename for an item."""
        self._editing_item = item
        self._original_name = item.text(0)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
        self.editItem(item, 0)

    def _on_item_renamed(self, item, column):
        """Handle item rename completion."""
        if item != self._editing_item or column != 0:
            return

        new_name = item.text(0)
        if new_name == self._original_name:
            self._editing_item = None
            return

        path = item.data(0, Qt.ItemDataRole.UserRole)
        if path:
            new_path = os.path.join(os.path.dirname(path), new_name)

            try:
                os.rename(path, new_path)
                item.setData(0, Qt.ItemDataRole.UserRole, new_path)
                self.rename_requested.emit(path, new_name)
                log_event(f"Renamed {path} to {new_path}")
            except Exception as e:
                item.setText(0, self._original_name)
                QMessageBox.warning(self, "Error", f"Could not rename: {e}")

        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self._editing_item = None

    def _add_tag(self, path):
        """Add a tag to a file."""
        tag, ok = QInputDialog.getText(self, "Add Tag", "Enter tag name:")
        if ok and tag.strip():
            if self.tags_db:
                self.tags_db.add_tag(path, tag.strip())
                # Refresh the item's tag display
                self._update_tags_column(path)

    def _remove_tag(self, path, tag):
        """Remove a tag from a file."""
        if self.tags_db:
            self.tags_db.remove_tag(path, tag)
            self._update_tags_column(path)

    def _update_tags_column(self, path):
        """Update the tags column for a specific file."""
        # Find the item with this path
        iterator = QTreeWidgetItemIterator(self)
        while iterator.value():
            item = iterator.value()
            if item.data(0, Qt.ItemDataRole.UserRole) == path:
                tags = self.tags_db.get_tags(path) if self.tags_db else []
                item.setText(5, ", ".join(tags) if tags else "")
                break
            iterator += 1

    def _create_folder(self, parent_path):
        """Create a new subfolder."""
        name, ok = QInputDialog.getText(self, "Create Folder", "Folder name:")
        if ok and name.strip():
            new_path = os.path.join(parent_path, name.strip())
            try:
                os.makedirs(new_path, exist_ok=True)
                self.folder_created.emit(new_path)
                log_event(f"Created folder: {new_path}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not create folder: {e}")

    def get_processing_status(self, file_path):
        """Get processing status for a file."""
        if not self.processing_log:
            return ""

        logs = self.processing_log.get_file_processing_status(file_path)
        if not logs:
            return ""

        # Get unique successful task types
        successful = set()
        for log in logs:
            if log["status"] == "success":
                task = log["task_type"].lower()
                if "ocr" in task:
                    successful.add("OCR")
                elif "summar" in task:
                    successful.add("SUM")
                elif "timeline" in task:
                    successful.add("TL")

        return ", ".join(sorted(successful)) if successful else ""

    def dropEvent(self, event):
        """Handle drop events for file moving."""
        target_item = self.itemAt(event.position().toPoint())
        if not target_item:
            event.ignore()
            return

        target_path = target_item.data(0, Qt.ItemDataRole.UserRole)
        target_type = target_item.data(0, Qt.ItemDataRole.UserRole + 1)

        if target_type == "file":
            target_folder = os.path.dirname(target_path)
        else:
            target_folder = target_path

        # Handle External Files (URLs)
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            external_paths = [u.toLocalFile() for u in urls if u.isLocalFile()]
            if external_paths:
                event.accept()
                copied_any = False
                for src in external_paths:
                    if os.path.exists(src):
                        dst = os.path.join(target_folder, os.path.basename(src))
                        try:
                            if os.path.isdir(src):
                                shutil.copytree(src, dst)
                            else:
                                shutil.copy2(src, dst)
                            copied_any = True
                        except Exception as e:
                            log_event(f"Error copying {src}: {e}", "error")

                if copied_any:
                    self.item_moved.emit("", "")
                return

        # Handle Internal Moves
        items = self.selectedItems()
        moved_any = False
        for item in items:
            old_path = item.data(0, Qt.ItemDataRole.UserRole)
            if not old_path or not os.path.exists(old_path):
                continue

            filename = os.path.basename(old_path)
            new_path = os.path.join(target_folder, filename)

            if old_path == new_path:
                continue

            try:
                os.rename(old_path, new_path)
                moved_any = True
            except Exception as e:
                log_event(f"Failed to move {filename}: {e}", "error")

        if moved_any:
            self.item_moved.emit("", "")
            event.accept()
        else:
            super().dropEvent(event)
