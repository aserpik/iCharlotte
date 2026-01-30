"""
Enhanced Case View Tab Components

This module contains improved UI components for the Case View Tab including:
- Enhanced Agent Buttons with running indicator, last docket date, and settings
- Enhanced File Tree with processing status and tags columns
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
            "date_format": "MM/DD/YYYY",
            "from_summaries": True,  # Extract from summaries by default
            "selected_summaries": None  # None means all summaries
        },
        "detect_contradictions.py": {
            "selected_summaries": None  # None means all summaries
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
# Contradiction Detector Settings Dialog
# =============================================================================

class ContradictionSettingsDialog(QDialog):
    """Custom settings dialog for the Contradiction Detector agent."""

    def __init__(self, file_number, settings_db, parent=None):
        super().__init__(parent)
        self.file_number = file_number
        self.settings_db = settings_db
        self.script_name = "detect_contradictions.py"
        self.setWindowTitle("Contradiction Detector Settings")
        self.setMinimumSize(600, 500)

        # Load current settings
        self.current_settings = self.settings_db.get_settings(self.script_name)

        # Load prompt from file
        self.prompt_path = os.path.join(SCRIPTS_DIR, "CONTRADICTION_DETECTION_PROMPT.txt")
        self.original_prompt = self._load_prompt()

        self._setup_ui()
        self._load_summaries()

    def _load_prompt(self):
        """Load the prompt from file."""
        try:
            if os.path.exists(self.prompt_path):
                with open(self.prompt_path, 'r', encoding='utf-8') as f:
                    return f.read()
        except Exception as e:
            log_event(f"Error loading prompt: {e}", "error")
        return ""

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # Create tab widget for organization
        tabs = QTabWidget()
        layout.addWidget(tabs)

        # --- Tab 1: Summary Selection ---
        summary_tab = QWidget()
        summary_layout = QVBoxLayout(summary_tab)

        # Header with toggle and delete buttons
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>Select Summaries to Analyze:</b>"))
        header_layout.addStretch()

        self.toggle_btn = QPushButton("Deselect All")
        self.toggle_btn.setFixedWidth(100)
        self.toggle_btn.clicked.connect(self._toggle_selection)
        header_layout.addWidget(self.toggle_btn)

        self.delete_btn = QPushButton("Delete Selected")
        self.delete_btn.setFixedWidth(110)
        self.delete_btn.setStyleSheet("background-color: #f44336; color: white;")
        self.delete_btn.clicked.connect(self._delete_selected)
        header_layout.addWidget(self.delete_btn)

        summary_layout.addLayout(header_layout)

        # Summary list with checkboxes
        self.summary_list = QListWidget()
        self.summary_list.setAlternatingRowColors(True)
        self.summary_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.summary_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.summary_list.customContextMenuRequested.connect(self._show_context_menu)
        self.summary_list.installEventFilter(self)
        summary_layout.addWidget(self.summary_list)

        # Info label
        self.info_label = QLabel("Loading summaries...")
        self.info_label.setStyleSheet("color: #666; font-style: italic;")
        summary_layout.addWidget(self.info_label)

        tabs.addTab(summary_tab, "Summaries")

        # --- Tab 2: Prompt Editor ---
        prompt_tab = QWidget()
        prompt_layout = QVBoxLayout(prompt_tab)

        prompt_header = QHBoxLayout()
        prompt_header.addWidget(QLabel("<b>Detection Prompt:</b>"))
        prompt_header.addStretch()

        reset_prompt_btn = QPushButton("Reset to Default")
        reset_prompt_btn.clicked.connect(self._reset_prompt)
        prompt_header.addWidget(reset_prompt_btn)

        prompt_layout.addLayout(prompt_header)

        self.prompt_edit = QTextEdit()
        self.prompt_edit.setPlainText(self.original_prompt)
        self.prompt_edit.setStyleSheet("font-family: Consolas, monospace; font-size: 11px;")
        prompt_layout.addWidget(self.prompt_edit)

        tabs.addTab(prompt_tab, "Prompt")

        # --- Buttons ---
        btn_layout = QHBoxLayout()

        reset_all_btn = QPushButton("Reset All to Defaults")
        reset_all_btn.clicked.connect(self._reset_all)
        btn_layout.addWidget(reset_all_btn)

        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px 16px;")
        save_btn.clicked.connect(self._save)
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)

    def _load_summaries(self):
        """Load available summaries from Document Registry, AI_OUTPUT.docx, and CaseDataManager."""
        self.summary_list.clear()

        if not self.file_number:
            self.info_label.setText("No case selected. Open a case first.")
            return

        try:
            import sys
            sys.path.insert(0, SCRIPTS_DIR)

            # Get previously selected summaries from settings
            selected_summaries = self.current_settings.get("selected_summaries", None)
            summary_count = 0
            seen_names = set()

            # --- Source 1: Document Registry (preferred) ---
            try:
                from document_registry import get_available_documents
                registry_docs = get_available_documents(self.file_number)

                for doc in registry_docs:
                    name = doc['name']
                    if name in seen_names:
                        continue
                    seen_names.add(name)

                    item = QListWidgetItem(f"{name}")
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setData(Qt.ItemDataRole.UserRole, name)  # Store clean name

                    if selected_summaries is None:
                        item.setCheckState(Qt.CheckState.Checked)
                    else:
                        item.setCheckState(
                            Qt.CheckState.Checked if name in selected_summaries else Qt.CheckState.Unchecked
                        )

                    doc_type = doc.get('document_type', 'Unknown')
                    agent = doc.get('agent', 'Unknown')
                    item.setToolTip(f"Type: {doc_type}\nAgent: {agent}\nChars: {doc.get('char_count', 'N/A')}")

                    # Color-code by document type
                    if 'Deposition' in doc_type:
                        item.setForeground(QColor("#2196F3"))  # Blue
                    elif 'Interrogatories' in doc_type or 'Request' in doc_type:
                        item.setForeground(QColor("#4CAF50"))  # Green
                    elif 'Medical' in doc_type:
                        item.setForeground(QColor("#FF9800"))  # Orange

                    self.summary_list.addItem(item)
                    summary_count += 1

            except ImportError:
                log_event("Document registry not available", "warning")

            # --- Source 2: AI_OUTPUT.docx ---
            try:
                from detect_contradictions import find_ai_output_docx, gather_summaries_from_docx
                from icharlotte_core.agent_logger import AgentLogger

                docx_path = find_ai_output_docx(self.file_number)
                if docx_path:
                    # Create a simple logger that does nothing
                    class SilentLogger:
                        def info(self, msg): pass
                        def warning(self, msg): pass
                        def error(self, msg): pass

                    docx_summaries = gather_summaries_from_docx(docx_path, SilentLogger())

                    for name in docx_summaries.keys():
                        if name in seen_names:
                            continue
                        seen_names.add(name)

                        item = QListWidgetItem(f"{name}")
                        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                        item.setData(Qt.ItemDataRole.UserRole, name)

                        if selected_summaries is None:
                            item.setCheckState(Qt.CheckState.Checked)
                        else:
                            item.setCheckState(
                                Qt.CheckState.Checked if name in selected_summaries else Qt.CheckState.Unchecked
                            )

                        item.setToolTip(f"Source: AI_OUTPUT.docx\nLength: {len(docx_summaries[name])} chars")
                        item.setForeground(QColor("#9C27B0"))  # Purple for DOCX-only

                        self.summary_list.addItem(item)
                        summary_count += 1

            except Exception as e:
                log_event(f"Error loading from AI_OUTPUT.docx: {e}", "warning")

            # --- Source 3: CaseDataManager (fallback for other summaries) ---
            try:
                from case_data_manager import CaseDataManager

                data_manager = CaseDataManager()
                variables = data_manager.get_all_variables(self.file_number, flatten=False)

                for var_name, var_data in variables.items():
                    if var_name in seen_names:
                        continue

                    # Filter for summary-type variables
                    if any(tag in var_name.lower() for tag in ['summary', 'depo', 'extraction']):
                        value = var_data.get('value', '') if isinstance(var_data, dict) else var_data
                        if value and len(str(value)) > 100:
                            seen_names.add(var_name)

                            item = QListWidgetItem(var_name)
                            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                            item.setData(Qt.ItemDataRole.UserRole, var_name)

                            if selected_summaries is None:
                                item.setCheckState(Qt.CheckState.Checked)
                            else:
                                item.setCheckState(
                                    Qt.CheckState.Checked if var_name in selected_summaries else Qt.CheckState.Unchecked
                                )

                            source = var_data.get('source', 'Unknown') if isinstance(var_data, dict) else 'legacy'
                            item.setToolTip(f"Source: {source}\nLength: {len(str(value))} chars")
                            item.setForeground(QColor("#607D8B"))  # Gray for case data

                            self.summary_list.addItem(item)
                            summary_count += 1

            except Exception as e:
                log_event(f"Error loading from CaseDataManager: {e}", "warning")

            if summary_count == 0:
                self.info_label.setText("No summaries found. Run summarization agents first.")
            else:
                self.info_label.setText(f"Found {summary_count} summaries for case {self.file_number}")
                self._update_toggle_button()

        except Exception as e:
            log_event(f"Error loading summaries: {e}", "error")
            self.info_label.setText(f"Error loading summaries: {e}")

    def _toggle_selection(self):
        """Toggle between select all and deselect all."""
        # Check if all are currently selected
        all_checked = all(
            self.summary_list.item(i).checkState() == Qt.CheckState.Checked
            for i in range(self.summary_list.count())
        )

        new_state = Qt.CheckState.Unchecked if all_checked else Qt.CheckState.Checked

        for i in range(self.summary_list.count()):
            self.summary_list.item(i).setCheckState(new_state)

        self._update_toggle_button()

    def _update_toggle_button(self):
        """Update toggle button text based on current selection."""
        if self.summary_list.count() == 0:
            return

        all_checked = all(
            self.summary_list.item(i).checkState() == Qt.CheckState.Checked
            for i in range(self.summary_list.count())
        )

        self.toggle_btn.setText("Deselect All" if all_checked else "Select All")

    def _reset_prompt(self):
        """Reset prompt to the file default."""
        self.prompt_edit.setPlainText(self.original_prompt)

    def _reset_all(self):
        """Reset all settings to defaults."""
        # Reset prompt
        self._reset_prompt()

        # Reset summary selection (all checked)
        for i in range(self.summary_list.count()):
            self.summary_list.item(i).setCheckState(Qt.CheckState.Checked)
        self._update_toggle_button()

    def _save(self):
        """Save all settings."""
        # Gather selected summaries (use stored name from UserRole if available)
        selected_summaries = []
        for i in range(self.summary_list.count()):
            item = self.summary_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                # Prefer UserRole data (clean name), fallback to display text
                name = item.data(Qt.ItemDataRole.UserRole) or item.text()
                selected_summaries.append(name)

        if not selected_summaries:
            QMessageBox.warning(self, "No Selection", "Please select at least one summary to analyze.")
            return

        # Save settings (model is configured in LLM Settings)
        settings = {
            "selected_summaries": selected_summaries,
        }
        self.settings_db.update_settings(self.script_name, settings)

        # Save prompt if modified
        current_prompt = self.prompt_edit.toPlainText()
        if current_prompt != self.original_prompt:
            try:
                with open(self.prompt_path, 'w', encoding='utf-8') as f:
                    f.write(current_prompt)
                log_event("Saved updated contradiction detection prompt")
            except Exception as e:
                QMessageBox.warning(self, "Warning", f"Could not save prompt: {e}")

        self.accept()

    def eventFilter(self, obj, event):
        """Handle key press events for delete functionality."""
        if obj == self.summary_list and event.type() == event.Type.KeyPress:
            if event.key() == Qt.Key.Key_Delete:
                self._delete_selected()
                return True
        return super().eventFilter(obj, event)

    def _show_context_menu(self, position):
        """Show context menu for the summary list."""
        menu = QMenu(self)

        delete_action = menu.addAction("Delete Selected")
        delete_action.triggered.connect(self._delete_selected)

        # Only show if there are selected items
        selected_items = self.summary_list.selectedItems()
        if not selected_items:
            delete_action.setEnabled(False)

        menu.exec(self.summary_list.mapToGlobal(position))

    def _delete_selected(self):
        """Delete selected items from the list and document registry."""
        selected_items = self.summary_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select items to delete (click to select, Ctrl+click for multiple).")
            return

        # Confirm deletion
        count = len(selected_items)
        reply = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Delete {count} item(s) from the document registry?\n\nThis will remove them from the list but NOT delete the actual summary files.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        # Delete from registry and list
        try:
            from document_registry import DocumentRegistry
            registry = DocumentRegistry()

            deleted_count = 0
            for item in selected_items:
                name = item.data(Qt.ItemDataRole.UserRole) or item.text()

                # Remove from document registry
                if registry.remove_document(self.file_number, name):
                    deleted_count += 1
                    log_event(f"Removed '{name}' from document registry")

                # Remove from list widget
                row = self.summary_list.row(item)
                self.summary_list.takeItem(row)

            # Update info label
            remaining = self.summary_list.count()
            self.info_label.setText(f"Found {remaining} summaries for case {self.file_number} ({deleted_count} deleted)")
            self._update_toggle_button()

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error deleting items: {e}")
            log_event(f"Error deleting from registry: {e}", "error")

    def get_selected_summaries(self):
        """Return list of selected summary names."""
        selected = []
        for i in range(self.summary_list.count()):
            item = self.summary_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected.append(item.text())
        return selected


# =============================================================================
# Timeline Extraction Settings Dialog
# =============================================================================

class TimelineSettingsDialog(QDialog):
    """Custom settings dialog for the Timeline Extraction agent."""

    def __init__(self, file_number, settings_db, parent=None):
        super().__init__(parent)
        self.file_number = file_number
        self.settings_db = settings_db
        self.script_name = "extract_timeline.py"
        self.setWindowTitle("Timeline Extraction Settings")
        self.setMinimumSize(600, 500)

        # Load current settings
        self.current_settings = self.settings_db.get_settings(self.script_name)

        # Load prompt from file
        self.prompt_path = os.path.join(SCRIPTS_DIR, "TIMELINE_EXTRACTION_PROMPT.txt")
        self.original_prompt = self._load_prompt()

        self._setup_ui()
        self._load_summaries()

    def _load_prompt(self):
        """Load the prompt from file."""
        try:
            if os.path.exists(self.prompt_path):
                with open(self.prompt_path, 'r', encoding='utf-8') as f:
                    return f.read()
        except Exception as e:
            log_event(f"Error loading prompt: {e}", "error")
        return ""

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # Mode selection
        mode_group = QGroupBox("Extraction Mode")
        mode_layout = QVBoxLayout(mode_group)

        self.mode_files = QCheckBox("Extract from selected files (traditional)")
        self.mode_summaries = QCheckBox("Extract from existing summaries (faster, uses AI_OUTPUT.docx)")
        self.mode_summaries.setChecked(self.current_settings.get("from_summaries", True))
        self.mode_files.setChecked(not self.current_settings.get("from_summaries", True))

        # Make them mutually exclusive
        self.mode_files.toggled.connect(lambda checked: self.mode_summaries.setChecked(not checked) if checked else None)
        self.mode_summaries.toggled.connect(lambda checked: self.mode_files.setChecked(not checked) if checked else None)
        self.mode_summaries.toggled.connect(self._on_mode_changed)

        mode_layout.addWidget(self.mode_files)
        mode_layout.addWidget(self.mode_summaries)
        layout.addWidget(mode_group)

        # Create tab widget
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # --- Tab 1: Summary Selection ---
        summary_tab = QWidget()
        summary_layout = QVBoxLayout(summary_tab)

        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>Select Documents to Include:</b>"))
        header_layout.addStretch()

        self.toggle_btn = QPushButton("Deselect All")
        self.toggle_btn.setFixedWidth(100)
        self.toggle_btn.clicked.connect(self._toggle_selection)
        header_layout.addWidget(self.toggle_btn)

        self.delete_btn = QPushButton("Delete Selected")
        self.delete_btn.setFixedWidth(110)
        self.delete_btn.setStyleSheet("background-color: #f44336; color: white;")
        self.delete_btn.clicked.connect(self._delete_selected)
        header_layout.addWidget(self.delete_btn)

        summary_layout.addLayout(header_layout)

        self.summary_list = QListWidget()
        self.summary_list.setAlternatingRowColors(True)
        self.summary_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.summary_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.summary_list.customContextMenuRequested.connect(self._show_context_menu)
        self.summary_list.installEventFilter(self)
        summary_layout.addWidget(self.summary_list)

        self.info_label = QLabel("Loading summaries...")
        self.info_label.setStyleSheet("color: #666; font-style: italic;")
        summary_layout.addWidget(self.info_label)

        self.tabs.addTab(summary_tab, "Documents")

        # --- Tab 2: Prompt Editor ---
        prompt_tab = QWidget()
        prompt_layout = QVBoxLayout(prompt_tab)

        prompt_header = QHBoxLayout()
        prompt_header.addWidget(QLabel("<b>Extraction Prompt:</b>"))
        prompt_header.addStretch()

        reset_prompt_btn = QPushButton("Reset to Default")
        reset_prompt_btn.clicked.connect(self._reset_prompt)
        prompt_header.addWidget(reset_prompt_btn)

        prompt_layout.addLayout(prompt_header)

        self.prompt_edit = QTextEdit()
        self.prompt_edit.setPlainText(self.original_prompt)
        self.prompt_edit.setStyleSheet("font-family: Consolas, monospace; font-size: 11px;")
        prompt_layout.addWidget(self.prompt_edit)

        self.tabs.addTab(prompt_tab, "Prompt")

        # --- Buttons ---
        btn_layout = QHBoxLayout()

        reset_all_btn = QPushButton("Reset All to Defaults")
        reset_all_btn.clicked.connect(self._reset_all)
        btn_layout.addWidget(reset_all_btn)

        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px 16px;")
        save_btn.clicked.connect(self._save)
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)

        # Initial state
        self._on_mode_changed(self.mode_summaries.isChecked())

    def _on_mode_changed(self, from_summaries):
        """Update UI based on mode selection."""
        self.tabs.setTabEnabled(0, from_summaries)  # Documents tab only for summary mode

    def _load_summaries(self):
        """Load available summaries from Document Registry, AI_OUTPUT.docx, and CaseDataManager."""
        self.summary_list.clear()

        if not self.file_number:
            self.info_label.setText("No case selected. Open a case first.")
            return

        try:
            import sys
            sys.path.insert(0, SCRIPTS_DIR)

            selected_summaries = self.current_settings.get("selected_summaries", None)
            summary_count = 0
            seen_names = set()

            # --- Source 1: Document Registry ---
            try:
                from document_registry import get_available_documents
                registry_docs = get_available_documents(self.file_number)

                for doc in registry_docs:
                    name = doc['name']
                    if name in seen_names:
                        continue
                    seen_names.add(name)

                    item = QListWidgetItem(f"{name}")
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setData(Qt.ItemDataRole.UserRole, name)

                    if selected_summaries is None:
                        item.setCheckState(Qt.CheckState.Checked)
                    else:
                        item.setCheckState(
                            Qt.CheckState.Checked if name in selected_summaries else Qt.CheckState.Unchecked
                        )

                    doc_type = doc.get('document_type', 'Unknown')
                    item.setToolTip(f"Type: {doc_type}\nAgent: {doc.get('agent', 'Unknown')}")

                    if 'Deposition' in doc_type:
                        item.setForeground(QColor("#2196F3"))
                    elif 'Interrogatories' in doc_type or 'Request' in doc_type:
                        item.setForeground(QColor("#4CAF50"))
                    elif 'Medical' in doc_type:
                        item.setForeground(QColor("#FF9800"))

                    self.summary_list.addItem(item)
                    summary_count += 1

            except ImportError:
                pass

            # --- Source 2: AI_OUTPUT.docx ---
            try:
                from extract_timeline import find_ai_output_docx, gather_summaries_from_docx
                from icharlotte_core.agent_logger import AgentLogger

                docx_path = find_ai_output_docx(self.file_number)
                if docx_path:
                    class SilentLogger:
                        def info(self, msg): pass
                        def warning(self, msg): pass
                        def error(self, msg): pass

                    docx_summaries = gather_summaries_from_docx(docx_path, SilentLogger())

                    for name in docx_summaries.keys():
                        if name in seen_names:
                            continue
                        seen_names.add(name)

                        item = QListWidgetItem(f"{name}")
                        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                        item.setData(Qt.ItemDataRole.UserRole, name)

                        if selected_summaries is None:
                            item.setCheckState(Qt.CheckState.Checked)
                        else:
                            item.setCheckState(
                                Qt.CheckState.Checked if name in selected_summaries else Qt.CheckState.Unchecked
                            )

                        item.setToolTip(f"Source: AI_OUTPUT.docx")
                        item.setForeground(QColor("#9C27B0"))

                        self.summary_list.addItem(item)
                        summary_count += 1

            except Exception as e:
                log_event(f"Error loading from AI_OUTPUT.docx: {e}", "warning")

            # --- Source 3: CaseDataManager ---
            try:
                from case_data_manager import CaseDataManager

                data_manager = CaseDataManager()
                variables = data_manager.get_all_variables(self.file_number, flatten=False)

                for var_name, var_data in variables.items():
                    if var_name in seen_names:
                        continue

                    if any(tag in var_name.lower() for tag in ['summary', 'depo', 'extraction']):
                        value = var_data.get('value', '') if isinstance(var_data, dict) else var_data
                        if value and len(str(value)) > 100:
                            seen_names.add(var_name)

                            item = QListWidgetItem(var_name)
                            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                            item.setData(Qt.ItemDataRole.UserRole, var_name)

                            if selected_summaries is None:
                                item.setCheckState(Qt.CheckState.Checked)
                            else:
                                item.setCheckState(
                                    Qt.CheckState.Checked if var_name in selected_summaries else Qt.CheckState.Unchecked
                                )

                            item.setToolTip(f"Source: Case Data")
                            item.setForeground(QColor("#607D8B"))

                            self.summary_list.addItem(item)
                            summary_count += 1

            except Exception as e:
                log_event(f"Error loading from CaseDataManager: {e}", "warning")

            if summary_count == 0:
                self.info_label.setText("No summaries found. Run summarization agents first.")
            else:
                self.info_label.setText(f"Found {summary_count} summaries for case {self.file_number}")
                self._update_toggle_button()

        except Exception as e:
            log_event(f"Error loading summaries: {e}", "error")
            self.info_label.setText(f"Error loading summaries: {e}")

    def _toggle_selection(self):
        """Toggle between select all and deselect all."""
        all_checked = all(
            self.summary_list.item(i).checkState() == Qt.CheckState.Checked
            for i in range(self.summary_list.count())
        )

        new_state = Qt.CheckState.Unchecked if all_checked else Qt.CheckState.Checked

        for i in range(self.summary_list.count()):
            self.summary_list.item(i).setCheckState(new_state)

        self._update_toggle_button()

    def _update_toggle_button(self):
        """Update toggle button text."""
        if self.summary_list.count() == 0:
            return

        all_checked = all(
            self.summary_list.item(i).checkState() == Qt.CheckState.Checked
            for i in range(self.summary_list.count())
        )

        self.toggle_btn.setText("Deselect All" if all_checked else "Select All")

    def _reset_prompt(self):
        """Reset prompt to default."""
        self.prompt_edit.setPlainText(self.original_prompt)

    def _reset_all(self):
        """Reset all settings."""
        self._reset_prompt()
        self.mode_summaries.setChecked(True)

        for i in range(self.summary_list.count()):
            self.summary_list.item(i).setCheckState(Qt.CheckState.Checked)
        self._update_toggle_button()

    def _save(self):
        """Save settings."""
        # Gather selected summaries
        selected_summaries = []
        for i in range(self.summary_list.count()):
            item = self.summary_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                name = item.data(Qt.ItemDataRole.UserRole) or item.text()
                selected_summaries.append(name)

        if self.mode_summaries.isChecked() and not selected_summaries:
            QMessageBox.warning(self, "No Selection", "Please select at least one document.")
            return

        settings = {
            "from_summaries": self.mode_summaries.isChecked(),
            "selected_summaries": selected_summaries if self.mode_summaries.isChecked() else None,
        }
        self.settings_db.update_settings(self.script_name, settings)

        # Save prompt if modified
        current_prompt = self.prompt_edit.toPlainText()
        if current_prompt != self.original_prompt:
            try:
                with open(self.prompt_path, 'w', encoding='utf-8') as f:
                    f.write(current_prompt)
                log_event("Saved updated timeline extraction prompt")
            except Exception as e:
                QMessageBox.warning(self, "Warning", f"Could not save prompt: {e}")

        self.accept()

    def eventFilter(self, obj, event):
        """Handle key press events for delete functionality."""
        if obj == self.summary_list and event.type() == event.Type.KeyPress:
            if event.key() == Qt.Key.Key_Delete:
                self._delete_selected()
                return True
        return super().eventFilter(obj, event)

    def _show_context_menu(self, position):
        """Show context menu for the summary list."""
        menu = QMenu(self)

        delete_action = menu.addAction("Delete Selected")
        delete_action.triggered.connect(self._delete_selected)

        selected_items = self.summary_list.selectedItems()
        if not selected_items:
            delete_action.setEnabled(False)

        menu.exec(self.summary_list.mapToGlobal(position))

    def _delete_selected(self):
        """Delete selected items from the list and document registry."""
        selected_items = self.summary_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select items to delete (click to select, Ctrl+click for multiple).")
            return

        count = len(selected_items)
        reply = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Delete {count} item(s) from the document registry?\n\nThis will remove them from the list but NOT delete the actual summary files.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            from document_registry import DocumentRegistry
            registry = DocumentRegistry()

            deleted_count = 0
            for item in selected_items:
                name = item.data(Qt.ItemDataRole.UserRole) or item.text()

                if registry.remove_document(self.file_number, name):
                    deleted_count += 1
                    log_event(f"Removed '{name}' from document registry")

                row = self.summary_list.row(item)
                self.summary_list.takeItem(row)

            remaining = self.summary_list.count()
            self.info_label.setText(f"Found {remaining} documents for case {self.file_number} ({deleted_count} deleted)")
            self._update_toggle_button()

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error deleting items: {e}")
            log_event(f"Error deleting from registry: {e}", "error")


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
    """Split view preview pane for files with PDF navigation."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)  # Enable keyboard focus
        self.current_file = None
        self._doc_worker = None  # For async .doc text extraction

        # PDF navigation state
        self._pdf_doc = None
        self._pdf_current_page = 0
        self._pdf_total_pages = 0
        self._pdf_zoom = 1.5  # Default zoom level

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

        # Page navigation bar (for PDFs)
        self.nav_bar = QWidget()
        nav_layout = QHBoxLayout(self.nav_bar)
        nav_layout.setContentsMargins(0, 2, 0, 2)
        nav_layout.setSpacing(4)

        self.prev_btn = QPushButton("◀")
        self.prev_btn.setFixedSize(28, 24)
        self.prev_btn.clicked.connect(self._prev_page)
        self.prev_btn.setEnabled(False)
        self.prev_btn.setToolTip("Previous page (Left/Up/PageUp)")
        nav_layout.addWidget(self.prev_btn)

        self.page_label = QLabel("Page 1 / 1")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_label.setMinimumWidth(80)
        self.page_label.setToolTip("Home=First, End=Last")
        nav_layout.addWidget(self.page_label)

        self.next_btn = QPushButton("▶")
        self.next_btn.setFixedSize(28, 24)
        self.next_btn.clicked.connect(self._next_page)
        self.next_btn.setEnabled(False)
        self.next_btn.setToolTip("Next page (Right/Down/PageDown/Space)")
        nav_layout.addWidget(self.next_btn)

        nav_layout.addStretch()
        self.nav_bar.hide()  # Hidden until PDF is loaded
        layout.addWidget(self.nav_bar)

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
        self.image_scroll = QScrollArea()
        self.image_scroll.setWidget(self.image_label)
        self.image_scroll.setWidgetResizable(True)
        self.preview_stack.addWidget(self.image_scroll)

        # No preview placeholder
        self.no_preview_label = QLabel("No preview available.\nDouble-click to open file.")
        self.no_preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.no_preview_label.setStyleSheet("color: gray;")
        self.preview_stack.addWidget(self.no_preview_label)

        # HTML preview for Word documents (using QWebEngineView)
        self.html_preview = None  # Lazy-loaded to avoid startup overhead

        layout.addWidget(self.preview_stack)

        self.preview_stack.setCurrentWidget(self.no_preview_label)

    def show_file(self, file_path):
        """Display preview for the given file."""
        # Close any previously open PDF when switching files
        self._close_pdf()
        self.nav_bar.hide()

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
        elif ext == '.docx':
            self._show_docx_preview(file_path)
        elif ext == '.doc':
            self._show_doc_preview(file_path)
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

    def _show_image_preview(self, file_path, pixmap=None, scale_to_fit=True):
        """Show image preview. Can accept a pre-created pixmap or load from file.

        Args:
            file_path: Path to the file (used if pixmap is None)
            pixmap: Pre-rendered QPixmap (optional)
            scale_to_fit: If True, scale to fit pane. If False, show at actual size (for zoom).
        """
        from PySide6.QtGui import QPixmap
        if pixmap is None:
            pixmap = QPixmap(file_path)
        if not pixmap.isNull():
            if scale_to_fit:
                # Scale to fit preview pane width while maintaining aspect ratio
                available_width = max(400, self.width() - 20)
                available_height = max(500, self.height() - 80)
                scaled = pixmap.scaled(
                    available_width, available_height,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.image_label.setPixmap(scaled)
            else:
                # Show at actual size (respects zoom level, scrollable)
                self.image_label.setPixmap(pixmap)
            self.preview_stack.setCurrentIndex(1)
        else:
            self.preview_stack.setCurrentWidget(self.no_preview_label)

    def _show_pdf_info(self, file_path):
        """Show PDF preview - renders pages as images using PyMuPDF with navigation."""
        # Close any previously open PDF
        self._close_pdf()

        try:
            import fitz  # PyMuPDF

            # Open PDF and store reference for navigation
            self._pdf_doc = fitz.open(file_path)
            self._pdf_total_pages = self._pdf_doc.page_count
            self._pdf_current_page = 0

            if self._pdf_total_pages > 0:
                # Update info label with page count
                self.info_label.setText(
                    f"{self.info_label.text()} | Pages: {self._pdf_total_pages}"
                )

                # Show navigation bar
                self._update_nav_bar()
                self.nav_bar.show()

                # Render first page
                self._render_pdf_page(0)
                return

            self._close_pdf()
        except ImportError:
            pass  # Fall back to text info if PyMuPDF not available
        except Exception as e:
            self._close_pdf()
            pass  # Fall back on any error

        # Fallback: show text info with OCR if available
        self.nav_bar.hide()
        self._show_pdf_text_fallback(file_path)

    def _render_pdf_page(self, page_num):
        """Render a specific PDF page."""
        if not self._pdf_doc or page_num < 0 or page_num >= self._pdf_total_pages:
            return

        try:
            import fitz
            from PySide6.QtGui import QPixmap, QImage

            page = self._pdf_doc[page_num]
            # Render at current zoom level
            mat = fitz.Matrix(self._pdf_zoom, self._pdf_zoom)
            pix = page.get_pixmap(matrix=mat)

            # Convert to QPixmap
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(img)

            self._pdf_current_page = page_num
            self._update_nav_bar()

            # Show in image preview without scaling (zoom is already applied in rendering)
            self._show_image_preview(self.current_file, pixmap, scale_to_fit=False)
        except Exception as e:
            self.no_preview_label.setText(f"Error rendering page {page_num + 1}:\n{str(e)}")
            self.preview_stack.setCurrentWidget(self.no_preview_label)

    def _update_nav_bar(self):
        """Update navigation bar state."""
        zoom_pct = int(self._pdf_zoom / 1.5 * 100)  # Show as percentage relative to default
        self.page_label.setText(f"Page {self._pdf_current_page + 1} / {self._pdf_total_pages} ({zoom_pct}%)")
        self.prev_btn.setEnabled(self._pdf_current_page > 0)
        self.next_btn.setEnabled(self._pdf_current_page < self._pdf_total_pages - 1)

    def _prev_page(self):
        """Navigate to previous page."""
        if self._pdf_current_page > 0:
            self._render_pdf_page(self._pdf_current_page - 1)

    def _next_page(self):
        """Navigate to next page."""
        if self._pdf_current_page < self._pdf_total_pages - 1:
            self._render_pdf_page(self._pdf_current_page + 1)

    def _close_pdf(self):
        """Close currently open PDF document."""
        if self._pdf_doc:
            try:
                self._pdf_doc.close()
            except:
                pass
            self._pdf_doc = None
        self._pdf_current_page = 0
        self._pdf_total_pages = 0
        self._pdf_zoom = 1.5  # Reset zoom to default

    def keyPressEvent(self, event):
        """Handle keyboard navigation for PDF pages."""
        from PySide6.QtCore import Qt
        if self._pdf_doc:
            if event.key() in (Qt.Key.Key_Right, Qt.Key.Key_Down, Qt.Key.Key_PageDown, Qt.Key.Key_Space):
                self._next_page()
                event.accept()
                return
            elif event.key() in (Qt.Key.Key_Left, Qt.Key.Key_Up, Qt.Key.Key_PageUp, Qt.Key.Key_Backspace):
                self._prev_page()
                event.accept()
                return
            elif event.key() == Qt.Key.Key_Home:
                self._render_pdf_page(0)
                event.accept()
                return
            elif event.key() == Qt.Key.Key_End:
                self._render_pdf_page(self._pdf_total_pages - 1)
                event.accept()
                return
            # Zoom with +/- keys
            elif event.key() in (Qt.Key.Key_Plus, Qt.Key.Key_Equal):
                self._zoom_in()
                event.accept()
                return
            elif event.key() == Qt.Key.Key_Minus:
                self._zoom_out()
                event.accept()
                return
            elif event.key() == Qt.Key.Key_0:
                self._zoom_reset()
                event.accept()
                return
        super().keyPressEvent(event)

    def wheelEvent(self, event):
        """Handle mouse wheel for page navigation and zoom."""
        from PySide6.QtCore import Qt
        if self._pdf_doc:
            # Ctrl+Wheel = Zoom
            if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
                delta = event.angleDelta().y()
                # Get cursor position relative to scroll area for zoom-to-cursor
                cursor_pos = event.position()
                if delta > 0:
                    self._zoom_in(cursor_pos)
                elif delta < 0:
                    self._zoom_out(cursor_pos)
                event.accept()
                return
            else:
                # Regular scroll = page navigation
                delta = event.angleDelta().y()
                if delta < 0:  # Scroll down = next page
                    self._next_page()
                elif delta > 0:  # Scroll up = previous page
                    self._prev_page()
                event.accept()
                return
        super().wheelEvent(event)

    def _zoom_in(self, cursor_pos=None):
        """Zoom in on PDF with zoom-to-cursor."""
        if self._pdf_doc and self._pdf_zoom < 4.0:
            old_zoom = self._pdf_zoom
            self._pdf_zoom = min(4.0, self._pdf_zoom + 0.25)
            self._render_pdf_page_with_cursor_focus(old_zoom, cursor_pos)

    def _zoom_out(self, cursor_pos=None):
        """Zoom out on PDF with zoom-to-cursor."""
        if self._pdf_doc and self._pdf_zoom > 0.5:
            old_zoom = self._pdf_zoom
            self._pdf_zoom = max(0.5, self._pdf_zoom - 0.25)
            self._render_pdf_page_with_cursor_focus(old_zoom, cursor_pos)

    def _zoom_reset(self):
        """Reset zoom to default."""
        if self._pdf_doc:
            self._pdf_zoom = 1.5
            self._render_pdf_page(self._pdf_current_page)

    def _render_pdf_page_with_cursor_focus(self, old_zoom, cursor_pos=None):
        """Render PDF page and adjust scroll to keep cursor position fixed."""
        if cursor_pos is None:
            # No cursor position, just render normally
            self._render_pdf_page(self._pdf_current_page)
            return

        # Get current scroll position and cursor position relative to content
        h_scroll = self.image_scroll.horizontalScrollBar()
        v_scroll = self.image_scroll.verticalScrollBar()

        # Map cursor position to scroll area viewport
        scroll_viewport_pos = self.image_scroll.mapFrom(self, cursor_pos.toPoint())

        # Calculate the document position under the cursor before zoom
        # doc_x = (scroll_position + cursor_in_viewport) / old_zoom
        doc_x = (h_scroll.value() + scroll_viewport_pos.x()) / old_zoom
        doc_y = (v_scroll.value() + scroll_viewport_pos.y()) / old_zoom

        # Render the page at new zoom level
        self._render_pdf_page(self._pdf_current_page)

        # Calculate new scroll position to keep the same document point under cursor
        # new_scroll = doc_pos * new_zoom - cursor_in_viewport
        new_h = int(doc_x * self._pdf_zoom - scroll_viewport_pos.x())
        new_v = int(doc_y * self._pdf_zoom - scroll_viewport_pos.y())

        # Apply new scroll positions (clamped to valid range)
        h_scroll.setValue(max(0, min(new_h, h_scroll.maximum())))
        v_scroll.setValue(max(0, min(new_v, v_scroll.maximum())))

    def _show_pdf_text_fallback(self, file_path):
        """Fallback PDF preview showing OCR text if available."""
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

    def _get_html_preview(self):
        """Lazy-load QWebEngineView for HTML preview."""
        if self.html_preview is None:
            try:
                from PySide6.QtWebEngineWidgets import QWebEngineView
                self.html_preview = QWebEngineView()
                self.html_preview.setStyleSheet("background-color: white;")
                self.preview_stack.addWidget(self.html_preview)
            except ImportError:
                return None
        return self.html_preview

    def _show_docx_preview(self, file_path):
        """Show preview for .docx files using mammoth (fast HTML conversion)."""
        try:
            import mammoth

            # Convert docx to HTML - this is fast, pure Python
            with open(file_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html_content = result.value

            if not html_content.strip():
                self.no_preview_label.setText("Document appears to be empty.")
                self.preview_stack.setCurrentWidget(self.no_preview_label)
                return

            # Wrap in styled HTML for better display
            styled_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {{
                        font-family: Calibri, Arial, sans-serif;
                        font-size: 11pt;
                        line-height: 1.4;
                        padding: 20px;
                        max-width: 100%;
                        background-color: white;
                        color: black;
                    }}
                    table {{
                        border-collapse: collapse;
                        margin: 10px 0;
                        width: 100%;
                    }}
                    td, th {{
                        border: 1px solid #ccc;
                        padding: 6px 10px;
                        text-align: left;
                    }}
                    th {{
                        background-color: #f0f0f0;
                        font-weight: bold;
                    }}
                    img {{
                        max-width: 100%;
                        height: auto;
                    }}
                    p {{
                        margin: 0 0 10px 0;
                    }}
                    h1, h2, h3, h4, h5, h6 {{
                        margin: 15px 0 10px 0;
                    }}
                    ul, ol {{
                        margin: 10px 0;
                        padding-left: 30px;
                    }}
                </style>
            </head>
            <body>
                {html_content}
            </body>
            </html>
            """

            # Try to use QWebEngineView for rich rendering
            web_view = self._get_html_preview()
            if web_view:
                web_view.setHtml(styled_html)
                self.preview_stack.setCurrentWidget(web_view)
            else:
                # Fallback to QTextEdit (basic HTML support)
                self.text_preview.setHtml(styled_html)
                self.preview_stack.setCurrentWidget(self.text_preview)

        except ImportError:
            self.no_preview_label.setText("mammoth not installed.\nInstall with: pip install mammoth")
            self.preview_stack.setCurrentWidget(self.no_preview_label)
        except Exception as e:
            self.no_preview_label.setText(f"Error reading .docx file:\n{str(e)}")
            self.preview_stack.setCurrentWidget(self.no_preview_label)

    def _show_doc_preview(self, file_path):
        """Show preview for legacy .doc files - text extraction only."""
        from PySide6.QtCore import QThread, Signal

        # Show loading message
        self.text_preview.setPlainText("Loading .doc preview...")
        self.preview_stack.setCurrentWidget(self.text_preview)

        class DocExtractWorker(QThread):
            finished = Signal(str, str)  # content, error

            def __init__(self, path):
                super().__init__()
                self.path = path

            def run(self):
                try:
                    import win32com.client
                    import pythoncom

                    pythoncom.CoInitialize()
                    try:
                        word = win32com.client.Dispatch("Word.Application")
                        word.Visible = False
                        word.DisplayAlerts = False

                        doc = word.Documents.Open(self.path, ReadOnly=True)
                        content = doc.Content.Text
                        doc.Close(False)
                        word.Quit()

                        if len(content) > 50000:
                            content = content[:50000] + '\n\n... (content truncated)'
                        self.finished.emit(content, "")
                    finally:
                        pythoncom.CoUninitialize()

                except ImportError:
                    self.finished.emit("", "pywin32 not installed.\nCannot preview .doc files.\n\nConvert to .docx for formatted preview.")
                except Exception as e:
                    self.finished.emit("", f"Error reading .doc file:\n{str(e)}\n\nTry opening with Microsoft Word.")

        def on_doc_extracted(content, error):
            if error:
                self.no_preview_label.setText(error)
                self.preview_stack.setCurrentWidget(self.no_preview_label)
            elif content.strip():
                self.text_preview.setPlainText(content)
                self.preview_stack.setCurrentWidget(self.text_preview)
            else:
                self.no_preview_label.setText("Document appears to be empty.")
                self.preview_stack.setCurrentWidget(self.no_preview_label)
            self._doc_worker = None

        self._doc_worker = DocExtractWorker(file_path)
        self._doc_worker.finished.connect(on_doc_extracted)
        self._doc_worker.start()

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
        self._close_pdf()
        self.nav_bar.hide()
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
        ext = os.path.splitext(path)[1].lower()

        # Handle .docx files with mammoth
        if ext == '.docx':
            self._show_docx_preview(path)
            return

        # Skip binary files that can't be previewed as text
        binary_extensions = {'.pdf', '.doc', '.xlsx', '.xls', '.pptx', '.ppt',
                           '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.zip', '.rar'}
        if ext in binary_extensions:
            self.output_preview.setPlainText(
                f"Binary file cannot be previewed as text.\n\n"
                f"File: {os.path.basename(path)}\n"
                f"Type: {ext}\n\n"
                f"Double-click to open with default application."
            )
            return

        try:
            content = None

            # Try different encodings in order of likelihood
            encodings = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'cp1252', 'latin-1']

            for encoding in encodings:
                try:
                    with open(path, 'r', encoding=encoding) as f:
                        content = f.read()

                    # Check if content looks like valid text (not garbled)
                    # If more than 10% are control characters (except newlines/tabs), try next encoding
                    if content:
                        control_chars = sum(1 for c in content[:1000] if ord(c) < 32 and c not in '\n\r\t')
                        sample_len = min(len(content), 1000)
                        if sample_len > 0 and control_chars / sample_len > 0.1:
                            content = None
                            continue
                        break  # Content looks valid
                except (UnicodeDecodeError, UnicodeError):
                    continue

            if content:
                self.output_preview.setPlainText(content)
            else:
                # Fallback: read as binary and try to decode
                with open(path, 'rb') as f:
                    raw = f.read()
                # Try to detect BOM
                if raw.startswith(b'\xff\xfe'):
                    content = raw.decode('utf-16-le', errors='replace')
                elif raw.startswith(b'\xfe\xff'):
                    content = raw.decode('utf-16-be', errors='replace')
                elif raw.startswith(b'\xef\xbb\xbf'):
                    content = raw[3:].decode('utf-8', errors='replace')
                else:
                    content = raw.decode('utf-8', errors='replace')
                self.output_preview.setPlainText(content)

        except Exception as e:
            self.output_preview.setPlainText(f"Error reading file: {e}")

    def _show_docx_preview(self, file_path):
        """Show preview for .docx files using mammoth."""
        try:
            import mammoth

            with open(file_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html_content = result.value

            if not html_content.strip():
                self.output_preview.setPlainText("Document appears to be empty.")
                return

            # Style the HTML for better display
            styled_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {{
                        font-family: Calibri, Arial, sans-serif;
                        font-size: 11pt;
                        line-height: 1.4;
                        padding: 20px;
                        background-color: white;
                        color: black;
                    }}
                    table {{
                        border-collapse: collapse;
                        margin: 10px 0;
                        width: 100%;
                    }}
                    td, th {{
                        border: 1px solid #ccc;
                        padding: 6px 10px;
                        text-align: left;
                    }}
                    th {{
                        background-color: #f0f0f0;
                        font-weight: bold;
                    }}
                    p {{
                        margin: 0 0 10px 0;
                    }}
                    h1, h2, h3, h4, h5, h6 {{
                        margin: 15px 0 10px 0;
                    }}
                    ul, ol {{
                        margin: 10px 0;
                        padding-left: 30px;
                    }}
                </style>
            </head>
            <body>
                {html_content}
            </body>
            </html>
            """

            # QTextEdit supports basic HTML rendering
            self.output_preview.setHtml(styled_html)

        except ImportError:
            self.output_preview.setPlainText(
                "mammoth library not installed.\n\n"
                "Install with: pip install mammoth\n\n"
                "Double-click to open with default application."
            )
        except Exception as e:
            self.output_preview.setPlainText(f"Error reading .docx file: {e}")

    def _open_selected_output(self):
        items = self.output_list.selectedItems()
        if items:
            path = items[0].data(Qt.ItemDataRole.UserRole)
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def _read_text_file(self, path):
        """Read a text file with automatic encoding detection."""
        ext = os.path.splitext(path)[1].lower()

        # Handle .docx files using mammoth
        if ext == '.docx':
            try:
                import mammoth
                with open(path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    return result.value
            except:
                return None

        # Handle PDF files using pypdf
        if ext == '.pdf':
            try:
                from pypdf import PdfReader
                reader = PdfReader(path)
                text_parts = []
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)
                if text_parts:
                    return "\n\n".join(text_parts)
                return "(No text could be extracted from this PDF)"
            except Exception as e:
                return f"(Error extracting PDF text: {e})"

        binary_extensions = {'.doc', '.xlsx', '.xls', '.pptx', '.ppt',
                           '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.zip', '.rar'}
        if ext in binary_extensions:
            return None

        encodings = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'cp1252', 'latin-1']

        for encoding in encodings:
            try:
                with open(path, 'r', encoding=encoding) as f:
                    content = f.read()
                # Validate content looks like text
                if content:
                    control_chars = sum(1 for c in content[:1000] if ord(c) < 32 and c not in '\n\r\t')
                    sample_len = min(len(content), 1000)
                    if sample_len > 0 and control_chars / sample_len > 0.1:
                        continue
                    return content
            except (UnicodeDecodeError, UnicodeError):
                continue

        # Fallback with BOM detection
        try:
            with open(path, 'rb') as f:
                raw = f.read()
            if raw.startswith(b'\xff\xfe'):
                return raw.decode('utf-16-le', errors='replace')
            elif raw.startswith(b'\xfe\xff'):
                return raw.decode('utf-16-be', errors='replace')
            elif raw.startswith(b'\xef\xbb\xbf'):
                return raw[3:].decode('utf-8', errors='replace')
            else:
                return raw.decode('utf-8', errors='replace')
        except:
            return None

    def _do_compare(self):
        left_path = self.compare_left.currentData()
        right_path = self.compare_right.currentData()

        if not left_path or not right_path:
            return

        try:
            left_content = self._read_text_file(left_path)
            right_content = self._read_text_file(right_path)

            if left_content is None:
                left_content = f"Cannot preview binary file: {os.path.basename(left_path)}"
            if right_content is None:
                right_content = f"Cannot preview binary file: {os.path.basename(right_path)}"

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
                content = self._read_text_file(output["path"])
                if content is None:
                    continue  # Skip binary files

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

                # Delete
                delete_action = QAction("Delete", self)
                delete_action.triggered.connect(lambda: self._delete_item(item))
                menu.addAction(delete_action)

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

    def keyPressEvent(self, event):
        """Handle keyboard shortcuts including Delete key."""
        if event.key() == Qt.Key.Key_Delete:
            items = self.selectedItems()
            if items:
                self._delete_item(items[0])
                event.accept()
                return
        super().keyPressEvent(event)

    def _delete_item(self, item):
        """Delete a file or folder after confirmation."""
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if not path or not os.path.exists(path):
            return

        item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
        item_name = os.path.basename(path)

        if item_type == "dir":
            msg = f"Are you sure you want to delete the folder '{item_name}' and all its contents?\n\nThis action cannot be undone."
        else:
            msg = f"Are you sure you want to delete '{item_name}'?\n\nThis action cannot be undone."

        confirm = QMessageBox.question(
            self,
            "Confirm Delete",
            msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path)
                else:
                    os.remove(path)
                log_event(f"Deleted: {path}")
                # Emit signal to refresh the tree
                self.item_moved.emit("", "")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not delete: {e}")
                log_event(f"Error deleting {path}: {e}", "error")

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
        # Tags column has been removed from the UI
        pass

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
