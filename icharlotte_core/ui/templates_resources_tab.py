"""
Templates / Resources Tab for iCharlotte

Provides a split-pane interface for managing:
- Templates: Editable document templates with variable placeholders
- Resources: Reference materials (PDFs, manuals, books) with tagging
"""

import os
import re
import json
import html
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QFrame,
    QLabel, QPushButton, QComboBox, QLineEdit, QTreeWidget,
    QTreeWidgetItem, QStackedWidget, QMenu, QMessageBox,
    QDialog, QDialogButtonBox, QListWidget, QListWidgetItem,
    QAbstractItemView, QApplication, QFileIconProvider, QInputDialog
)
from PySide6.QtCore import Qt, QUrl, QFileInfo, Signal, QObject, Slot, QTimer, QEvent
from PySide6.QtGui import QAction, QKeyEvent
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineSettings
from PySide6.QtWebChannel import QWebChannel

from ..config import TEMPLATES_DIR, RESOURCES_DIR, TEMPLATE_EXTENSIONS, RESOURCE_EXTENSIONS, GEMINI_DATA_DIR
from ..templates_db import TemplatesDatabase
from ..utils import log_event
from ..bridge import LocalFileSchemeHandler


class EditorBridge(QObject):
    """Bridge for QWebChannel communication between JS editor and Python."""

    content_changed = Signal()

    @Slot()
    def notifyContentChanged(self):
        """Called from JavaScript when editor content changes."""
        self.content_changed.emit()


class TemplateTreeWidget(QTreeWidget):
    """Custom tree widget with context menu and tagging support."""

    item_selected = Signal(str, str)  # path, type ('template' or 'resource')
    item_deleted = Signal(str)  # path of deleted item
    request_refresh = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setDragEnabled(False)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_context_menu)
        self.itemClicked.connect(self.on_item_clicked)
        self.itemDoubleClicked.connect(self.on_item_double_clicked)
        self.db = TemplatesDatabase()
        self.mode = 'templates'  # 'templates' or 'resources'
        self.icon_provider = QFileIconProvider()

    def set_mode(self, mode):
        """Set tree mode to 'templates' or 'resources'."""
        self.mode = mode

    def on_item_clicked(self, item, column):
        path = item.data(0, Qt.ItemDataRole.UserRole)
        item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if path and item_type == 'file':
            self.item_selected.emit(path, self.mode)

    def on_item_double_clicked(self, item, column):
        path = item.data(0, Qt.ItemDataRole.UserRole)
        item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if path and item_type == 'file':
            try:
                os.startfile(path)
            except Exception as e:
                log_event(f"Error opening file: {e}", "error")

    def open_context_menu(self, position):
        item = self.itemAt(position)
        if not item:
            return

        path = item.data(0, Qt.ItemDataRole.UserRole)
        item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)

        if not path:
            return

        menu = QMenu()

        if item_type == 'file':
            open_action = QAction("Open", self)
            open_action.triggered.connect(lambda: os.startfile(path))
            menu.addAction(open_action)

            menu.addSeparator()

            tags_action = QAction("Edit Tags...", self)
            tags_action.triggered.connect(lambda: self.edit_tags(path))
            menu.addAction(tags_action)

            if self.mode == 'templates':
                category_action = QAction("Set Category...", self)
                category_action.triggered.connect(lambda: self.edit_category(path))
                menu.addAction(category_action)

            menu.addSeparator()

        reveal_action = QAction("Reveal in Explorer", self)
        reveal_action.triggered.connect(lambda: self.reveal_in_explorer(path))
        menu.addAction(reveal_action)

        copy_path_action = QAction("Copy Path", self)
        copy_path_action.triggered.connect(lambda: QApplication.clipboard().setText(path))
        menu.addAction(copy_path_action)

        if item_type == 'file':
            menu.addSeparator()
            delete_action = QAction("Delete", self)
            delete_action.triggered.connect(lambda: self.delete_item(path))
            menu.addAction(delete_action)

        menu.exec(self.viewport().mapToGlobal(position))

    def reveal_in_explorer(self, path):
        import subprocess
        try:
            subprocess.run(['explorer', '/select,', path])
        except Exception as e:
            log_event(f"Error revealing file: {e}", "error")

    def edit_tags(self, path):
        """Show dialog to edit tags for a file."""
        dialog = TagEditDialog(self, path, self.mode, self.db)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.request_refresh.emit()

    def edit_category(self, path):
        """Show dialog to set category for a template."""
        relative_path = os.path.relpath(path, TEMPLATES_DIR)
        template = self.db.get_template_by_path(relative_path)

        current_category = template['category'] if template else ''
        categories = [''] + self.db.get_all_categories()

        category, ok = QInputDialog.getItem(
            self, "Set Category", "Category:",
            categories, categories.index(current_category) if current_category in categories else 0,
            editable=True
        )

        if ok and template:
            self.db.update_template_category(template['id'], category)
            self.request_refresh.emit()

    def keyPressEvent(self, event):
        """Handle key press events - Delete key deletes selected template."""
        if event.key() == Qt.Key.Key_Delete:
            item = self.currentItem()
            if item:
                path = item.data(0, Qt.ItemDataRole.UserRole)
                item_type = item.data(0, Qt.ItemDataRole.UserRole + 1)
                if path and item_type == 'file':
                    self.delete_item(path)
                    return
        super().keyPressEvent(event)

    def delete_item(self, path):
        """Delete a template or resource file after confirmation."""
        filename = os.path.basename(path)

        reply = QMessageBox.question(
            self, "Delete File",
            f"Are you sure you want to delete '{filename}'?\n\nThis cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                # Remove from database if it's a template
                if self.mode == 'templates':
                    relative_path = os.path.relpath(path, TEMPLATES_DIR)
                    template = self.db.get_template_by_path(relative_path)
                    if template:
                        self.db.delete_template(template['id'])

                # Delete the file
                os.remove(path)
                log_event(f"Deleted {self.mode[:-1]}: {filename}", "info")
                self.item_deleted.emit(path)
                self.request_refresh.emit()

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete file: {e}")
                log_event(f"Error deleting file: {e}", "error")


class TagEditDialog(QDialog):
    """Dialog for editing tags on a file."""

    def __init__(self, parent, path, mode, db):
        super().__init__(parent)
        self.path = path
        self.mode = mode
        self.db = db
        self.setWindowTitle("Edit Tags")
        self.setMinimumWidth(300)
        self.setup_ui()
        self.load_tags()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Current Tags:"))

        self.tag_list = QListWidget()
        self.tag_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        layout.addWidget(self.tag_list)

        # Add new tag
        add_layout = QHBoxLayout()
        self.new_tag_input = QLineEdit()
        self.new_tag_input.setPlaceholderText("New tag...")
        self.new_tag_input.returnPressed.connect(self.add_tag)
        add_layout.addWidget(self.new_tag_input)

        add_btn = QPushButton("Add")
        add_btn.clicked.connect(self.add_tag)
        add_layout.addWidget(add_btn)
        layout.addLayout(add_layout)

        # Remove selected
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self.remove_selected)
        layout.addWidget(remove_btn)

        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def load_tags(self):
        self.tag_list.clear()
        if self.mode == 'templates':
            relative_path = os.path.relpath(self.path, TEMPLATES_DIR)
            template = self.db.get_template_by_path(relative_path)
            if template:
                tags = self.db.get_template_tags(template['id'])
                for tag in tags:
                    self.tag_list.addItem(tag)
        else:
            tags = self.db.get_resource_tags(self.path)
            for tag in tags:
                self.tag_list.addItem(tag)

    def add_tag(self):
        tag = self.new_tag_input.text().strip()
        if not tag:
            return

        if self.mode == 'templates':
            relative_path = os.path.relpath(self.path, TEMPLATES_DIR)
            template = self.db.get_template_by_path(relative_path)
            if template:
                self.db.add_template_tag(template['id'], tag)
        else:
            self.db.add_resource_tag(self.path, tag)

        self.new_tag_input.clear()
        self.load_tags()

    def remove_selected(self):
        selected = self.tag_list.selectedItems()
        for item in selected:
            tag = item.text()
            if self.mode == 'templates':
                relative_path = os.path.relpath(self.path, TEMPLATES_DIR)
                template = self.db.get_template_by_path(relative_path)
                if template:
                    self.db.remove_template_tag(template['id'], tag)
            else:
                self.db.remove_resource_tag(self.path, tag)
        self.load_tags()


class PlaceholderMappingDialog(QDialog):
    """Dialog for mapping a custom placeholder to a case variable."""

    # Available case variables grouped by category
    VARIABLE_GROUPS = [
        ("Case Info", [
            ("Case Name", "CASE_NAME"),
            ("Case Number", "CASE_NUMBER"),
            ("File Number", "FILE_NUMBER"),
            ("Venue County", "VENUE_COUNTY"),
        ]),
        ("Parties", [
            ("Plaintiff", "PLAINTIFF"),
            ("Plaintiffs (all)", "PLAINTIFFS"),
            ("Defendant", "DEFENDANT"),
            ("Defendants (all)", "DEFENDANTS"),
            ("Client Name", "CLIENT_NAME"),
        ]),
        ("Dates", [
            ("Today's Date", "TODAY_DATE"),
            ("Today (Long)", "TODAY_DATE_LONG"),
            ("Trial Date", "TRIAL_DATE"),
            ("Incident Date", "INCIDENT_DATE"),
            ("Filing Date", "FILING_DATE"),
        ]),
        ("Insurance", [
            ("Claim Number", "CLAIM_NUMBER"),
            ("Adjuster Name", "ADJUSTER_NAME"),
            ("Adjuster Email", "ADJUSTER_EMAIL"),
        ]),
        ("Other", [
            ("Plaintiff Counsel", "PLAINTIFF_COUNSEL"),
        ]),
    ]

    def __init__(self, parent, placeholder_name, db, existing_mapping=None):
        super().__init__(parent)
        self.placeholder_name = placeholder_name.upper()
        self.db = db
        self.existing_mapping = existing_mapping
        self.selected_mapping = None
        self.setWindowTitle("Map Placeholder")
        self.setMinimumWidth(350)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Placeholder name display
        name_label = QLabel(f"Custom Placeholder: <b>{{{{{self.placeholder_name}}}}}</b>")
        name_label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(name_label)

        layout.addSpacing(10)

        # Mapping dropdown
        layout.addWidget(QLabel("Map to case variable:"))

        self.mapping_combo = QComboBox()
        self.mapping_combo.setMinimumWidth(280)

        # Add "None" option first
        self.mapping_combo.addItem("(None - keep as custom placeholder)", None)

        # Add grouped variables
        for group_name, variables in self.VARIABLE_GROUPS:
            # Add group header (disabled)
            self.mapping_combo.addItem(f"-- {group_name} --")
            idx = self.mapping_combo.count() - 1
            self.mapping_combo.model().item(idx).setEnabled(False)

            # Add variables in group
            for label, value in variables:
                self.mapping_combo.addItem(f"  {label} ({{{{{value}}}}})", value)

        # Select existing mapping if present
        if self.existing_mapping:
            for i in range(self.mapping_combo.count()):
                if self.mapping_combo.itemData(i) == self.existing_mapping:
                    self.mapping_combo.setCurrentIndex(i)
                    break

        layout.addWidget(self.mapping_combo)

        layout.addSpacing(10)

        # Info text
        info_label = QLabel(
            "When you 'Copy Filled', this placeholder will be replaced\n"
            "with the value of the selected case variable."
        )
        info_label.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(info_label)

        layout.addSpacing(10)

        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # Add delete button if editing existing mapping
        if self.existing_mapping:
            delete_btn = buttons.addButton("Remove Mapping", QDialogButtonBox.ButtonRole.DestructiveRole)
            delete_btn.clicked.connect(self.remove_mapping)

        layout.addWidget(buttons)

    def accept(self):
        self.selected_mapping = self.mapping_combo.currentData()
        super().accept()

    def remove_mapping(self):
        self.selected_mapping = None
        self.db.delete_placeholder_mapping(self.placeholder_name)
        super().accept()

    def get_mapping(self):
        """Return the selected mapping (None if no mapping selected)."""
        return self.selected_mapping


class TemplatesResourcesTab(QWidget):
    """Main tab widget for Templates and Resources management."""

    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.db = TemplatesDatabase()
        self.current_item_path = None
        self.current_item_type = None  # 'templates' or 'resources'
        self.current_mode = 'templates'
        self.editor_modified = False

        # Auto-save timer (debounced - fires 1.5 seconds after last change)
        self.auto_save_timer = QTimer(self)
        self.auto_save_timer.setSingleShot(True)
        self.auto_save_timer.setInterval(1500)  # 1.5 seconds
        self.auto_save_timer.timeout.connect(self._perform_auto_save)

        # Bridge for JS-Python communication
        self.editor_bridge = EditorBridge()
        self.editor_bridge.content_changed.connect(self._on_editor_content_changed)

        self.setup_ui()
        self.ensure_directories()
        self.refresh_tree()

    def ensure_directories(self):
        """Create Templates and Resources directories if they don't exist."""
        os.makedirs(TEMPLATES_DIR, exist_ok=True)
        os.makedirs(RESOURCES_DIR, exist_ok=True)

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # Top Toolbar
        toolbar = self.setup_toolbar()
        main_layout.addWidget(toolbar)

        # Main Splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left Panel - Tree View
        left_panel = self.setup_left_panel()
        splitter.addWidget(left_panel)

        # Right Panel - Preview/Editor
        right_panel = self.setup_right_panel()
        splitter.addWidget(right_panel)

        splitter.setSizes([300, 700])
        main_layout.addWidget(splitter)

    def setup_toolbar(self):
        """Create the top toolbar - compact design."""
        toolbar = QFrame()
        toolbar.setObjectName("mainToolbar")
        toolbar.setMaximumHeight(36)  # Constrain height to single row
        toolbar.setStyleSheet("""
            #mainToolbar {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fafafa, stop:1 #f0f0f0);
                border: 1px solid #d0d0d0;
                border-radius: 3px;
            }
            #mainToolbar QLabel { color: #555; font-size: 11px; }
            #mainToolbar QComboBox {
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 2px 4px;
                background: white;
                min-height: 20px;
            }
            #mainToolbar QLineEdit {
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 2px 6px;
                background: white;
                min-height: 20px;
            }
            #mainToolbar QPushButton {
                border: 1px solid #bbb;
                border-radius: 3px;
                padding: 3px 8px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fff, stop:1 #e8e8e8);
                min-height: 20px;
            }
            #mainToolbar QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fff, stop:1 #ddd);
                border-color: #999;
            }
        """)
        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(6, 3, 6, 3)
        layout.setSpacing(4)

        # Mode selector
        self.mode_selector = QComboBox()
        self.mode_selector.addItems(["Templates", "Resources"])
        self.mode_selector.setFixedWidth(90)
        self.mode_selector.currentTextChanged.connect(self.on_mode_changed)
        layout.addWidget(self.mode_selector)

        # Separator
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.VLine)
        sep1.setStyleSheet("color: #ccc;")
        layout.addWidget(sep1)

        # Search
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setFixedWidth(140)
        self.search_input.textChanged.connect(self.filter_tree)
        layout.addWidget(self.search_input)

        # Category filter (templates only)
        self.category_filter = QComboBox()
        self.category_filter.setFixedWidth(100)
        self.category_filter.addItem("All Categories")
        self.category_filter.currentTextChanged.connect(self.on_category_changed)
        layout.addWidget(self.category_filter)

        # Tag filter
        self.tag_filter_btn = QPushButton("Tags")
        self.tag_filter_btn.setFixedWidth(50)
        self.tag_filter_btn.clicked.connect(self.show_tag_filter)
        layout.addWidget(self.tag_filter_btn)
        self.active_tags = []

        layout.addStretch()

        # Refresh
        refresh_btn = QPushButton("↻")
        refresh_btn.setToolTip("Refresh")
        refresh_btn.setFixedWidth(28)
        refresh_btn.clicked.connect(self.refresh_tree)
        layout.addWidget(refresh_btn)

        return toolbar

    def setup_left_panel(self):
        """Create the left panel with tree view."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # Tree view
        self.tree = TemplateTreeWidget()
        self.tree.item_selected.connect(self.on_item_selected)
        self.tree.item_deleted.connect(self.on_item_deleted)
        self.tree.request_refresh.connect(self.refresh_tree)
        layout.addWidget(self.tree)

        # Action buttons
        actions = QFrame()
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(5, 5, 5, 5)

        self.new_btn = QPushButton("New Template")
        self.new_btn.clicked.connect(self.create_new_template)
        actions_layout.addWidget(self.new_btn)

        self.open_folder_btn = QPushButton("Open Folder")
        self.open_folder_btn.clicked.connect(self.open_current_folder)
        actions_layout.addWidget(self.open_folder_btn)

        layout.addWidget(actions)

        return panel

    def setup_right_panel(self):
        """Create the right panel with preview/editor."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        # Placeholder toolbar (for templates)
        self.placeholder_toolbar = self.setup_placeholder_toolbar()
        layout.addWidget(self.placeholder_toolbar)

        # Format toolbar (for templates)
        self.format_toolbar = self.setup_format_toolbar()
        layout.addWidget(self.format_toolbar)

        # Stacked widget for different views
        self.content_stack = QStackedWidget()

        # Page 0: Placeholder
        placeholder = QLabel("Select a template or resource to preview")
        placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        placeholder.setStyleSheet("color: #888; font-size: 14px;")
        self.content_stack.addWidget(placeholder)

        # Page 1: PDF Preview
        self.pdf_preview = QWebEngineView()
        # Configure PDF preview settings
        settings = self.pdf_preview.settings()
        settings.setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.PdfViewerEnabled, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        # Install scheme handler for local-resource:// URLs
        self.scheme_handler = LocalFileSchemeHandler(self)
        self.pdf_preview.page().profile().installUrlSchemeHandler(b"local-resource", self.scheme_handler)
        self.content_stack.addWidget(self.pdf_preview)

        # Page 2: Rich Text Editor
        self.editor = QWebEngineView()
        self.editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.editor.customContextMenuRequested.connect(self.show_editor_context_menu)

        # Install event filter to catch Tab key before Qt handles it
        self.editor.installEventFilter(self)
        # Also install on the focus proxy (the actual widget that receives key events)
        QTimer.singleShot(100, self._install_editor_event_filter)

        # Set up QWebChannel for JS-Python communication
        self.web_channel = QWebChannel(self.editor.page())
        self.web_channel.registerObject("bridge", self.editor_bridge)
        self.editor.page().setWebChannel(self.web_channel)

        self.editor.setHtml(self.get_editor_html(""))
        self.content_stack.addWidget(self.editor)

        layout.addWidget(self.content_stack, 1)

        # Bottom actions
        bottom = QFrame()
        bottom.setStyleSheet("QFrame { background: #f5f5f5; border-top: 1px solid #ddd; }")
        bottom_layout = QHBoxLayout(bottom)
        bottom_layout.setContentsMargins(10, 8, 10, 8)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666;")
        bottom_layout.addWidget(self.status_label)

        bottom_layout.addStretch()

        self.copy_raw_btn = QPushButton("Copy Raw")
        self.copy_raw_btn.setToolTip("Copy template with placeholders intact")
        self.copy_raw_btn.clicked.connect(self.copy_raw)
        self.copy_raw_btn.setEnabled(False)
        bottom_layout.addWidget(self.copy_raw_btn)

        self.copy_filled_btn = QPushButton("Copy Filled")
        self.copy_filled_btn.setToolTip("Copy template with case data filled in")
        self.copy_filled_btn.clicked.connect(self.copy_filled)
        self.copy_filled_btn.setEnabled(False)
        self.copy_filled_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        bottom_layout.addWidget(self.copy_filled_btn)

        layout.addWidget(bottom)

        return panel

    def setup_placeholder_toolbar(self):
        """Create placeholder insertion toolbar for templates."""
        toolbar = QFrame()
        toolbar.setObjectName("placeholderToolbar")
        toolbar.setStyleSheet("""
            #placeholderToolbar {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #e8f4e8, stop:1 #d4e8d4);
                border: 1px solid #a8c8a8;
                border-radius: 3px;
            }
            #placeholderToolbar QLabel { color: #2d5a2d; font-size: 11px; font-weight: bold; }
            #placeholderToolbar QComboBox {
                border: 1px solid #8ab88a;
                border-radius: 3px;
                padding: 2px 4px;
                background: white;
                min-height: 20px;
            }
            #placeholderToolbar QPushButton {
                border: 1px solid #7aa87a;
                border-radius: 3px;
                padding: 3px 8px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fff, stop:1 #e0f0e0);
                color: #2d5a2d;
                min-height: 20px;
                font-size: 11px;
            }
            #placeholderToolbar QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fff, stop:1 #c8e0c8);
                border-color: #5a8a5a;
            }
            #placeholderToolbar QPushButton#insertBtn {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5cb85c, stop:1 #449d44);
                color: white;
                border-color: #398439;
                font-weight: bold;
            }
            #placeholderToolbar QPushButton#insertBtn:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6cc86c, stop:1 #54ad54);
            }
        """)
        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(6, 3, 6, 3)
        layout.setSpacing(6)

        # Label
        layout.addWidget(QLabel("Insert Placeholder:"))

        # Placeholder dropdown with common placeholders
        self.placeholder_combo = QComboBox()
        self.placeholder_combo.setMinimumWidth(200)
        self.placeholder_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.populate_placeholder_combo()
        layout.addWidget(self.placeholder_combo)

        # Insert button
        insert_btn = QPushButton("Insert")
        insert_btn.setObjectName("insertBtn")
        insert_btn.setFixedWidth(55)
        insert_btn.clicked.connect(self.insert_placeholder)
        layout.addWidget(insert_btn)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setStyleSheet("color: #a8c8a8;")
        layout.addWidget(sep)

        # Create custom placeholder
        layout.addWidget(QLabel("Custom:"))
        self.custom_placeholder_input = QLineEdit()
        self.custom_placeholder_input.setPlaceholderText("NEW_FIELD")
        self.custom_placeholder_input.setFixedWidth(100)
        self.custom_placeholder_input.returnPressed.connect(self.insert_custom_placeholder)
        layout.addWidget(self.custom_placeholder_input)

        create_btn = QPushButton("+ Create")
        create_btn.setFixedWidth(60)
        create_btn.clicked.connect(self.insert_custom_placeholder)
        layout.addWidget(create_btn)

        # Separator
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.VLine)
        sep2.setStyleSheet("color: #a8c8a8;")
        layout.addWidget(sep2)

        # Manage mappings button
        mappings_btn = QPushButton("Mappings")
        mappings_btn.setFixedWidth(65)
        mappings_btn.setToolTip("Manage custom placeholder mappings")
        mappings_btn.clicked.connect(self.show_manage_mappings_dialog)
        layout.addWidget(mappings_btn)

        layout.addStretch()

        # Help button
        help_btn = QPushButton("?")
        help_btn.setFixedWidth(24)
        help_btn.setToolTip("Placeholders use {{NAME}} format and auto-fill from case data")
        help_btn.clicked.connect(self.show_placeholder_help)
        layout.addWidget(help_btn)

        toolbar.hide()  # Hidden by default
        return toolbar

    def populate_placeholder_combo(self):
        """Populate the placeholder dropdown with available placeholders."""
        self.placeholder_combo.clear()

        # Group placeholders by category
        placeholders = [
            ("-- Case Info --", None),
            ("Case Name", "CASE_NAME"),
            ("Case Number", "CASE_NUMBER"),
            ("File Number", "FILE_NUMBER"),
            ("Venue County", "VENUE_COUNTY"),
            ("-- Parties --", None),
            ("Plaintiff", "PLAINTIFF"),
            ("Plaintiffs (all)", "PLAINTIFFS"),
            ("Defendant", "DEFENDANT"),
            ("Defendants (all)", "DEFENDANTS"),
            ("Client Name", "CLIENT_NAME"),
            ("-- Dates --", None),
            ("Today's Date", "TODAY_DATE"),
            ("Today (Long)", "TODAY_DATE_LONG"),
            ("Trial Date", "TRIAL_DATE"),
            ("Incident Date", "INCIDENT_DATE"),
            ("Filing Date", "FILING_DATE"),
            ("-- Insurance --", None),
            ("Claim Number", "CLAIM_NUMBER"),
            ("Adjuster Name", "ADJUSTER_NAME"),
            ("Adjuster Email", "ADJUSTER_EMAIL"),
            ("-- Other --", None),
            ("Plaintiff Counsel", "PLAINTIFF_COUNSEL"),
        ]

        first_selectable_idx = None
        for label, value in placeholders:
            if value is None:
                self.placeholder_combo.addItem(label)
                # Make header items non-selectable
                idx = self.placeholder_combo.count() - 1
                self.placeholder_combo.model().item(idx).setEnabled(False)
            else:
                self.placeholder_combo.addItem(f"{label} ({{{{{value}}}}})", value)
                if first_selectable_idx is None:
                    first_selectable_idx = self.placeholder_combo.count() - 1

        # Add custom placeholders from database
        custom_mappings = self.db.get_all_placeholder_mappings()
        if custom_mappings:
            # Add Custom header
            self.placeholder_combo.addItem("-- Custom --")
            idx = self.placeholder_combo.count() - 1
            self.placeholder_combo.model().item(idx).setEnabled(False)

            # Add each custom placeholder
            for custom_name, maps_to in sorted(custom_mappings.items()):
                display = f"{custom_name} \u2192 {maps_to} ({{{{{custom_name}}}}})"
                self.placeholder_combo.addItem(display, custom_name)

        # Select first real placeholder by default
        if first_selectable_idx is not None:
            self.placeholder_combo.setCurrentIndex(first_selectable_idx)

    def insert_placeholder(self):
        """Insert the selected placeholder at cursor position in editor."""
        if self.current_item_type != 'templates':
            return

        value = self.placeholder_combo.currentData()
        if not value:
            return

        placeholder_text = f"{{{{{value}}}}}"
        self.insert_text_at_cursor(placeholder_text)

    def insert_custom_placeholder(self):
        """Insert a custom placeholder at cursor position."""
        if self.current_item_type != 'templates':
            QMessageBox.information(self, "No Template", "Please select a template first.")
            return

        name = self.custom_placeholder_input.text().strip().upper()
        if not name:
            return

        # Validate: only alphanumeric and underscores
        if not re.match(r'^[A-Z][A-Z0-9_]*$', name):
            QMessageBox.warning(self, "Invalid Name",
                "Placeholder names must start with a letter and contain only letters, numbers, and underscores.")
            return

        # Check if mapping already exists
        existing_mapping = self.db.get_placeholder_mapping(name)

        # Show mapping dialog
        dialog = PlaceholderMappingDialog(self, name, self.db, existing_mapping)
        result = dialog.exec()

        if result == QDialog.DialogCode.Accepted:
            mapping = dialog.get_mapping()
            if mapping:
                # Save the mapping
                self.db.add_placeholder_mapping(name, mapping)
                # Refresh dropdown to show new custom placeholder
                self.populate_placeholder_combo()
                # Update JavaScript variables for Ctrl+C to work
                self.update_editor_variables()

            # Insert the placeholder
            placeholder_text = f"{{{{{name}}}}}"
            self.insert_text_at_cursor(placeholder_text)
            self.custom_placeholder_input.clear()

    def insert_text_at_cursor(self, text):
        """Insert text at the current cursor position in the editor."""
        # Escape for JavaScript string only (backslashes and quotes)
        # Note: Braces in the text don't need escaping - f-string substitution inserts them as-is
        escaped_text = text.replace('\\', '\\\\').replace("'", "\\'")
        js = f"""
        (function() {{
            var sel = window.getSelection();
            if (sel.rangeCount > 0) {{
                var range = sel.getRangeAt(0);
                range.deleteContents();
                var textNode = document.createTextNode('{escaped_text}');
                range.insertNode(textNode);
                range.setStartAfter(textNode);
                range.setEndAfter(textNode);
                sel.removeAllRanges();
                sel.addRange(range);
            }} else {{
                document.getElementById('editor').innerHTML += '{escaped_text}';
            }}
        }})();
        """
        self.editor.page().runJavaScript(js)
        # Trigger content change handler (for auto-save)
        self._on_editor_content_changed()

    def update_editor_variables(self):
        """Update the JavaScript templateVariables with current case variables."""
        if self.current_item_type != 'templates':
            return
        variables = self.get_case_variables()
        vars_js = json.dumps(variables)
        js = f"templateVariables = {vars_js};"
        self.editor.page().runJavaScript(js)

    def show_editor_context_menu(self, position):
        """Show context menu with placeholder options when right-clicking in editor."""
        # Only show placeholder menu when editing templates
        if self.current_item_type != 'templates':
            return

        menu = QMenu(self)

        # Add "Edit Placeholder Mapping" option at top
        edit_mapping_action = QAction("Edit Placeholder Mapping...", self)
        edit_mapping_action.triggered.connect(self.edit_placeholder_mapping_from_selection)
        menu.addAction(edit_mapping_action)

        # Add "Manage All Mappings" option
        manage_mappings_action = QAction("Manage All Mappings...", self)
        manage_mappings_action.triggered.connect(self.show_manage_mappings_dialog)
        menu.addAction(manage_mappings_action)

        menu.addSeparator()

        # Add placeholder submenu
        placeholder_menu = menu.addMenu("Insert Placeholder")

        # Define placeholder categories and items
        placeholder_groups = [
            ("Case Info", [
                ("Case Name", "CASE_NAME"),
                ("Case Number", "CASE_NUMBER"),
                ("File Number", "FILE_NUMBER"),
                ("Venue County", "VENUE_COUNTY"),
            ]),
            ("Parties", [
                ("Plaintiff", "PLAINTIFF"),
                ("Plaintiffs (all)", "PLAINTIFFS"),
                ("Defendant", "DEFENDANT"),
                ("Defendants (all)", "DEFENDANTS"),
                ("Client Name", "CLIENT_NAME"),
            ]),
            ("Dates", [
                ("Today's Date", "TODAY_DATE"),
                ("Today (Long)", "TODAY_DATE_LONG"),
                ("Trial Date", "TRIAL_DATE"),
                ("Incident Date", "INCIDENT_DATE"),
                ("Filing Date", "FILING_DATE"),
            ]),
            ("Insurance", [
                ("Claim Number", "CLAIM_NUMBER"),
                ("Adjuster Name", "ADJUSTER_NAME"),
                ("Adjuster Email", "ADJUSTER_EMAIL"),
            ]),
            ("Other", [
                ("Plaintiff Counsel", "PLAINTIFF_COUNSEL"),
            ]),
        ]

        for group_name, items in placeholder_groups:
            group_menu = placeholder_menu.addMenu(group_name)
            for label, value in items:
                action = QAction(f"{label}  {{{{{value}}}}}", self)
                action.setData(value)
                action.triggered.connect(lambda checked, v=value: self.insert_placeholder_from_menu(v))
                group_menu.addAction(action)

        # Add separator and custom placeholder option
        placeholder_menu.addSeparator()
        custom_action = QAction("Custom Placeholder...", self)
        custom_action.triggered.connect(self.prompt_custom_placeholder)
        placeholder_menu.addAction(custom_action)

        # Show menu at cursor position
        menu.exec(self.editor.mapToGlobal(position))

    def insert_placeholder_from_menu(self, value):
        """Insert a placeholder from the context menu."""
        placeholder_text = f"{{{{{value}}}}}"
        self.insert_text_at_cursor(placeholder_text)

    def prompt_custom_placeholder(self):
        """Prompt user for a custom placeholder name and insert it."""
        name, ok = QInputDialog.getText(
            self, "Custom Placeholder",
            "Enter placeholder name (letters, numbers, underscores):"
        )
        if ok and name:
            name = name.strip().upper()
            if not re.match(r'^[A-Z][A-Z0-9_]*$', name):
                QMessageBox.warning(self, "Invalid Name",
                    "Placeholder names must start with a letter and contain only letters, numbers, and underscores.")
                return

            # Show mapping dialog
            existing_mapping = self.db.get_placeholder_mapping(name)
            dialog = PlaceholderMappingDialog(self, name, self.db, existing_mapping)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                mapping = dialog.get_mapping()
                if mapping:
                    self.db.add_placeholder_mapping(name, mapping)
                    self.populate_placeholder_combo()  # Refresh dropdown
                    self.update_editor_variables()  # Update JS variables

                placeholder_text = f"{{{{{name}}}}}"
                self.insert_text_at_cursor(placeholder_text)

    def edit_placeholder_mapping_from_selection(self):
        """Edit mapping for a placeholder - prompts user for the placeholder name."""
        name, ok = QInputDialog.getText(
            self, "Edit Placeholder Mapping",
            "Enter placeholder name to edit mapping for:"
        )
        if ok and name:
            name = name.strip().upper()
            if not re.match(r'^[A-Z][A-Z0-9_]*$', name):
                QMessageBox.warning(self, "Invalid Name",
                    "Placeholder names must start with a letter and contain only letters, numbers, and underscores.")
                return

            existing_mapping = self.db.get_placeholder_mapping(name)
            dialog = PlaceholderMappingDialog(self, name, self.db, existing_mapping)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                mapping = dialog.get_mapping()
                if mapping:
                    self.db.add_placeholder_mapping(name, mapping)
                    self.status_label.setText(f"Mapped {{{{{name}}}}} to {{{{{mapping}}}}}")
                else:
                    self.status_label.setText(f"Removed mapping for {{{{{name}}}}}")
                # Refresh dropdown and JS variables to reflect changes
                self.populate_placeholder_combo()
                self.update_editor_variables()

    def show_manage_mappings_dialog(self):
        """Show dialog to view and manage all placeholder mappings."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Manage Placeholder Mappings")
        dialog.setMinimumSize(450, 350)
        # Refresh dropdown and JS variables when dialog closes
        dialog.finished.connect(lambda: (self.populate_placeholder_combo(), self.update_editor_variables()))
        layout = QVBoxLayout(dialog)

        layout.addWidget(QLabel("All custom placeholder mappings:"))

        # List widget showing all mappings
        mapping_list = QListWidget()
        mapping_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        def refresh_list():
            mapping_list.clear()
            mappings = self.db.get_all_placeholder_mappings()
            if not mappings:
                item = QListWidgetItem("(No custom mappings defined)")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                mapping_list.addItem(item)
            else:
                for custom_name, maps_to in sorted(mappings.items()):
                    item = QListWidgetItem(f"{{{{{custom_name}}}}}  \u2192  {{{{{maps_to}}}}}")
                    item.setData(Qt.ItemDataRole.UserRole, custom_name)
                    mapping_list.addItem(item)

        refresh_list()
        layout.addWidget(mapping_list)

        # Buttons
        btn_layout = QHBoxLayout()

        add_btn = QPushButton("Add New...")
        def on_add():
            name, ok = QInputDialog.getText(dialog, "New Mapping", "Placeholder name:")
            if ok and name:
                name = name.strip().upper()
                if not re.match(r'^[A-Z][A-Z0-9_]*$', name):
                    QMessageBox.warning(dialog, "Invalid Name",
                        "Names must start with a letter and contain only letters, numbers, underscores.")
                    return
                existing = self.db.get_placeholder_mapping(name)
                map_dialog = PlaceholderMappingDialog(dialog, name, self.db, existing)
                if map_dialog.exec() == QDialog.DialogCode.Accepted:
                    mapping = map_dialog.get_mapping()
                    if mapping:
                        self.db.add_placeholder_mapping(name, mapping)
                    refresh_list()

        add_btn.clicked.connect(on_add)
        btn_layout.addWidget(add_btn)

        edit_btn = QPushButton("Edit...")
        def on_edit():
            item = mapping_list.currentItem()
            if not item:
                return
            name = item.data(Qt.ItemDataRole.UserRole)
            if not name:
                return
            existing = self.db.get_placeholder_mapping(name)
            map_dialog = PlaceholderMappingDialog(dialog, name, self.db, existing)
            if map_dialog.exec() == QDialog.DialogCode.Accepted:
                mapping = map_dialog.get_mapping()
                if mapping:
                    self.db.add_placeholder_mapping(name, mapping)
                refresh_list()

        edit_btn.clicked.connect(on_edit)
        btn_layout.addWidget(edit_btn)

        delete_btn = QPushButton("Delete")
        def on_delete():
            item = mapping_list.currentItem()
            if not item:
                return
            name = item.data(Qt.ItemDataRole.UserRole)
            if not name:
                return
            reply = QMessageBox.question(
                dialog, "Delete Mapping",
                f"Delete mapping for {{{{{name}}}}}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.db.delete_placeholder_mapping(name)
                refresh_list()

        delete_btn.clicked.connect(on_delete)
        btn_layout.addWidget(delete_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)
        dialog.exec()

    def show_placeholder_help(self):
        """Show help dialog explaining placeholders."""
        help_text = """<h3>Template Placeholders</h3>
        <p>Placeholders are special markers that get replaced with actual case data.</p>

        <h4>Format</h4>
        <p>Use double curly braces: <code>{{PLACEHOLDER_NAME}}</code></p>

        <h4>Available Placeholders</h4>
        <table style='border-collapse: collapse; width: 100%;'>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{CASE_NAME}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>Full case name</td></tr>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{CASE_NUMBER}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>Court case number</td></tr>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{FILE_NUMBER}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>Internal file number</td></tr>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{CLIENT_NAME}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>Client's name</td></tr>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{PLAINTIFF}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>First plaintiff</td></tr>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{DEFENDANT}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>First defendant</td></tr>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{TODAY_DATE}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>Current date (MM/DD/YYYY)</td></tr>
        <tr><td style='padding: 4px; border: 1px solid #ddd;'><b>{{TODAY_DATE_LONG}}</b></td><td style='padding: 4px; border: 1px solid #ddd;'>Current date (Month DD, YYYY)</td></tr>
        </table>

        <h4>Custom Placeholders</h4>
        <p>Create your own placeholders for fields not in the list. They'll remain as <code>{{NAME}}</code> in the output if not found in case data.</p>
        """

        msg = QMessageBox(self)
        msg.setWindowTitle("Placeholder Help")
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(help_text)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def setup_format_toolbar(self):
        """Create formatting toolbar for template editor."""
        toolbar = QFrame()
        toolbar.setObjectName("formatToolbar")
        toolbar.setStyleSheet("""
            #formatToolbar {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f8f9fa, stop:1 #e9ecef);
                border: 1px solid #dee2e6;
                border-radius: 3px;
            }
            #formatToolbar QPushButton {
                border: 1px solid transparent;
                border-radius: 3px;
                background: transparent;
                padding: 2px 4px;
                min-height: 20px;
            }
            #formatToolbar QPushButton:hover { background: #dee2e6; border: 1px solid #adb5bd; }
            #formatToolbar QComboBox {
                background: white;
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 2px;
                min-height: 20px;
            }
        """)
        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(4, 2, 4, 2)
        layout.setSpacing(3)

        # Font family
        self.font_family_combo = QComboBox()
        self.font_family_combo.addItems(["Times New Roman", "Arial", "Calibri", "Courier New", "Georgia"])
        self.font_family_combo.setFixedWidth(115)
        self.font_family_combo.currentTextChanged.connect(lambda f: self.exec_format_cmd("fontName", f))
        layout.addWidget(self.font_family_combo)

        # Font size
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(["10", "11", "12", "14", "16", "18", "20", "24"])
        self.font_size_combo.setCurrentText("12")
        self.font_size_combo.setFixedWidth(45)
        self.font_size_combo.currentTextChanged.connect(lambda sz: self.exec_format_cmd("fontSize", sz))
        layout.addWidget(self.font_size_combo)

        # Separator
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.VLine)
        sep1.setStyleSheet("color: #ccc;")
        layout.addWidget(sep1)

        def create_btn(text, tooltip, cmd, value=None):
            btn = QPushButton(text)
            btn.setToolTip(tooltip)
            btn.setFixedSize(24, 24)
            btn.clicked.connect(lambda: self.exec_format_cmd(cmd, value))
            return btn

        layout.addWidget(create_btn("B", "Bold (Ctrl+B)", "bold"))
        layout.addWidget(create_btn("I", "Italic (Ctrl+I)", "italic"))
        layout.addWidget(create_btn("U", "Underline (Ctrl+U)", "underline"))

        # Separator
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.VLine)
        sep2.setStyleSheet("color: #ccc;")
        layout.addWidget(sep2)

        layout.addWidget(create_btn("⫷", "Align Left", "justifyLeft"))
        layout.addWidget(create_btn("≡", "Center", "justifyCenter"))
        layout.addWidget(create_btn("⫸", "Align Right", "justifyRight"))

        layout.addStretch()

        # Save button
        self.save_btn = QPushButton("Save")
        self.save_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5bc0de, stop:1 #2196F3);
                color: white;
                border: 1px solid #1976D2;
                border-radius: 3px;
                padding: 3px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6bd0ee, stop:1 #42a6f3);
            }
            QPushButton:disabled {
                background: #ccc;
                border-color: #aaa;
                color: #888;
            }
        """)
        self.save_btn.clicked.connect(self.save_template_content)
        self.save_btn.setEnabled(False)
        layout.addWidget(self.save_btn)

        toolbar.hide()  # Hidden by default
        return toolbar

    def exec_format_cmd(self, cmd, value=None):
        """Execute a formatting command on the editor."""
        if self.current_item_type != 'templates':
            return

        js = ""
        if cmd == "fontSize":
            size_map = {"10": "2", "11": "2", "12": "3", "14": "4", "16": "4", "18": "5", "20": "5", "24": "6"}
            js = f"document.execCommand('fontSize', false, '{size_map.get(value, '3')}');"
        elif cmd == "fontName":
            js = f"document.execCommand('fontName', false, '{value}');"
        else:
            js = f"document.execCommand('{cmd}', false, null);"

        self.editor.page().runJavaScript(js)
        # Trigger content change handler (for auto-save)
        self._on_editor_content_changed()

    def on_mode_changed(self, mode_text):
        """Handle mode selector change."""
        self.current_mode = 'templates' if mode_text == "Templates" else 'resources'
        self.tree.set_mode(self.current_mode)

        # Category filter is visible for both modes
        self.category_filter.setVisible(True)

        # Tags button only for templates
        is_templates = self.current_mode == 'templates'
        self.tag_filter_btn.setVisible(is_templates)
        self.new_btn.setText("New Template" if is_templates else "Open Folder")

        # Show/hide placeholder toolbar
        self.placeholder_toolbar.setVisible(is_templates)

        self.refresh_tree()
        self.clear_preview()

    def on_category_changed(self, category):
        """Handle category filter change."""
        self.refresh_tree()

    def show_tag_filter(self):
        """Show tag filter dialog."""
        all_tags = (self.db.get_all_template_tags() if self.current_mode == 'templates'
                   else self.db.get_all_resource_tags())

        if not all_tags:
            QMessageBox.information(self, "No Tags", "No tags have been created yet.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Filter by Tags")
        dialog.setMinimumWidth(250)
        layout = QVBoxLayout(dialog)

        layout.addWidget(QLabel("Select tags to filter by:"))

        tag_list = QListWidget()
        tag_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        for tag in all_tags:
            item = QListWidgetItem(tag)
            tag_list.addItem(item)
            if tag in self.active_tags:
                item.setSelected(True)
        layout.addWidget(tag_list)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        clear_btn = buttons.addButton("Clear", QDialogButtonBox.ButtonRole.ResetRole)

        def on_ok():
            self.active_tags = [item.text() for item in tag_list.selectedItems()]
            self.refresh_tree()
            dialog.accept()

        def on_clear():
            self.active_tags = []
            self.refresh_tree()
            dialog.accept()

        buttons.accepted.connect(on_ok)
        buttons.rejected.connect(dialog.reject)
        clear_btn.clicked.connect(on_clear)
        layout.addWidget(buttons)

        dialog.exec()

    def filter_tree(self, text):
        """Filter tree items by search text."""
        text = text.lower()
        root = self.tree.invisibleRootItem()
        self._filter_item(root, text)

    def _filter_item(self, item, text):
        """Recursively filter tree items."""
        child_visible = False
        for i in range(item.childCount()):
            child = item.child(i)
            if self._filter_item(child, text):
                child_visible = True

        item_text = item.text(0).lower() if item.text(0) else ""
        is_match = text in item_text or text == ""

        should_show = is_match or child_visible
        item.setHidden(not should_show)

        if should_show and child_visible:
            item.setExpanded(True)

        return should_show

    def refresh_tree(self):
        """Refresh the tree view."""
        self.tree.clear()

        # Update category filter (block signals to prevent recursion)
        current_cat = self.category_filter.currentText()
        self.category_filter.blockSignals(True)
        self.category_filter.clear()
        self.category_filter.addItem("All")

        if self.current_mode == 'templates':
            for cat in self.db.get_all_categories():
                self.category_filter.addItem(cat)
        else:
            # For resources, categories are subfolders
            for cat in self.get_resource_categories():
                self.category_filter.addItem(cat)

        idx = self.category_filter.findText(current_cat)
        if idx >= 0:
            self.category_filter.setCurrentIndex(idx)
        self.category_filter.blockSignals(False)

        if self.current_mode == 'templates':
            self.populate_templates_tree()
        else:
            self.populate_resources_tree()

        # Re-apply filter
        self.filter_tree(self.search_input.text())

    def get_resource_categories(self):
        """Get list of subfolder names in Resources directory as categories."""
        categories = []
        if os.path.exists(RESOURCES_DIR):
            for item in sorted(os.listdir(RESOURCES_DIR)):
                item_path = os.path.join(RESOURCES_DIR, item)
                if os.path.isdir(item_path):
                    categories.append(item)
        return categories

    def populate_templates_tree(self):
        """Populate tree with templates."""
        if not os.path.exists(TEMPLATES_DIR):
            return

        category_filter = self.category_filter.currentText()
        if category_filter == "All":
            category_filter = None

        icon_provider = QFileIconProvider()

        # Get templates that match tag filter
        if self.active_tags:
            matching_paths = set()
            for tag in self.active_tags:
                templates = self.db.get_templates_by_tag(tag)
                for t in templates:
                    matching_paths.add(t['relative_path'])
        else:
            matching_paths = None  # No filter

        for root, dirs, files in os.walk(TEMPLATES_DIR):
            rel_root = os.path.relpath(root, TEMPLATES_DIR)

            if rel_root == '.':
                parent_item = self.tree.invisibleRootItem()
            else:
                # Find or create parent items
                parts = rel_root.split(os.sep)
                parent_item = self.tree.invisibleRootItem()
                for part in parts:
                    found = None
                    for i in range(parent_item.childCount()):
                        child = parent_item.child(i)
                        if child.text(0) == part and child.data(0, Qt.ItemDataRole.UserRole + 1) == 'folder':
                            found = child
                            break
                    if found:
                        parent_item = found
                    else:
                        folder_item = QTreeWidgetItem([part])
                        folder_path = os.path.join(TEMPLATES_DIR, *parts[:parts.index(part)+1])
                        folder_item.setData(0, Qt.ItemDataRole.UserRole, folder_path)
                        folder_item.setData(0, Qt.ItemDataRole.UserRole + 1, 'folder')
                        folder_item.setIcon(0, icon_provider.icon(QFileIconProvider.IconType.Folder))
                        parent_item.addChild(folder_item)
                        parent_item = folder_item

            # Add files
            for filename in sorted(files):
                ext = os.path.splitext(filename)[1].lower()
                if ext not in TEMPLATE_EXTENSIONS:
                    continue

                file_path = os.path.join(root, filename)
                relative_path = os.path.relpath(file_path, TEMPLATES_DIR)

                # Apply tag filter
                if matching_paths is not None and relative_path not in matching_paths:
                    continue

                # Apply category filter
                if category_filter:
                    template = self.db.get_template_by_path(relative_path)
                    if not template or template.get('category') != category_filter:
                        continue

                # Ensure template is in database
                self.db.upsert_template(relative_path, filename)

                file_item = QTreeWidgetItem([filename])
                file_item.setData(0, Qt.ItemDataRole.UserRole, file_path)
                file_item.setData(0, Qt.ItemDataRole.UserRole + 1, 'file')
                file_item.setIcon(0, icon_provider.icon(QFileInfo(file_path)))
                parent_item.addChild(file_item)

        self.tree.expandAll()

    def populate_resources_tree(self):
        """Populate tree with resources organized by category (subfolder)."""
        if not os.path.exists(RESOURCES_DIR):
            return

        icon_provider = QFileIconProvider()

        # Get selected category filter
        category_filter = self.category_filter.currentText()
        if category_filter == "All":
            category_filter = None

        # Determine the base directory to scan
        if category_filter:
            # Only scan the selected category subfolder
            scan_dir = os.path.join(RESOURCES_DIR, category_filter)
            if not os.path.exists(scan_dir):
                return
        else:
            scan_dir = RESOURCES_DIR

        for root, dirs, files in os.walk(scan_dir):
            rel_root = os.path.relpath(root, scan_dir)

            if rel_root == '.':
                parent_item = self.tree.invisibleRootItem()
            else:
                parts = rel_root.split(os.sep)
                parent_item = self.tree.invisibleRootItem()
                for part in parts:
                    found = None
                    for i in range(parent_item.childCount()):
                        child = parent_item.child(i)
                        if child.text(0) == part and child.data(0, Qt.ItemDataRole.UserRole + 1) == 'folder':
                            found = child
                            break
                    if found:
                        parent_item = found
                    else:
                        folder_item = QTreeWidgetItem([part])
                        folder_path = os.path.join(scan_dir, *parts[:parts.index(part)+1])
                        folder_item.setData(0, Qt.ItemDataRole.UserRole, folder_path)
                        folder_item.setData(0, Qt.ItemDataRole.UserRole + 1, 'folder')
                        folder_item.setIcon(0, icon_provider.icon(QFileIconProvider.IconType.Folder))
                        parent_item.addChild(folder_item)
                        parent_item = folder_item

            for filename in sorted(files):
                ext = os.path.splitext(filename)[1].lower()
                if ext not in RESOURCE_EXTENSIONS:
                    continue

                file_path = os.path.join(root, filename)

                file_item = QTreeWidgetItem([filename])
                file_item.setData(0, Qt.ItemDataRole.UserRole, file_path)
                file_item.setData(0, Qt.ItemDataRole.UserRole + 1, 'file')
                file_item.setIcon(0, icon_provider.icon(QFileInfo(file_path)))
                parent_item.addChild(file_item)

        self.tree.expandAll()

    def on_item_selected(self, path, item_type):
        """Handle item selection in tree."""
        # Stop any pending auto-save for the previous template
        self.auto_save_timer.stop()

        # With auto-save, changes are saved automatically, but check just in case
        if self.editor_modified:
            # Perform a final save before switching
            self._save_template_silent()

        self.current_item_path = path
        self.current_item_type = item_type
        self.editor_modified = False
        self.save_btn.setEnabled(False)

        ext = os.path.splitext(path)[1].lower()

        if ext == '.pdf':
            self.show_pdf_preview(path)
        elif item_type == 'templates' or ext in ['.txt', '.html', '.docx', '.doc']:
            self.show_template_editor(path)
        else:
            self.show_pdf_preview(path)  # Try to preview other files

    def on_item_deleted(self, path):
        """Handle item deletion - clear preview if the deleted item was being viewed."""
        if self.current_item_path == path:
            self.clear_preview()
            self.status_label.setText("Template deleted")

    def _install_editor_event_filter(self):
        """Install event filter on the editor's focus proxy widget."""
        focus_proxy = self.editor.focusProxy()
        if focus_proxy:
            focus_proxy.installEventFilter(self)
            self._editor_focus_proxy = focus_proxy

    def eventFilter(self, obj, event):
        """Filter events to catch Tab key in editor before Qt handles it."""
        # Check if event is from editor or its focus proxy
        is_editor_event = (obj == self.editor or
                          (hasattr(self, '_editor_focus_proxy') and obj == self._editor_focus_proxy))
        if is_editor_event and event.type() == QEvent.Type.KeyPress:
            key_event = event
            if key_event.key() == Qt.Key.Key_Tab:
                # Only handle Tab if we're editing a template
                if self.current_item_type == 'templates':
                    shift_pressed = bool(key_event.modifiers() & Qt.KeyboardModifier.ShiftModifier)
                    self.apply_tab_indent(shift_pressed)
                    return True  # Event handled, don't propagate
        return super().eventFilter(obj, event)

    def apply_tab_indent(self, remove_indent=False):
        """Apply or remove indent via JavaScript. Handles paragraphs and list items."""
        js = """
        (function() {
            var sel = window.getSelection();
            if (sel.rangeCount === 0) return;
            var range = sel.getRangeAt(0);
            var node = range.startContainer;
            var editor = document.getElementById('editor');
            var removeIndent = """ + ('true' if remove_indent else 'false') + """;

            // Check if we're in a list item
            var listItem = node;
            while (listItem && listItem !== editor) {
                if (listItem.nodeType === 1 && listItem.tagName === 'LI') {
                    break;
                }
                listItem = listItem.parentNode;
            }

            if (listItem && listItem.tagName === 'LI') {
                // Handle list item indentation
                var currentMargin = parseInt(listItem.style.marginLeft) || 0;
                if (removeIndent) {
                    // Decrease indent (minimum 0)
                    var newMargin = Math.max(0, currentMargin - 36); // 0.5in ≈ 36px
                    listItem.style.marginLeft = newMargin > 0 ? newMargin + 'px' : '';
                } else {
                    // Increase indent
                    listItem.style.marginLeft = (currentMargin + 36) + 'px';
                }
                return;
            }

            // Check if we're in a list (UL/OL) but not in a specific LI
            var list = node;
            while (list && list !== editor) {
                if (list.nodeType === 1 && (list.tagName === 'UL' || list.tagName === 'OL')) {
                    break;
                }
                list = list.parentNode;
            }

            if (list && (list.tagName === 'UL' || list.tagName === 'OL')) {
                // Indent the entire list
                var currentMargin = parseInt(list.style.marginLeft) || 0;
                if (removeIndent) {
                    var newMargin = Math.max(0, currentMargin - 36);
                    list.style.marginLeft = newMargin > 0 ? newMargin + 'px' : '';
                } else {
                    list.style.marginLeft = (currentMargin + 36) + 'px';
                }
                return;
            }

            // Handle regular paragraph/block element
            var block = node;
            while (block && block !== editor && block.parentNode !== editor) {
                block = block.parentNode;
            }

            // If we're in a text node directly under editor, wrap it in a p tag
            if (!block || block === editor || block.nodeType === 3) {
                if (node.nodeType === 3) {
                    var p = document.createElement('p');
                    p.style.margin = '0';
                    node.parentNode.insertBefore(p, node);
                    p.appendChild(node);
                    block = p;
                    // Restore cursor
                    var newRange = document.createRange();
                    newRange.setStart(node, range.startOffset);
                    newRange.collapse(true);
                    sel.removeAllRanges();
                    sel.addRange(newRange);
                } else {
                    return;
                }
            }

            // Apply first-line indent to paragraph
            if (block && block.nodeType === 1) {
                if (removeIndent) {
                    block.style.textIndent = '';
                } else {
                    block.style.textIndent = '0.5in';
                }
            }
        })();
        """
        self.editor.page().runJavaScript(js)
        self._on_editor_content_changed()

    def show_pdf_preview(self, path):
        """Display PDF in preview."""
        self.format_toolbar.hide()
        self.placeholder_toolbar.hide()
        self.copy_raw_btn.setEnabled(False)
        self.copy_filled_btn.setEnabled(False)

        clean_path = path.replace('\\', '/')
        url = f"local-resource:///{clean_path}"
        self.pdf_preview.setUrl(QUrl(url))
        self.content_stack.setCurrentIndex(1)

        self.status_label.setText(f"Viewing: {os.path.basename(path)}")

    def show_template_editor(self, path):
        """Load and display template in editor."""
        self.format_toolbar.show()
        self.placeholder_toolbar.show()
        self.copy_raw_btn.setEnabled(True)
        self.copy_filled_btn.setEnabled(True)

        content = self.load_template_content(path)
        variables = self.get_case_variables()
        self.editor.setHtml(self.get_editor_html(content, variables))
        self.content_stack.setCurrentIndex(2)

        # Update status with variable info
        placeholders = self.extract_placeholders(content)
        resolved = sum(1 for p in placeholders if p in variables and variables[p])

        if placeholders:
            self.status_label.setText(f"Found {len(placeholders)} variables, {resolved} resolved")
        else:
            self.status_label.setText("No variables in this template")

    def load_template_content(self, path):
        """Load template file content."""
        ext = os.path.splitext(path)[1].lower()

        try:
            if ext == '.txt':
                with open(path, 'r', encoding='utf-8') as f:
                    text = f.read()
                return html.escape(text).replace('\n', '<br>')

            elif ext == '.html':
                with open(path, 'r', encoding='utf-8') as f:
                    return f.read()

            elif ext == '.docx':
                try:
                    from docx import Document
                    doc = Document(path)
                    html_parts = []
                    for para in doc.paragraphs:
                        html_parts.append(f"<p>{html.escape(para.text)}</p>")
                    return '\n'.join(html_parts)
                except ImportError:
                    return "<p>python-docx not installed. Run: pip install python-docx</p>"

            elif ext == '.doc':
                # Use win32com on Windows to read .doc files
                try:
                    import win32com.client
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(path)
                    text = doc.Content.Text
                    doc.Close(False)
                    word.Quit()
                    return html.escape(text).replace('\n', '<br>').replace('\r', '')
                except ImportError:
                    return "<p>pywin32 not installed. Run: pip install pywin32</p>"
                except Exception as e:
                    return f"<p>Error reading .doc file: {html.escape(str(e))}</p>"

        except Exception as e:
            log_event(f"Error loading template: {e}", "error")
            return f"<p>Error loading file: {e}</p>"

        return ""

    def _on_editor_content_changed(self):
        """Called when editor content changes - starts/restarts the auto-save timer."""
        if self.current_item_type == 'templates' and self.current_item_path:
            self.editor_modified = True
            self.save_btn.setEnabled(True)
            # Restart the auto-save timer (debounce)
            self.auto_save_timer.stop()
            self.auto_save_timer.start()

    def _perform_auto_save(self):
        """Perform the actual auto-save when the timer fires."""
        if self.current_item_type == 'templates' and self.current_item_path and self.editor_modified:
            self._save_template_silent()

    def _save_template_silent(self):
        """Save template content without showing dialogs (for auto-save)."""
        if not self.current_item_path:
            return

        def on_html_ready(html_content):
            if html_content is None:
                return

            ext = os.path.splitext(self.current_item_path)[1].lower()

            try:
                if ext == '.txt':
                    # Strip HTML tags for plain text
                    text = re.sub(r'<br\s*/?>', '\n', html_content)
                    text = re.sub(r'<[^>]+>', '', text)
                    text = html.unescape(text)
                    with open(self.current_item_path, 'w', encoding='utf-8') as f:
                        f.write(text)

                elif ext == '.html':
                    with open(self.current_item_path, 'w', encoding='utf-8') as f:
                        f.write(html_content)

                elif ext == '.docx':
                    # Can't auto-save docx, skip silently
                    return

                self.editor_modified = False
                self.save_btn.setEnabled(False)
                self.status_label.setText("Auto-saved")
                log_event(f"Auto-saved template: {self.current_item_path}", "info")

            except Exception as e:
                log_event(f"Error auto-saving template: {e}", "error")
                self.status_label.setText(f"Auto-save failed: {str(e)[:30]}")

        # Get content from editor
        self.editor.page().runJavaScript(
            "document.getElementById('editor').innerHTML",
            on_html_ready
        )

    def save_template_content(self):
        """Save editor content back to file."""
        if not self.current_item_path:
            return
        # Stop auto-save timer since we're saving manually
        self.auto_save_timer.stop()

        def on_html_ready(html_content):
            ext = os.path.splitext(self.current_item_path)[1].lower()

            try:
                if ext == '.txt':
                    # Strip HTML tags for plain text
                    import re
                    text = re.sub(r'<br\s*/?>', '\n', html_content)
                    text = re.sub(r'<[^>]+>', '', text)
                    text = html.unescape(text)
                    with open(self.current_item_path, 'w', encoding='utf-8') as f:
                        f.write(text)

                elif ext == '.html':
                    with open(self.current_item_path, 'w', encoding='utf-8') as f:
                        f.write(html_content)

                elif ext == '.docx':
                    QMessageBox.warning(self, "Save",
                        "Saving to .docx is not yet supported. Save as .html instead.")
                    return

                self.editor_modified = False
                self.save_btn.setEnabled(False)
                log_event(f"Saved template: {self.current_item_path}", "info")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save: {e}")
                log_event(f"Error saving template: {e}", "error")

        # Get content from editor
        self.editor.page().runJavaScript(
            "document.getElementById('editor').innerHTML",
            on_html_ready
        )

    def get_editor_html(self, content, variables=None):
        """Generate HTML wrapper for editor."""
        # Convert variables to JavaScript object
        if variables:
            vars_js = json.dumps(variables)
        else:
            vars_js = '{}'

        return f"""<!DOCTYPE html>
<html>
<head>
    <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
    <style>
        body {{ margin: 0; padding: 0; }}
        #editor {{
            font-family: 'Times New Roman', serif;
            font-size: 12pt;
            padding: 1in;
            min-height: 100vh;
            outline: none;
            line-height: 1.5;
        }}
        .placeholder {{
            background-color: #fff3cd;
            border: 1px dashed #856404;
            padding: 0 4px;
            border-radius: 2px;
        }}
    </style>
</head>
<body>
    <div id="editor" contenteditable="true">{content}</div>
    <script>
        // Variables for placeholder replacement
        var templateVariables = {vars_js};
        var bridge = null;

        // Initialize QWebChannel for Python communication
        new QWebChannel(qt.webChannelTransport, function(channel) {{
            bridge = channel.objects.bridge;
        }});

        // Function to replace placeholders in text
        function fillPlaceholders(text) {{
            return text.replace(/\\{{\\{{(\\w+)\\}}\\}}/g, function(match, varName) {{
                return templateVariables[varName] !== undefined ? templateVariables[varName] : match;
            }});
        }}

        // Intercept copy event to fill placeholders
        document.getElementById('editor').addEventListener('copy', function(e) {{
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {{
                var selectedText = selection.toString();
                if (selectedText) {{
                    var filledText = fillPlaceholders(selectedText);
                    e.clipboardData.setData('text/plain', filledText);
                    e.preventDefault();
                }}
            }}
        }});

        // Notify Python when content changes (for auto-save)
        document.getElementById('editor').addEventListener('input', function() {{
            if (bridge) {{
                bridge.notifyContentChanged();
            }}
        }});
    </script>
</body>
</html>"""

    def clear_preview(self):
        """Clear preview and reset state."""
        self.current_item_path = None
        self.current_item_type = None
        self.content_stack.setCurrentIndex(0)
        self.format_toolbar.hide()
        self.placeholder_toolbar.hide()
        self.copy_raw_btn.setEnabled(False)
        self.copy_filled_btn.setEnabled(False)
        self.status_label.setText("")

    def get_case_variables(self):
        """Build variable dictionary from currently selected case."""
        variables = {}

        # Add computed variables first (always available)
        variables['TODAY_DATE'] = datetime.now().strftime('%m/%d/%Y')
        variables['TODAY_DATE_LONG'] = datetime.now().strftime('%B %d, %Y')

        # Get file number from main window
        file_number = None
        if self.main_window:
            if hasattr(self.main_window, 'file_number'):
                file_number = self.main_window.file_number

        if not file_number:
            # Still apply custom mappings even without a case loaded
            mappings = self.db.get_all_placeholder_mappings()
            for custom_name, maps_to in mappings.items():
                if maps_to in variables:
                    variables[custom_name] = variables[maps_to]
            return variables

        variables['FILE_NUMBER'] = file_number

        # Load case JSON
        json_path = os.path.join(GEMINI_DATA_DIR, f"{file_number}.json")
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    case_data = json.load(f)

                # Map JSON fields to variable names
                field_map = {
                    'CASE_NAME': 'case_name',
                    'CASE_NUMBER': 'case_number',
                    'FILE_NUMBER': 'file_number',
                    'CLIENT_NAME': 'client_name',
                    'VENUE_COUNTY': 'venue_county',
                    'TRIAL_DATE': 'trial_date',
                    'CLAIM_NUMBER': 'claim_number',
                    'INCIDENT_DATE': 'incident_date',
                    'FILING_DATE': 'filing_date',
                    'ADJUSTER_NAME': 'adjuster_name',
                    'ADJUSTER_EMAIL': 'adjuster_email',
                    'PLAINTIFF_COUNSEL': 'plaintiff_counsel',
                    'FACTUAL_BACKGROUND': 'factual_background',
                    'PROCEDURAL_HISTORY': 'procedural_history',
                }

                for var_name, json_key in field_map.items():
                    if json_key in case_data:
                        val = case_data[json_key]
                        if isinstance(val, dict) and 'value' in val:
                            val = val['value']
                        variables[var_name] = str(val) if val else ''

                # Handle plaintiffs/defendants (arrays)
                if 'plaintiffs' in case_data:
                    pls = case_data['plaintiffs']
                    if isinstance(pls, dict) and 'value' in pls:
                        pls = pls['value']
                    if isinstance(pls, list):
                        variables['PLAINTIFF'] = pls[0] if pls else ''
                        variables['PLAINTIFFS'] = ', '.join(pls)
                    else:
                        variables['PLAINTIFF'] = str(pls) if pls else ''
                        variables['PLAINTIFFS'] = str(pls) if pls else ''

                if 'defendants' in case_data:
                    dfs = case_data['defendants']
                    if isinstance(dfs, dict) and 'value' in dfs:
                        dfs = dfs['value']
                    if isinstance(dfs, list):
                        variables['DEFENDANT'] = dfs[0] if dfs else ''
                        variables['DEFENDANTS'] = ', '.join(dfs)
                    else:
                        variables['DEFENDANT'] = str(dfs) if dfs else ''
                        variables['DEFENDANTS'] = str(dfs) if dfs else ''

            except Exception as e:
                log_event(f"Error loading case data: {e}", "error")

        # Apply custom placeholder mappings
        mappings = self.db.get_all_placeholder_mappings()
        for custom_name, maps_to in mappings.items():
            if maps_to in variables:
                variables[custom_name] = variables[maps_to]

        return variables

    def extract_placeholders(self, content):
        """Find all {{VARIABLE}} placeholders in content."""
        pattern = r'\{\{(\w+)\}\}'
        matches = re.findall(pattern, content)
        return list(set(matches))

    def fill_variables(self, content, variables):
        """Replace all {{VARIABLE}} placeholders with actual values."""
        def replacer(match):
            var_name = match.group(1)
            return variables.get(var_name, match.group(0))

        return re.sub(r'\{\{(\w+)\}\}', replacer, content)

    def copy_raw(self):
        """Copy template content with placeholders intact."""
        if not self.current_item_path:
            return

        def on_html_ready(html_content):
            if html_content is None:
                return
            # Convert HTML to plain text
            text = re.sub(r'<br\s*/?>', '\n', html_content)
            text = re.sub(r'<[^>]+>', '', text)
            text = html.unescape(text)

            # Clear and set clipboard
            clipboard = QApplication.clipboard()
            clipboard.clear()
            clipboard.setText(text)
            self.status_label.setText("Copied raw template to clipboard")

        self.editor.page().runJavaScript(
            "document.getElementById('editor').innerHTML",
            on_html_ready
        )

    def copy_filled(self):
        """Copy template content with case data filled in."""
        if not self.current_item_path:
            return

        variables = self.get_case_variables()

        def on_html_ready(html_content):
            if html_content is None:
                return

            # Convert HTML to plain text
            text = re.sub(r'<br\s*/?>', '\n', html_content)
            text = re.sub(r'<[^>]+>', '', text)
            text = html.unescape(text)

            # Fill in variables
            filled = self.fill_variables(text, variables)

            # Clear and set clipboard
            clipboard = QApplication.clipboard()
            clipboard.clear()
            clipboard.setText(filled)

            # Check for unfilled placeholders
            remaining = self.extract_placeholders(filled)
            if remaining:
                self.status_label.setText(f"Copied! ({len(remaining)} unfilled: {', '.join(remaining[:3])}...)")
            else:
                self.status_label.setText("Copied filled template to clipboard")

        self.editor.page().runJavaScript(
            "document.getElementById('editor').innerHTML",
            on_html_ready
        )

    def create_new_template(self):
        """Create a new template file."""
        if self.current_mode != 'templates':
            self.open_current_folder()
            return

        name, ok = QInputDialog.getText(
            self, "New Template", "Template name (without extension):"
        )

        if ok and name:
            filename = f"{name}.html"
            filepath = os.path.join(TEMPLATES_DIR, filename)

            if os.path.exists(filepath):
                QMessageBox.warning(self, "Error", "A template with this name already exists.")
                return

            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write("<p>New template content...</p>")

                self.refresh_tree()
                log_event(f"Created new template: {filename}", "info")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to create template: {e}")

    def open_current_folder(self):
        """Open the current folder in Explorer."""
        if self.current_mode == 'templates':
            folder = TEMPLATES_DIR
        else:
            # For resources, open the selected category subfolder if one is selected
            category = self.category_filter.currentText()
            if category and category != "All":
                folder = os.path.join(RESOURCES_DIR, category)
            else:
                folder = RESOURCES_DIR
        try:
            os.startfile(folder)
        except Exception as e:
            log_event(f"Error opening folder: {e}", "error")
