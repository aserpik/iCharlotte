"""
Word Tab - Embeds Microsoft Word directly into iCharlotte.

Uses win32gui to reparent the Word application window into a Qt widget,
providing full Word functionality including the ribbon toolbar.
"""

import os
import json
import time

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QMessageBox, QSplitter, QListWidget, QListWidgetItem,
    QFrame, QApplication, QSizePolicy
)
from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QWindow

try:
    import win32com.client
    import win32gui
    import win32con
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

from icharlotte_core.config import GEMINI_DATA_DIR

# Windows API constants
GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_CHILD = 0x40000000
WS_VISIBLE = 0x10000000
WS_CLIPSIBLINGS = 0x04000000
WS_CLIPCHILDREN = 0x02000000


class WordEmbedWidget(QWidget):
    """Widget that hosts an embedded Word window via win32gui reparenting."""

    word_closed = Signal()
    document_saved = Signal(str)
    document_opened = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.word_app = None
        self.word_doc = None
        self.word_hwnd = 0
        self.original_style = 0
        self.original_parent = 0
        self._com_initialized = False

        self.setMinimumSize(400, 300)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # Create placeholder label
        self.placeholder = QLabel("Click 'New Document' or 'Open Document' to start Word")
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder.setStyleSheet("color: #666; font-size: 14px; background-color: #f5f5f5;")
        self.placeholder.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.placeholder)

        # Timer to check if Word is still running
        self.check_timer = QTimer(self)
        self.check_timer.timeout.connect(self._check_word_alive)

        # Timer for delayed embedding
        self.embed_timer = QTimer(self)
        self.embed_timer.setSingleShot(True)
        self.embed_timer.timeout.connect(self._do_embed)

    def _init_com(self):
        """Initialize COM for this thread."""
        if not self._com_initialized:
            pythoncom.CoInitialize()
            self._com_initialized = True

    def new_document(self):
        """Create a new Word document."""
        self._init_com()

        # If Word is already embedded, close it first and create fresh instance
        # This is more reliable than trying to add a document to an embedded instance
        if self.word_hwnd:
            self.close_word(save=False)

        return self._launch_word(None)

    def open_document(self, file_path):
        """Open an existing document."""
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "File Not Found", f"File does not exist:\n{file_path}")
            return False

        self._init_com()

        # If Word is already embedded, close it first and relaunch with the new document
        # This is more reliable than trying to open a new document in an embedded instance
        if self.word_hwnd:
            self.close_word(save=False)

        return self._launch_word(file_path)

    def _launch_word(self, file_path=None):
        """Launch Word and embed it."""
        self.placeholder.setText("Starting Microsoft Word...")
        self.placeholder.show()
        QApplication.processEvents()

        try:
            # Reinitialize COM for fresh state
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            pythoncom.CoInitialize()
            self._com_initialized = True

            # Create fresh Word application
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True

            # Prevent Word from showing its own dialogs
            self.word_app.DisplayAlerts = False

            # Open or create document
            if file_path:
                self.word_doc = self.word_app.Documents.Open(file_path)
                self.document_opened.emit(file_path)
            else:
                self.word_doc = self.word_app.Documents.Add()

            # Give Word time to fully initialize before embedding
            self.embed_timer.start(800)
            return True

        except Exception as e:
            self.placeholder.setText(f"Failed to start Word: {e}")
            QMessageBox.warning(self, "Word Error", f"Could not start Microsoft Word:\n{e}")
            return False

    def _do_embed(self):
        """Find and embed the Word window."""
        try:
            # Find the Word window
            self.word_hwnd = self._find_word_window()

            if self.word_hwnd:
                self._embed_window(self.word_hwnd)
                self.placeholder.hide()
                self.check_timer.start(1000)  # Check every second if Word is still alive
                print(f"Word embedded successfully, hwnd={self.word_hwnd}")
            else:
                self.placeholder.setText("Word started but window not found.\nTry clicking 'New Document' again.")
                print("Could not find Word window")

        except Exception as e:
            self.placeholder.setText(f"Embedding failed: {e}")
            print(f"Embedding error: {e}")

    def _find_word_window(self):
        """Find the Word application window handle."""
        results = []

        def enum_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                class_name = win32gui.GetClassName(hwnd)
                # Word's main window class is OpusApp
                if class_name == "OpusApp":
                    results.append(hwnd)
            return True

        win32gui.EnumWindows(enum_callback, None)

        if results:
            # Return the first (most recent) Word window
            return results[0]
        return 0

    def _embed_window(self, hwnd):
        """Embed the Word window into this widget."""
        try:
            # Store original window properties for restoration later
            self.original_style = win32gui.GetWindowLong(hwnd, GWL_STYLE)
            self.original_parent = win32gui.GetParent(hwnd)

            # Modify style: remove title bar/borders, make it a child window
            new_style = (self.original_style & ~win32con.WS_POPUP & ~win32con.WS_CAPTION
                        & ~win32con.WS_THICKFRAME & ~win32con.WS_MINIMIZEBOX
                        & ~win32con.WS_MAXIMIZEBOX & ~win32con.WS_SYSMENU)
            new_style = new_style | WS_CHILD | WS_VISIBLE

            win32gui.SetWindowLong(hwnd, GWL_STYLE, new_style)

            # Get our widget's native window handle
            self_hwnd = int(self.winId())

            # Reparent the Word window to our widget
            win32gui.SetParent(hwnd, self_hwnd)

            # Resize to fill our widget
            self._resize_word_window()

            # Force a repaint
            win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE |
                                  win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)

        except Exception as e:
            print(f"Error embedding Word window: {e}")
            raise

    def _resize_word_window(self):
        """Resize the embedded Word window to fill this widget."""
        if self.word_hwnd and win32gui.IsWindow(self.word_hwnd):
            rect = self.rect()
            try:
                win32gui.MoveWindow(
                    self.word_hwnd,
                    0, 0,
                    rect.width(), rect.height(),
                    True
                )
            except Exception as e:
                print(f"Error resizing Word window: {e}")

    def resizeEvent(self, event):
        """Handle resize to update embedded window size."""
        super().resizeEvent(event)
        self._resize_word_window()

    def _check_word_alive(self):
        """Check if Word is still running."""
        if self.word_hwnd:
            if not win32gui.IsWindow(self.word_hwnd):
                self._on_word_closed()

    def _on_word_closed(self):
        """Handle Word being closed externally."""
        self.check_timer.stop()
        self.word_app = None
        self.word_doc = None
        self.word_hwnd = 0
        self.placeholder.setText("Word was closed.\nClick 'New Document' or 'Open Document' to restart.")
        self.placeholder.show()
        self.word_closed.emit()

    def save_document(self):
        """Save the current document."""
        if self.word_doc:
            try:
                if self.word_doc.Path:
                    self.word_doc.Save()
                    self.document_saved.emit(self.word_doc.FullName)
                    return True
                else:
                    return self.save_document_as()
            except Exception as e:
                QMessageBox.warning(self, "Save Error", f"Could not save document:\n{e}")
        return False

    def save_document_as(self, file_path=None):
        """Save the document with a new name."""
        if not self.word_doc:
            return False

        if not file_path:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Document As",
                "", "Word Documents (*.docx);;All Files (*.*)"
            )

        if file_path:
            try:
                if not file_path.lower().endswith('.docx'):
                    file_path += '.docx'
                self.word_doc.SaveAs2(file_path)
                self.document_saved.emit(file_path)
                return True
            except Exception as e:
                QMessageBox.warning(self, "Save Error", f"Could not save document:\n{e}")
        return False

    def get_current_file(self):
        """Get the path of the currently open file."""
        if self.word_doc:
            try:
                if self.word_doc.Path:
                    return self.word_doc.FullName
            except:
                pass
        return None

    def close_word(self, save=True):
        """Close Word and restore the window."""
        self.check_timer.stop()
        self.embed_timer.stop()

        # Don't try to use COM to close - it's unreliable after embedding
        # Just force close the window and kill the process

        # First, try to restore the window (so it can close properly)
        if self.word_hwnd and self.original_style:
            try:
                if win32gui.IsWindow(self.word_hwnd):
                    win32gui.SetWindowLong(self.word_hwnd, GWL_STYLE, self.original_style)
                    win32gui.SetParent(self.word_hwnd, 0)
            except Exception as e:
                print(f"Error restoring window: {e}")

        # Clear COM references (don't try to call methods on them)
        self.word_doc = None
        self.word_app = None

        # Force close the Word window
        if self.word_hwnd:
            try:
                if win32gui.IsWindow(self.word_hwnd):
                    # Send close message
                    win32gui.PostMessage(self.word_hwnd, win32con.WM_CLOSE, 0, 0)
                    time.sleep(0.3)

                    # If still exists, force destroy
                    if win32gui.IsWindow(self.word_hwnd):
                        win32gui.DestroyWindow(self.word_hwnd)
            except Exception as e:
                print(f"Error closing window: {e}")

        # Kill any Word processes as a last resort
        self._kill_word_processes()

        self.word_hwnd = 0
        self.original_style = 0
        self.original_parent = 0
        self._com_initialized = False  # Reset so we reinitialize COM

        # Wait for Word to fully close
        time.sleep(0.5)
        QApplication.processEvents()

        self.placeholder.setText("Click 'New Document' or 'Open Document' to start Word")
        self.placeholder.show()

    def _kill_word_processes(self):
        """Kill any Word processes that might be hanging."""
        import subprocess
        try:
            # Use taskkill to ensure Word is closed
            subprocess.run(
                ['taskkill', '/F', '/IM', 'WINWORD.EXE'],
                capture_output=True,
                timeout=5
            )
        except Exception as e:
            print(f"Error killing Word processes: {e}")


class WordTab(QWidget):
    """Tab widget for embedded Microsoft Word functionality."""

    def __init__(self, main_window=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.recent_files = []

        if not HAS_WIN32:
            self._show_error_ui()
            return

        self.setup_ui()
        self.load_recent_files()

    def _show_error_ui(self):
        """Show error message if win32 modules are not available."""
        layout = QVBoxLayout(self)
        label = QLabel(
            "Microsoft Word embedding requires pywin32.\n\n"
            "Install it with: pip install pywin32"
        )
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)

    def setup_ui(self):
        """Build the UI."""
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(5, 5, 5, 5)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # Left panel - Controls and recent files
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_panel.setFixedWidth(200)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Document controls
        controls_label = QLabel("Document Controls")
        controls_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        left_layout.addWidget(controls_label)

        btn_new = QPushButton("New Document")
        btn_new.clicked.connect(self.new_document)
        left_layout.addWidget(btn_new)

        btn_open = QPushButton("Open Document")
        btn_open.clicked.connect(self.open_document)
        left_layout.addWidget(btn_open)

        btn_save = QPushButton("Save")
        btn_save.clicked.connect(self.save_document)
        left_layout.addWidget(btn_save)

        btn_save_as = QPushButton("Save As...")
        btn_save_as.clicked.connect(self.save_document_as)
        left_layout.addWidget(btn_save_as)

        left_layout.addSpacing(10)

        btn_close = QPushButton("Close Word")
        btn_close.setStyleSheet("background-color: #f44336; color: white;")
        btn_close.clicked.connect(self.close_word)
        left_layout.addWidget(btn_close)

        left_layout.addSpacing(20)

        # Recent files
        recent_label = QLabel("Recent Files")
        recent_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        left_layout.addWidget(recent_label)

        self.recent_list = QListWidget()
        self.recent_list.itemDoubleClicked.connect(self._open_recent_file)
        left_layout.addWidget(self.recent_list)

        left_layout.addStretch()

        # Quick actions for case integration
        if self.main_window:
            case_label = QLabel("Case Integration")
            case_label.setStyleSheet("font-weight: bold; font-size: 12px;")
            left_layout.addWidget(case_label)

            btn_open_case = QPushButton("Open from Case")
            btn_open_case.clicked.connect(self.open_from_case)
            left_layout.addWidget(btn_open_case)

            btn_save_case = QPushButton("Save to Case")
            btn_save_case.clicked.connect(self.save_to_case)
            left_layout.addWidget(btn_save_case)

        splitter.addWidget(left_panel)

        # Right panel - Embedded Word
        self.word_widget = WordEmbedWidget()
        self.word_widget.document_saved.connect(self._on_document_saved)
        self.word_widget.document_opened.connect(self._on_document_opened)
        splitter.addWidget(self.word_widget)

        splitter.setSizes([200, 800])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

    def new_document(self):
        """Create a new Word document."""
        self.word_widget.new_document()

    def open_document(self):
        """Open an existing document."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Document",
            "", "Word Documents (*.docx *.doc);;All Files (*.*)"
        )
        if file_path:
            if self.word_widget.open_document(file_path):
                self._add_to_recent(file_path)

    def save_document(self):
        """Save the current document."""
        self.word_widget.save_document()

    def save_document_as(self):
        """Save the document with a new name."""
        if self.word_widget.save_document_as():
            current = self.word_widget.get_current_file()
            if current:
                self._add_to_recent(current)

    def close_word(self):
        """Close Word."""
        if self.word_widget.word_doc:
            reply = QMessageBox.question(
                self, "Close Word",
                "Do you want to save changes before closing?",
                QMessageBox.StandardButton.Save |
                QMessageBox.StandardButton.Discard |
                QMessageBox.StandardButton.Cancel
            )
            if reply == QMessageBox.StandardButton.Save:
                self.word_widget.save_document()
                self.word_widget.close_word(save=False)
            elif reply == QMessageBox.StandardButton.Discard:
                self.word_widget.close_word(save=False)
            # Cancel does nothing
        else:
            self.word_widget.close_word(save=False)

    def open_from_case(self):
        """Open a document from the current case folder."""
        if not self.main_window or not self.main_window.case_path:
            QMessageBox.warning(self, "No Case", "No case is currently loaded.")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Document from Case",
            self.main_window.case_path,
            "Word Documents (*.docx *.doc);;All Files (*.*)"
        )
        if file_path:
            if self.word_widget.open_document(file_path):
                self._add_to_recent(file_path)

    def save_to_case(self):
        """Save the document to the current case folder."""
        if not self.main_window or not self.main_window.case_path:
            QMessageBox.warning(self, "No Case", "No case is currently loaded.")
            return

        if not self.word_widget.word_doc:
            QMessageBox.warning(self, "No Document", "No document is currently open.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Document to Case",
            self.main_window.case_path,
            "Word Documents (*.docx);;All Files (*.*)"
        )
        if file_path:
            if self.word_widget.save_document_as(file_path):
                self._add_to_recent(file_path)
                QMessageBox.information(self, "Saved", f"Document saved to:\n{file_path}")

    def _add_to_recent(self, file_path):
        """Add a file to the recent files list."""
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        self.recent_files.insert(0, file_path)
        self.recent_files = self.recent_files[:10]  # Keep last 10
        self._update_recent_list()
        self._save_recent_files()

    def _update_recent_list(self):
        """Update the recent files list widget."""
        self.recent_list.clear()
        for path in self.recent_files:
            item = QListWidgetItem(os.path.basename(path))
            item.setToolTip(path)
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.recent_list.addItem(item)

    def _open_recent_file(self, item):
        """Open a file from the recent list."""
        path = item.data(Qt.ItemDataRole.UserRole)
        if os.path.exists(path):
            self.word_widget.open_document(path)
        else:
            QMessageBox.warning(self, "File Not Found", f"The file no longer exists:\n{path}")
            self.recent_files.remove(path)
            self._update_recent_list()
            self._save_recent_files()

    def _on_document_saved(self, file_path):
        """Handle document saved event."""
        self._add_to_recent(file_path)

    def _on_document_opened(self, file_path):
        """Handle document opened event."""
        self._add_to_recent(file_path)

    def load_recent_files(self):
        """Load recent files from settings."""
        config_path = os.path.join(GEMINI_DATA_DIR, "..", "config", "word_recent.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r') as f:
                    self.recent_files = json.load(f)
                self._update_recent_list()
            except:
                pass

    def _save_recent_files(self):
        """Save recent files to settings."""
        config_dir = os.path.join(GEMINI_DATA_DIR, "..", "config")
        os.makedirs(config_dir, exist_ok=True)
        config_path = os.path.join(config_dir, "word_recent.json")
        try:
            with open(config_path, 'w') as f:
                json.dump(self.recent_files, f)
        except:
            pass

    def reset_state(self):
        """Reset state when switching cases."""
        # Don't close Word when switching cases
        pass

    def closeEvent(self, event):
        """Handle tab close - ensure Word is properly closed."""
        self.word_widget.close_word(save=False)
        super().closeEvent(event)
