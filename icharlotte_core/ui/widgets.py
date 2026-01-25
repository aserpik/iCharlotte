import os
import re
import shutil
import subprocess
import uuid
from PyQt6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QProgressBar, QTextEdit, QSizePolicy, QMessageBox, 
    QTreeWidget, QAbstractItemView, QMenu, QApplication
)
from PyQt6.QtCore import Qt, QProcess, QObject, pyqtSignal
from PyQt6.QtGui import QTextCursor, QAction, QDragEnterEvent, QDropEvent
from ..utils import log_event

# --- Status System Classes ---

class StatusWidget(QFrame):
    cancel_requested = pyqtSignal()

    def __init__(self, agent_name, details, parent=None):
        super().__init__(parent)
        self.task_id = str(uuid.uuid4())
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shadow.Raised)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        self.output_path = None
        self.is_finished = False
        
        layout = QVBoxLayout(self)
        
        # Header Row
        header_layout = QHBoxLayout()
        self.name_label = QLabel(f"<b>{agent_name}</b>")
        self.details_label = QLabel(details)
        self.details_label.setStyleSheet("color: gray;")
        
        header_layout.addWidget(self.name_label)
        header_layout.addWidget(self.details_label)
        header_layout.addStretch()
        
        self.open_output_btn = QPushButton("Open Output")
        self.open_output_btn.setFixedSize(100, 25)
        self.open_output_btn.setVisible(False)
        self.open_output_btn.clicked.connect(self.open_output)
        header_layout.addWidget(self.open_output_btn)

        self.toggle_log_btn = QPushButton("Show Log")
        self.toggle_log_btn.setCheckable(True)
        self.toggle_log_btn.setFixedSize(80, 25)
        self.toggle_log_btn.clicked.connect(self.toggle_log)
        header_layout.addWidget(self.toggle_log_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setFixedSize(60, 25)
        self.cancel_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        self.cancel_btn.clicked.connect(self.cancel_requested.emit)
        header_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(header_layout)
        
        # Progress Bar & Status Text
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        self.status_text_label = QLabel("Initializing...")
        self.status_text_label.setStyleSheet("font-style: italic; font-size: 11px;")
        layout.addWidget(self.status_text_label)
        
        # Log Output (Hidden by default)
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFixedHeight(150)
        self.log_output.setVisible(False)
        self.log_output.setStyleSheet("font-family: Consolas; font-size: 10px; background-color: #f0f0f0;")
        layout.addWidget(self.log_output)

    def update_progress(self, percent, message):
        self.progress_bar.setValue(percent)
        self.status_text_label.setText(message)
    
    def append_log(self, text):
        self.log_output.insertPlainText(text)
        if self.log_output.isVisible():
            self.log_output.moveCursor(QTextCursor.MoveOperation.End)
    
    def set_output_file(self, path):
        self.output_path = path.strip()
        self.open_output_btn.setVisible(True)
        self.open_output_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")

    def open_output(self):
        if self.output_path and os.path.exists(self.output_path):
            try:
                os.startfile(self.output_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")
        else:
            QMessageBox.warning(self, "Error", f"File not found: {self.output_path}")

    def set_finished(self, success=True):
        self.is_finished = True
        self.cancel_btn.setVisible(False)
        if success:
            self.progress_bar.setValue(100)
            self.status_text_label.setText("Completed Successfully.")
            self.status_text_label.setStyleSheet("color: green; font-weight: bold;")
        else:
            self.status_text_label.setText("Failed / Error.")
            self.status_text_label.setStyleSheet("color: red; font-weight: bold;")
            self.toggle_log_btn.setChecked(True)
            self.toggle_log(True)

    def toggle_log(self, checked):
        self.log_output.setVisible(checked)
        self.toggle_log_btn.setText("Hide Log" if checked else "Show Log")
        if checked:
            self.log_output.moveCursor(QTextCursor.MoveOperation.End)

    def to_dict(self):
        """Serialize the widget state to a dictionary."""
        return {
            "task_id": self.task_id,
            "agent_name": self.name_label.text().replace("<b>", "").replace("</b>", ""),
            "details": self.details_label.text(),
            "progress": self.progress_bar.value(),
            "status_text": self.status_text_label.text(),
            "status_style": self.status_text_label.styleSheet(),
            "log_content": self.log_output.toPlainText(),
            "output_path": self.output_path,
            "is_finished": self.is_finished
        }

    @classmethod
    def from_dict(cls, data, parent=None):
        """Create a StatusWidget from a dictionary."""
        widget = cls(data.get("agent_name", "Unknown"), data.get("details", ""), parent)
        if "task_id" in data:
            widget.task_id = data["task_id"]
        widget.progress_bar.setValue(data.get("progress", 0))
        
        widget.is_finished = data.get("is_finished", False)
        status_text = data.get("status_text", "Restored")
        widget.status_text_label.setText(status_text)
        
        if widget.is_finished:
             widget.status_text_label.setStyleSheet(data.get("status_style", ""))

        widget.log_output.setPlainText(data.get("log_content", ""))
        
        output_path = data.get("output_path")
        if output_path:
            widget.set_output_file(output_path)
            
        return widget

class AgentRunner(QObject):
    # Signals to update UI safely
    progress_update = pyqtSignal(int, str)
    log_update = pyqtSignal(str)
    output_file_found = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(self, command, args, status_widget=None, task_id=None, file_number=None):
        super().__init__()
        self.command = command
        self.args = args
        self.status_widget = status_widget
        self.task_id = task_id
        self.file_number = file_number
        self.process = QProcess()
        
        # State tracking for reconnection
        self.log_history = []
        self.last_progress = 0
        self.last_status = "Starting..."
        self.output_file = None
        self.success = None # None=Running, True=Success, False=Fail
        
        # Connect internal signals
        self.process.readyReadStandardOutput.connect(self.handle_stdout)
        self.process.readyReadStandardError.connect(self.handle_stderr)
        self.process.finished.connect(self.handle_finished)
        
        # Connect to widget if provided
        if self.status_widget:
            self.connect_widget(self.status_widget)

    def connect_widget(self, widget):
        self.status_widget = widget
        self.progress_update.connect(self.status_widget.update_progress)
        self.log_update.connect(self.status_widget.append_log)
        self.output_file_found.connect(self.status_widget.set_output_file)
        self.finished.connect(self.status_widget.set_finished)
        
        # Handle cancellation
        self.status_widget.cancel_requested.connect(self.cancel)

    def reconnect_widget(self, widget):
        """Reconnect a restored widget to this running agent."""
        # Disconnect old widget signals if possible (or just rely on new connections)
        # Note: Previous widget is likely deleted, so signals are already disconnected by Qt.
        
        self.connect_widget(widget)
        
        # Replay history
        if self.log_history:
             widget.log_output.setPlainText("".join(self.log_history))
             widget.log_output.moveCursor(QTextCursor.MoveOperation.End)
             
        widget.update_progress(self.last_progress, self.last_status)
        
        if self.output_file:
            widget.set_output_file(self.output_file)
            
        # If already finished, sync that state
        if self.success is not None:
             widget.set_finished(self.success)

    def start(self):
        self.process.start(self.command, self.args)

    def cancel(self):
        if self.process.state() != QProcess.ProcessState.NotRunning:
            self.process.kill()
            self.log_update.emit("\n--- PROCESS CANCELLED BY USER ---")
            if self.status_widget:
                self.status_widget.update_progress(0, "Cancelled")
            # We don't call set_finished(False) here because handle_finished will likely be called
            # when the process is killed, which will call set_finished.
            # However, handle_finished checks for NormalExit. 
            # If we kill it, it might be CrashExit or similar.
            # Let's verify handle_finished logic.
    
    def handle_stdout(self):
        data = self.process.readAllStandardOutput()
        text = bytes(data).decode("utf-8", errors="ignore")
        self.log_history.append(text)
        self.log_update.emit(text)
        self.parse_progress(text)

    def handle_stderr(self):
        data = self.process.readAllStandardError()
        text = bytes(data).decode("utf-8", errors="ignore")
        self.log_history.append(text)
        self.log_update.emit(text)
        # Often stderr contains useful warnings or errors, can also parse for progress if needed

    def handle_finished(self, exit_code, exit_status):
        success = (exit_code == 0 and exit_status == QProcess.ExitStatus.NormalExit)
        self.success = success
        self.finished.emit(success)
        # Ensure process is deleted later
        self.process.deleteLater()

    def parse_progress(self, text):
        # Scan for output file paths
        patterns = [
            r"Saved summary to: (.+)",
            r"Saved variables to (.+)",
            r"saved to (.+)",
            r"Saved to: (.+)",
            r"Generated: (.+)",
            r"Report updated successfully: (.+)"
        ]
        
        for pat in patterns:
            match = re.search(pat, text, re.IGNORECASE)
            if match:
                path = match.group(1).strip().rstrip('.') 
                self.output_file_found.emit(path)
                self.output_file = path
                break

        # Simple heuristic parser
        lower = text.lower()
        new_prog = -1
        new_stat = ""
        
        if "starting" in lower:
            new_prog, new_stat = 10, "Starting..."
        elif "downloading" in lower or "scraping" in lower:
            new_prog, new_stat = 20, "Downloading data..."
        elif "extracting" in lower:
            new_prog, new_stat = 30, "Extracting text..."
        elif "scanning" in lower:
            new_prog, new_stat = 40, "Scanning files..."
        elif "querying" in lower:
            new_prog, new_stat = 50, "Querying AI Model..."
        elif "processing" in lower:
            new_prog, new_stat = 60, "Processing data..."
        elif "generating" in lower:
            new_prog, new_stat = 75, "Generating report..."
        elif "saving" in lower or "updating" in lower:
            new_prog, new_stat = 90, "Saving output..."
        elif "completed" in lower or "finished" in lower:
            new_prog, new_stat = 100, "Finishing up..."
            
        if new_prog != -1:
            self.last_progress = new_prog
            self.last_status = new_stat
            self.progress_update.emit(new_prog, new_stat)

class FileTreeWidget(QTreeWidget):
    item_moved = pyqtSignal(str, str) # old_path, new_folder_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_context_menu)

    def open_context_menu(self, position):
        item = self.itemAt(position)
        if not item: return
        
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if not path: return
        
        menu = QMenu()
        open_action = QAction("Open", self)
        open_action.triggered.connect(lambda: self.open_file(path))
        menu.addAction(open_action)
        
        reveal_action = QAction("Reveal in Explorer", self)
        reveal_action.triggered.connect(lambda: self.reveal_in_explorer(path))
        menu.addAction(reveal_action)
        
        copy_path_action = QAction("Copy Path", self)
        copy_path_action.triggered.connect(lambda: QApplication.clipboard().setText(path))
        menu.addAction(copy_path_action)
        
        menu.exec(self.viewport().mapToGlobal(position))

    def open_file(self, path):
        try:
            os.startfile(path)
        except Exception as e:
            log_event(f"Error opening file: {e}", "error")

    def reveal_in_explorer(self, path):
        try:
            subprocess.run(['explorer', '/select,', path])
        except Exception as e:
            log_event(f"Error revealing file: {e}", "error")

    def dropEvent(self, event: QDropEvent):
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
                    # Trigger refresh
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
