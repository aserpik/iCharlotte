import os
import re
import shutil
import subprocess
import uuid
from PySide6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QTextEdit, QSizePolicy, QMessageBox,
    QTreeWidget, QAbstractItemView, QMenu, QApplication,
    QWidget, QGridLayout
)
from PySide6.QtCore import Qt, QProcess, QObject, Signal, QTimer
from PySide6.QtGui import QTextCursor, QAction, QDragEnterEvent, QDropEvent
from ..utils import log_event


# =============================================================================
# Pass Progress Tracking
# =============================================================================

class PassInfo:
    """Information about a processing pass."""

    def __init__(self, name: str, number: int = 0, total: int = 0):
        self.name = name
        self.number = number
        self.total = total
        self.status = "pending"  # pending, in_progress, completed, failed
        self.error_message = ""
        self.duration_sec = 0.0
        self.recoverable = True

    def __str__(self):
        if self.status == "in_progress":
            return f"{self.name} ({self.number}/{self.total})"
        elif self.status == "completed":
            return f"{self.name} - Done ({self.duration_sec:.1f}s)"
        elif self.status == "failed":
            return f"{self.name} - Failed"
        return self.name


class PassProgressWidget(QWidget):
    """Widget displaying progress for individual passes."""

    retry_requested = Signal(str)  # pass_name

    def __init__(self, parent=None):
        super().__init__(parent)
        self.passes = {}  # name -> PassInfo
        self.current_pass = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        # Header
        header_layout = QHBoxLayout()
        self.pass_label = QLabel("Pass: —")
        self.pass_label.setStyleSheet("font-weight: bold; font-size: 11px;")
        header_layout.addWidget(self.pass_label)
        header_layout.addStretch()

        self.retry_btn = QPushButton("Retry Pass")
        self.retry_btn.setFixedSize(80, 22)
        self.retry_btn.setStyleSheet("background-color: #ff9800; color: white; font-size: 10px;")
        self.retry_btn.setVisible(False)
        self.retry_btn.clicked.connect(self._on_retry_clicked)
        header_layout.addWidget(self.retry_btn)

        layout.addLayout(header_layout)

        # Pass indicators container
        self.indicators_widget = QWidget()
        self.indicators_layout = QHBoxLayout(self.indicators_widget)
        self.indicators_layout.setContentsMargins(0, 0, 0, 0)
        self.indicators_layout.setSpacing(4)
        layout.addWidget(self.indicators_widget)

        self.pass_indicators = {}  # name -> QLabel

    def set_total_passes(self, total: int, pass_names: list = None):
        """Initialize pass indicators."""
        # Clear existing
        for indicator in self.pass_indicators.values():
            indicator.deleteLater()
        self.pass_indicators.clear()
        self.passes.clear()

        names = pass_names or [f"Pass {i+1}" for i in range(total)]

        for i, name in enumerate(names):
            info = PassInfo(name, i + 1, total)
            self.passes[name] = info

            indicator = QLabel(f"○")
            indicator.setToolTip(name)
            indicator.setStyleSheet("color: gray; font-size: 12px;")
            indicator.setFixedWidth(20)
            indicator.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.indicators_layout.addWidget(indicator)
            self.pass_indicators[name] = indicator

        self.indicators_layout.addStretch()

    def pass_started(self, name: str, number: int, total: int):
        """Mark a pass as started."""
        if name not in self.passes:
            # Dynamic pass creation
            info = PassInfo(name, number, total)
            self.passes[name] = info

            indicator = QLabel("○")
            indicator.setToolTip(name)
            indicator.setStyleSheet("color: gray; font-size: 12px;")
            indicator.setFixedWidth(20)
            indicator.setAlignment(Qt.AlignmentFlag.AlignCenter)

            # Insert before stretch
            count = self.indicators_layout.count()
            self.indicators_layout.insertWidget(count - 1 if count > 0 else 0, indicator)
            self.pass_indicators[name] = indicator

        self.passes[name].status = "in_progress"
        self.passes[name].number = number
        self.passes[name].total = total
        self.current_pass = name

        self._update_indicator(name, "in_progress")
        self.pass_label.setText(f"Pass {number}/{total}: {name}")
        self.retry_btn.setVisible(False)

    def pass_completed(self, name: str, duration_sec: float = 0.0):
        """Mark a pass as completed."""
        if name in self.passes:
            self.passes[name].status = "completed"
            self.passes[name].duration_sec = duration_sec
            self._update_indicator(name, "completed")

            # Update label
            info = self.passes[name]
            self.pass_label.setText(f"Pass {info.number}/{info.total}: {name} - Done")

    def pass_failed(self, name: str, error: str, recoverable: bool = True):
        """Mark a pass as failed."""
        if name in self.passes:
            self.passes[name].status = "failed"
            self.passes[name].error_message = error
            self.passes[name].recoverable = recoverable
            self._update_indicator(name, "failed")

            # Show retry button for recoverable failures
            if recoverable:
                self.retry_btn.setVisible(True)

            # Update label
            info = self.passes[name]
            self.pass_label.setText(f"Pass {info.number}/{info.total}: {name} - FAILED")
            self.pass_label.setStyleSheet("font-weight: bold; font-size: 11px; color: red;")

    def _update_indicator(self, name: str, status: str):
        """Update the visual indicator for a pass."""
        if name not in self.pass_indicators:
            return

        indicator = self.pass_indicators[name]

        if status == "pending":
            indicator.setText("○")
            indicator.setStyleSheet("color: gray; font-size: 12px;")
        elif status == "in_progress":
            indicator.setText("◉")
            indicator.setStyleSheet("color: #2196F3; font-size: 12px;")  # Blue
        elif status == "completed":
            indicator.setText("●")
            indicator.setStyleSheet("color: #4CAF50; font-size: 12px;")  # Green
        elif status == "failed":
            indicator.setText("✗")
            indicator.setStyleSheet("color: #f44336; font-size: 12px;")  # Red

    def _on_retry_clicked(self):
        """Handle retry button click."""
        if self.current_pass:
            self.retry_requested.emit(self.current_pass)

    def reset(self):
        """Reset all passes to pending state."""
        for name in self.passes:
            self.passes[name].status = "pending"
            self._update_indicator(name, "pending")
        self.current_pass = None
        self.pass_label.setText("Pass: —")
        self.pass_label.setStyleSheet("font-weight: bold; font-size: 11px;")
        self.retry_btn.setVisible(False)


# =============================================================================
# Status System Classes
# =============================================================================

class StatusWidget(QFrame):
    cancel_requested = Signal()
    retry_pass_requested = Signal(str)  # pass_name

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

        # Pass Progress Widget (NEW)
        self.pass_progress = PassProgressWidget()
        self.pass_progress.retry_requested.connect(self.retry_pass_requested.emit)
        layout.addWidget(self.pass_progress)

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

    def update_pass_start(self, name: str, number: int, total: int):
        """Handle pass start event."""
        self.pass_progress.pass_started(name, number, total)

    def update_pass_complete(self, name: str, duration_sec: float = 0.0):
        """Handle pass complete event."""
        self.pass_progress.pass_completed(name, duration_sec)

    def update_pass_failed(self, name: str, error: str, recoverable: bool = True):
        """Handle pass failed event."""
        self.pass_progress.pass_failed(name, error, recoverable)

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
        # Serialize pass information
        passes_data = {}
        for name, info in self.pass_progress.passes.items():
            passes_data[name] = {
                "number": info.number,
                "total": info.total,
                "status": info.status,
                "error_message": info.error_message,
                "duration_sec": info.duration_sec,
                "recoverable": info.recoverable
            }

        return {
            "task_id": self.task_id,
            "agent_name": self.name_label.text().replace("<b>", "").replace("</b>", ""),
            "details": self.details_label.text(),
            "progress": self.progress_bar.value(),
            "status_text": self.status_text_label.text(),
            "status_style": self.status_text_label.styleSheet(),
            "log_content": self.log_output.toPlainText(),
            "output_path": self.output_path,
            "is_finished": self.is_finished,
            "passes": passes_data,
            "current_pass": self.pass_progress.current_pass
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

        # Restore pass information
        passes_data = data.get("passes", {})
        if passes_data:
            widget.pass_progress.set_total_passes(
                len(passes_data),
                list(passes_data.keys())
            )
            for name, info_data in passes_data.items():
                info = widget.pass_progress.passes.get(name)
                if info:
                    info.number = info_data.get("number", 0)
                    info.total = info_data.get("total", 0)
                    info.status = info_data.get("status", "pending")
                    info.error_message = info_data.get("error_message", "")
                    info.duration_sec = info_data.get("duration_sec", 0.0)
                    info.recoverable = info_data.get("recoverable", True)
                    widget.pass_progress._update_indicator(name, info.status)

            current = data.get("current_pass")
            if current:
                widget.pass_progress.current_pass = current
                if current in widget.pass_progress.passes:
                    info = widget.pass_progress.passes[current]
                    widget.pass_progress.pass_label.setText(
                        f"Pass {info.number}/{info.total}: {current}"
                    )

        return widget


class AgentRunner(QObject):
    # Signals to update UI safely
    progress_update = Signal(int, str)
    log_update = Signal(str)
    output_file_found = Signal(str)
    finished = Signal(bool)

    # Pass-level signals (NEW)
    pass_started = Signal(str, int, int)  # name, number, total
    pass_completed = Signal(str, float)   # name, duration_sec
    pass_failed = Signal(str, str, bool)  # name, error, recoverable

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

        # Pass tracking (NEW)
        self.current_pass = None
        self.pass_info = {}  # name -> {number, total, status}

        # Connect internal signals
        self.process.readyReadStandardOutput.connect(self.handle_stdout)
        self.process.readyReadStandardError.connect(self.handle_stderr)
        self.process.finished.connect(self.handle_finished)

        # Connect to widget if provided
        if self.status_widget:
            self.connect_widget(self.status_widget)

    def disconnect_widget(self):
        """Disconnect signals from the current widget before connecting a new one."""
        if self.status_widget is None:
            return
        try:
            self.progress_update.disconnect(self.status_widget.update_progress)
            self.log_update.disconnect(self.status_widget.append_log)
            self.output_file_found.disconnect(self.status_widget.set_output_file)
            self.finished.disconnect(self.status_widget.set_finished)
            self.pass_started.disconnect(self.status_widget.update_pass_start)
            self.pass_completed.disconnect(self.status_widget.update_pass_complete)
            self.pass_failed.disconnect(self.status_widget.update_pass_failed)
            self.status_widget.cancel_requested.disconnect(self.cancel)
            self.status_widget.retry_pass_requested.disconnect(self.retry_pass)
        except (TypeError, RuntimeError):
            # Signals may not be connected or widget may already be deleted
            pass
        self.status_widget = None

    def connect_widget(self, widget):
        # Disconnect from old widget first to prevent signals to deleted widgets
        self.disconnect_widget()

        self.status_widget = widget
        self.progress_update.connect(self.status_widget.update_progress)
        self.log_update.connect(self.status_widget.append_log)
        self.output_file_found.connect(self.status_widget.set_output_file)
        self.finished.connect(self.status_widget.set_finished)

        # Connect pass-level signals (NEW)
        self.pass_started.connect(self.status_widget.update_pass_start)
        self.pass_completed.connect(self.status_widget.update_pass_complete)
        self.pass_failed.connect(self.status_widget.update_pass_failed)

        # Handle cancellation
        self.status_widget.cancel_requested.connect(self.cancel)

        # Handle pass retry
        self.status_widget.retry_pass_requested.connect(self.retry_pass)

    def reconnect_widget(self, widget):
        """Reconnect a restored widget to this running agent."""
        self.connect_widget(widget)

        # Replay history
        if self.log_history:
             widget.log_output.setPlainText("".join(self.log_history))
             widget.log_output.moveCursor(QTextCursor.MoveOperation.End)

        widget.update_progress(self.last_progress, self.last_status)

        if self.output_file:
            widget.set_output_file(self.output_file)

        # Replay pass info
        for name, info in self.pass_info.items():
            if info["status"] == "started":
                widget.update_pass_start(name, info["number"], info["total"])
            elif info["status"] == "completed":
                widget.update_pass_start(name, info["number"], info["total"])
                widget.update_pass_complete(name, info.get("duration", 0.0))
            elif info["status"] == "failed":
                widget.update_pass_start(name, info["number"], info["total"])
                widget.update_pass_failed(name, info.get("error", ""), info.get("recoverable", True))

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

    def retry_pass(self, pass_name: str):
        """Retry a failed pass. (Placeholder for future implementation)"""
        # This would require caching intermediate results and re-running
        # For now, just log the request
        self.log_update.emit(f"\n--- RETRY REQUESTED FOR PASS: {pass_name} ---\n")
        # Future: Restart process with --retry-pass argument

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

    def handle_finished(self, exit_code, exit_status):
        success = (exit_code == 0 and exit_status == QProcess.ExitStatus.NormalExit)
        self.success = success
        self.finished.emit(success)
        self.process.deleteLater()

    def parse_progress(self, text):
        """Parse agent output for progress and pass information."""
        lines = text.split('\n')

        for line in lines:
            line = line.strip()

            # Parse structured pass messages (NEW)
            # Format: PASS_START:name:current:total
            if line.startswith("PASS_START:"):
                parts = line.split(":")
                if len(parts) >= 4:
                    name = parts[1]
                    try:
                        current = int(parts[2])
                        total = int(parts[3])
                        self.current_pass = name
                        self.pass_info[name] = {
                            "number": current,
                            "total": total,
                            "status": "started"
                        }
                        self.pass_started.emit(name, current, total)
                    except ValueError:
                        pass
                continue

            # Format: PASS_COMPLETE:name:status:duration
            if line.startswith("PASS_COMPLETE:"):
                parts = line.split(":")
                if len(parts) >= 3:
                    name = parts[1]
                    status = parts[2] if len(parts) > 2 else "success"
                    duration = 0.0
                    if len(parts) > 3:
                        try:
                            duration = float(parts[3])
                        except ValueError:
                            pass

                    if name in self.pass_info:
                        self.pass_info[name]["status"] = "completed"
                        self.pass_info[name]["duration"] = duration
                    self.pass_completed.emit(name, duration)
                continue

            # Format: PASS_FAILED:name:error:recoverable
            if line.startswith("PASS_FAILED:"):
                parts = line.split(":")
                if len(parts) >= 3:
                    name = parts[1]
                    error = parts[2] if len(parts) > 2 else "Unknown error"
                    recoverable = True
                    if len(parts) > 3:
                        recoverable = parts[3].lower() in ("true", "1", "yes", "recoverable")

                    if name in self.pass_info:
                        self.pass_info[name]["status"] = "failed"
                        self.pass_info[name]["error"] = error
                        self.pass_info[name]["recoverable"] = recoverable
                    self.pass_failed.emit(name, error, recoverable)
                continue

            # Format: PROGRESS:percent:message
            if line.startswith("PROGRESS:"):
                parts = line.split(":", 2)
                if len(parts) >= 3:
                    try:
                        percent = int(parts[1])
                        message = parts[2]
                        self.last_progress = percent
                        self.last_status = message
                        self.progress_update.emit(percent, message)
                    except ValueError:
                        pass
                continue

            # Format: OUTPUT_FILE:path
            if line.startswith("OUTPUT_FILE:"):
                path = line[12:].strip()
                if path:
                    self.output_file = path
                    self.output_file_found.emit(path)
                continue

        # Scan for output file paths (legacy patterns)
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

        # Simple heuristic parser (fallback for agents without structured output)
        lower = text.lower()
        new_prog = -1
        new_stat = ""

        # Only use heuristics if no structured progress received
        if not any(line.strip().startswith(("PASS_", "PROGRESS:")) for line in text.split('\n')):
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
    item_moved = Signal(str, str) # old_path, new_folder_path

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
