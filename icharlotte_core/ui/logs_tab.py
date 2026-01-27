import datetime
import json
import os

# Try to import PySide6, but provide fallback for testing without Qt
try:
    from PySide6.QtWidgets import (
        QWidget, QVBoxLayout, QHBoxLayout, QTextEdit,
        QComboBox, QLabel, QPushButton
    )
    from PySide6.QtCore import QObject, Signal, Qt
    PYSIDE6_AVAILABLE = True
except ImportError:
    PYSIDE6_AVAILABLE = False
    QObject = object
    QWidget = object

# Categories that should persist across sessions
PERSISTENT_CATEGORIES = {"Email Monitor", "Calendar"}

# Log file path
LOGS_DIR = os.path.join(os.getcwd(), ".gemini", "logs")
PERSISTENT_LOGS_FILE = os.path.join(LOGS_DIR, "persistent_logs.json")


class _LogManager(QObject if PYSIDE6_AVAILABLE else object):
    """Log manager that works with or without Qt/PySide6."""

    if PYSIDE6_AVAILABLE:
        log_added = Signal(str, str)  # category, message

    def __init__(self):
        if PYSIDE6_AVAILABLE:
            super().__init__()
        self.logs = {}  # {category: [messages]}
        self._callbacks = []  # For non-Qt callbacks
        self._load_persistent_logs()

    def _load_persistent_logs(self):
        """Load persistent logs from file."""
        if os.path.exists(PERSISTENT_LOGS_FILE):
            try:
                with open(PERSISTENT_LOGS_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for category in PERSISTENT_CATEGORIES:
                        if category in data:
                            self.logs[category] = data[category]
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Could not load persistent logs: {e}")

    def _save_persistent_logs(self):
        """Save persistent category logs to file."""
        try:
            os.makedirs(LOGS_DIR, exist_ok=True)
            data = {cat: self.logs.get(cat, []) for cat in PERSISTENT_CATEGORIES}
            with open(PERSISTENT_LOGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except IOError as e:
            print(f"Warning: Could not save persistent logs: {e}")

    def add_log(self, category, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_msg = f"[{timestamp}] {message}"

        if category not in self.logs:
            self.logs[category] = []

        self.logs[category].append(formatted_msg)

        # Save if this is a persistent category
        if category in PERSISTENT_CATEGORIES:
            self._save_persistent_logs()

        if PYSIDE6_AVAILABLE:
            self.log_added.emit(category, formatted_msg)
        else:
            # Non-Qt fallback: call registered callbacks
            for callback in self._callbacks:
                try:
                    callback(category, formatted_msg)
                except Exception:
                    pass

    def add_callback(self, callback):
        """Register a callback for log events (non-Qt environments)."""
        self._callbacks.append(callback)

    def get_logs(self, category):
        return self.logs.get(category, [])

    def clear_logs(self, category):
        if category in self.logs:
            self.logs[category] = []
        # Also clear from persistent storage if applicable
        if category in PERSISTENT_CATEGORIES:
            self._save_persistent_logs()

_instance = None

def LogManager():
    global _instance
    if _instance is None:
        _instance = _LogManager()
    return _instance

# LogsTab is only available when PySide6 is installed
if PYSIDE6_AVAILABLE:
    class LogsTab(QWidget):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.log_manager = LogManager()
            self.setup_ui()
            self.log_manager.log_added.connect(self.on_log_added)

        def setup_ui(self):
            layout = QVBoxLayout(self)

            # Controls
            controls_layout = QHBoxLayout()
            controls_layout.addWidget(QLabel("Select Category:"))

            self.category_combo = QComboBox()
            self.category_combo.addItem("All")
            self.category_combo.addItem("System")
            self.category_combo.addItem("Case View")
            self.category_combo.addItem("Status")
            self.category_combo.addItem("Index")
            self.category_combo.addItem("Note Taker")
            self.category_combo.addItem("Report Tab")
            self.category_combo.addItem("Email Tab")
            self.category_combo.addItem("Chat Tab")
            self.category_combo.addItem("Email Monitor")
            self.category_combo.addItem("Calendar")
            self.category_combo.currentTextChanged.connect(self.refresh_view)
            controls_layout.addWidget(self.category_combo)

            controls_layout.addStretch()

            clear_btn = QPushButton("Clear View")
            clear_btn.clicked.connect(self.clear_current_view)
            controls_layout.addWidget(clear_btn)

            layout.addLayout(controls_layout)

            # Log View
            self.text_edit = QTextEdit()
            self.text_edit.setReadOnly(True)
            self.text_edit.setStyleSheet("font-family: Consolas; font-size: 11px;")
            layout.addWidget(self.text_edit)

        def refresh_view(self):
            self.text_edit.clear()
            cat = self.category_combo.currentText()

            if cat == "All":
                # Combine all (sorted by time might be hard if stored separately,
                # but for now just dumping all categories)
                all_logs = []
                for c, msgs in self.log_manager.logs.items():
                    for m in msgs:
                        all_logs.append(f"[{c}] {m}")
                # Naive sort by string timestamp might fail if cross-day,
                # but acceptable for debugging session
                all_logs.sort()
                self.text_edit.setPlainText("\n".join(all_logs))
            else:
                logs = self.log_manager.get_logs(cat)
                self.text_edit.setPlainText("\n".join(logs))

            self.text_edit.moveCursor(self.text_edit.textCursor().MoveOperation.End)

        def on_log_added(self, category, message):
            current_cat = self.category_combo.currentText()
            if current_cat == "All":
                self.text_edit.append(f"[{category}] {message}")
            elif current_cat == category:
                self.text_edit.append(message)

        def clear_current_view(self):
            self.text_edit.clear()
else:
    # Placeholder when PySide6 is not available
    LogsTab = None