import datetime
import json
import os

# Try to import PySide6, but provide fallback for testing without Qt
try:
    from PySide6.QtCore import QObject, Signal
    PYSIDE6_AVAILABLE = True
except ImportError:
    PYSIDE6_AVAILABLE = False
    QObject = object

# Categories that should persist across sessions
PERSISTENT_CATEGORIES = {"Calendar"}

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
