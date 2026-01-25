import datetime
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTextEdit, 
    QComboBox, QLabel, QPushButton
)
from PyQt6.QtCore import QObject, pyqtSignal, Qt

class _LogManager(QObject):
    log_added = pyqtSignal(str, str) # category, message

    def __init__(self):
        super().__init__()
        self.logs = {} # {category: [messages]}

    def add_log(self, category, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {message}"
        
        if category not in self.logs:
            self.logs[category] = []
        
        self.logs[category].append(formatted_msg)
        self.log_added.emit(category, formatted_msg)

    def get_logs(self, category):
        return self.logs.get(category, [])
    
    def clear_logs(self, category):
        if category in self.logs:
            self.logs[category] = []

_instance = None

def LogManager():
    global _instance
    if _instance is None:
        _instance = _LogManager()
    return _instance

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