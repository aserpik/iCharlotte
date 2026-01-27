"""
Chat-specific dialogs for the iCharlotte app.
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QLineEdit, QTextEdit,
    QComboBox, QMessageBox, QDialogButtonBox, QGroupBox,
    QFormLayout
)
from PySide6.QtCore import Qt

from ..chat.models import QuickPrompt, BUILTIN_PROMPTS


class PromptTemplateDialog(QDialog):
    """Dialog for managing quick prompt templates."""

    def __init__(self, persistence, parent=None):
        super().__init__(parent)
        self.persistence = persistence
        self.setWindowTitle("Manage Quick Prompts")
        self.setMinimumSize(600, 500)
        self.setup_ui()
        self.load_prompts()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Quick Prompt Templates")
        header.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(header)

        # Main content
        content_layout = QHBoxLayout()

        # Left: Prompt list
        left_layout = QVBoxLayout()

        # Category filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Category:"))
        self.category_filter = QComboBox()
        self.category_filter.addItems(["All", "Summary", "Analysis", "Extraction", "Drafting", "Custom"])
        self.category_filter.currentTextChanged.connect(self.filter_prompts)
        filter_layout.addWidget(self.category_filter)
        left_layout.addLayout(filter_layout)

        self.prompt_list = QListWidget()
        self.prompt_list.currentItemChanged.connect(self.on_prompt_selected)
        left_layout.addWidget(self.prompt_list)

        # Buttons
        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add")
        self.add_btn.clicked.connect(self.add_prompt)
        btn_layout.addWidget(self.add_btn)

        self.delete_btn = QPushButton("Delete")
        self.delete_btn.clicked.connect(self.delete_prompt)
        self.delete_btn.setEnabled(False)
        btn_layout.addWidget(self.delete_btn)

        left_layout.addLayout(btn_layout)
        content_layout.addLayout(left_layout, 1)

        # Right: Prompt editor
        right_layout = QVBoxLayout()

        editor_group = QGroupBox("Prompt Details")
        editor_layout = QFormLayout(editor_group)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Enter prompt name...")
        editor_layout.addRow("Name:", self.name_edit)

        self.category_edit = QComboBox()
        self.category_edit.addItems(["Summary", "Analysis", "Extraction", "Drafting", "Custom"])
        self.category_edit.setEditable(True)
        editor_layout.addRow("Category:", self.category_edit)

        self.prompt_edit = QTextEdit()
        self.prompt_edit.setPlaceholderText("Enter the prompt template text...")
        self.prompt_edit.setMinimumHeight(200)
        editor_layout.addRow("Prompt:", self.prompt_edit)

        right_layout.addWidget(editor_group)

        self.save_btn = QPushButton("Save Changes")
        self.save_btn.clicked.connect(self.save_prompt)
        self.save_btn.setEnabled(False)
        right_layout.addWidget(self.save_btn)

        content_layout.addLayout(right_layout, 2)
        layout.addLayout(content_layout)

        # Dialog buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(self.accept)
        layout.addWidget(button_box)

        # Track editing state
        self.current_prompt_id = None
        self.is_builtin = False

    def load_prompts(self):
        """Load all prompts into the list."""
        self.prompt_list.clear()
        self.all_prompts = []

        # Add builtin prompts
        for prompt in BUILTIN_PROMPTS:
            self.all_prompts.append(prompt)

        # Add custom prompts from persistence
        if self.persistence:
            for prompt in self.persistence.get_quick_prompts():
                if not prompt.is_builtin:
                    self.all_prompts.append(prompt)

        self.filter_prompts(self.category_filter.currentText())

    def filter_prompts(self, category: str):
        """Filter prompts by category."""
        self.prompt_list.clear()

        for prompt in self.all_prompts:
            if category == "All" or prompt.category == category:
                item = QListWidgetItem(prompt.name)
                item.setData(Qt.ItemDataRole.UserRole, prompt.id)

                # Mark builtin prompts
                if prompt.is_builtin:
                    item.setText(f"{prompt.name} (built-in)")
                    item.setForeground(Qt.GlobalColor.gray)

                self.prompt_list.addItem(item)

    def on_prompt_selected(self, item):
        """Handle prompt selection."""
        if not item:
            self.clear_editor()
            return

        prompt_id = item.data(Qt.ItemDataRole.UserRole)
        prompt = self.find_prompt(prompt_id)

        if prompt:
            self.current_prompt_id = prompt_id
            self.is_builtin = prompt.is_builtin

            self.name_edit.setText(prompt.name)
            self.category_edit.setCurrentText(prompt.category)
            self.prompt_edit.setPlainText(prompt.prompt)

            # Disable editing for builtin prompts
            self.name_edit.setEnabled(not self.is_builtin)
            self.category_edit.setEnabled(not self.is_builtin)
            self.prompt_edit.setEnabled(not self.is_builtin)
            self.save_btn.setEnabled(not self.is_builtin)
            self.delete_btn.setEnabled(not self.is_builtin)

    def find_prompt(self, prompt_id: str):
        """Find a prompt by ID."""
        for prompt in self.all_prompts:
            if prompt.id == prompt_id:
                return prompt
        return None

    def clear_editor(self):
        """Clear the editor fields."""
        self.current_prompt_id = None
        self.is_builtin = False
        self.name_edit.clear()
        self.name_edit.setEnabled(True)
        self.category_edit.setCurrentIndex(0)
        self.category_edit.setEnabled(True)
        self.prompt_edit.clear()
        self.prompt_edit.setEnabled(True)
        self.save_btn.setEnabled(False)
        self.delete_btn.setEnabled(False)

    def add_prompt(self):
        """Add a new prompt."""
        self.clear_editor()
        self.name_edit.setFocus()
        self.save_btn.setEnabled(True)
        self.current_prompt_id = None  # Will create new

    def save_prompt(self):
        """Save the current prompt."""
        name = self.name_edit.text().strip()
        category = self.category_edit.currentText().strip()
        prompt_text = self.prompt_edit.toPlainText().strip()

        if not name:
            QMessageBox.warning(self, "Error", "Please enter a prompt name.")
            return

        if not prompt_text:
            QMessageBox.warning(self, "Error", "Please enter the prompt text.")
            return

        if not self.persistence:
            QMessageBox.warning(self, "Error", "No case loaded. Cannot save prompts.")
            return

        if self.current_prompt_id and not self.is_builtin:
            # Update existing
            self.persistence.update_quick_prompt(
                self.current_prompt_id,
                name=name,
                prompt=prompt_text,
                category=category
            )
        else:
            # Create new
            self.persistence.add_quick_prompt(name, prompt_text, category)

        self.load_prompts()
        QMessageBox.information(self, "Saved", "Prompt saved successfully.")

    def delete_prompt(self):
        """Delete the selected prompt."""
        if not self.current_prompt_id or self.is_builtin:
            return

        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this prompt?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.persistence:
                self.persistence.delete_quick_prompt(self.current_prompt_id)
            self.load_prompts()
            self.clear_editor()


class ConversationSearchDialog(QDialog):
    """Dialog for searching conversations."""

    def __init__(self, persistence, parent=None):
        super().__init__(parent)
        self.persistence = persistence
        self.selected_result = None
        self.setWindowTitle("Search Conversations")
        self.setMinimumSize(500, 400)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Search input
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter search terms...")
        self.search_input.returnPressed.connect(self.do_search)
        search_layout.addWidget(self.search_input)

        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self.do_search)
        search_layout.addWidget(search_btn)

        layout.addLayout(search_layout)

        # Results list
        self.results_list = QListWidget()
        self.results_list.itemDoubleClicked.connect(self.on_result_selected)
        layout.addWidget(self.results_list)

        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept_selection)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def do_search(self):
        """Perform search."""
        query = self.search_input.text().strip()
        if not query or not self.persistence:
            return

        results = self.persistence.search_conversations(query)
        self.results_list.clear()

        for result in results:
            text = result.get('match_text', '')
            conv_name = result.get('conversation_name', '')
            result_type = result.get('type', 'message')

            if result_type == 'conversation':
                display = f"[Conversation] {conv_name}"
            else:
                role = result.get('role', 'user')
                display = f"[{conv_name}] ({role}): {text[:80]}..."

            item = QListWidgetItem(display)
            item.setData(Qt.ItemDataRole.UserRole, result)
            self.results_list.addItem(item)

    def on_result_selected(self, item):
        """Handle double-click on result."""
        self.selected_result = item.data(Qt.ItemDataRole.UserRole)
        self.accept()

    def accept_selection(self):
        """Accept the current selection."""
        item = self.results_list.currentItem()
        if item:
            self.selected_result = item.data(Qt.ItemDataRole.UserRole)
        self.accept()

    def get_selected_result(self):
        """Get the selected search result."""
        return self.selected_result
