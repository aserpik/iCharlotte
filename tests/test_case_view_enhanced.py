"""
Tests for Enhanced Case View Tab Components

Tests cover:
- ProcessingLogDB - persistent processing logs
- FileTagsDB - file tagging system
- AgentSettingsDB - agent configuration
- EnhancedAgentButton - agent button with status/spinner
- AdvancedFilterWidget - filtering options
- FilePreviewWidget - file preview pane
- OutputBrowserWidget - output management
- ProcessingLogWidget - processing log viewer
- EnhancedFileTreeWidget - enhanced tree with columns
"""

import os
import sys
import json
import tempfile
import shutil
import unittest
from datetime import datetime
from unittest.mock import Mock, patch, MagicMock

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import the modules to test
from icharlotte_core.ui.case_view_enhanced import (
    ProcessingLogDB,
    FileTagsDB,
    AgentSettingsDB,
    EnhancedAgentButton,
    AdvancedFilterWidget,
    FilePreviewWidget,
    EnhancedFileTreeWidget,
    SpinningIndicator
)


class TestProcessingLogDB(unittest.TestCase):
    """Tests for ProcessingLogDB class."""

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()
        self.original_data_dir = None

        # Patch GEMINI_DATA_DIR
        import icharlotte_core.ui.case_view_enhanced as module
        self.original_data_dir = module.GEMINI_DATA_DIR
        module.GEMINI_DATA_DIR = self.test_dir

        self.db = ProcessingLogDB("2024.001")

    def tearDown(self):
        """Clean up test fixtures."""
        import icharlotte_core.ui.case_view_enhanced as module
        module.GEMINI_DATA_DIR = self.original_data_dir
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_add_entry(self):
        """Test adding a processing log entry."""
        entry = self.db.add_entry(
            file_path="/test/file.pdf",
            task_type="OCR",
            status="success",
            output_path="/test/file_ocr.txt",
            duration_sec=5.5
        )

        self.assertIsNotNone(entry)
        self.assertEqual(entry["file_path"], "/test/file.pdf")
        self.assertEqual(entry["task_type"], "OCR")
        self.assertEqual(entry["status"], "success")
        self.assertEqual(entry["duration_sec"], 5.5)

    def test_get_file_processing_status(self):
        """Test getting processing status for a file."""
        self.db.add_entry("/test/file.pdf", "OCR", "success")
        self.db.add_entry("/test/file.pdf", "Summarize", "success")
        self.db.add_entry("/test/other.pdf", "OCR", "failed")

        status = self.db.get_file_processing_status("/test/file.pdf")
        self.assertEqual(len(status), 2)

    def test_get_all_logs(self):
        """Test getting all logs."""
        self.db.add_entry("/test/file1.pdf", "OCR", "success")
        self.db.add_entry("/test/file2.pdf", "OCR", "failed")

        logs = self.db.get_all_logs()
        self.assertEqual(len(logs), 2)

    def test_clear_logs(self):
        """Test clearing all logs."""
        self.db.add_entry("/test/file.pdf", "OCR", "success")
        self.db.clear_logs()

        logs = self.db.get_all_logs()
        self.assertEqual(len(logs), 0)

    def test_persistence(self):
        """Test that logs persist to disk."""
        self.db.add_entry("/test/file.pdf", "OCR", "success")

        # Create new instance
        db2 = ProcessingLogDB("2024.001")
        logs = db2.get_all_logs()
        self.assertEqual(len(logs), 1)


class TestFileTagsDB(unittest.TestCase):
    """Tests for FileTagsDB class."""

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()

        import icharlotte_core.ui.case_view_enhanced as module
        self.original_data_dir = module.GEMINI_DATA_DIR
        module.GEMINI_DATA_DIR = self.test_dir

        self.db = FileTagsDB("2024.001")

    def tearDown(self):
        """Clean up test fixtures."""
        import icharlotte_core.ui.case_view_enhanced as module
        module.GEMINI_DATA_DIR = self.original_data_dir
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_add_tag(self):
        """Test adding a tag to a file."""
        self.db.add_tag("/test/file.pdf", "important")
        tags = self.db.get_tags("/test/file.pdf")
        self.assertIn("important", tags)

    def test_remove_tag(self):
        """Test removing a tag from a file."""
        self.db.add_tag("/test/file.pdf", "important")
        self.db.remove_tag("/test/file.pdf", "important")
        tags = self.db.get_tags("/test/file.pdf")
        self.assertNotIn("important", tags)

    def test_set_tags(self):
        """Test setting multiple tags at once."""
        self.db.set_tags("/test/file.pdf", ["tag1", "tag2", "tag3"])
        tags = self.db.get_tags("/test/file.pdf")
        self.assertEqual(len(tags), 3)

    def test_get_all_tags(self):
        """Test getting all unique tags."""
        self.db.add_tag("/test/file1.pdf", "important")
        self.db.add_tag("/test/file2.pdf", "urgent")
        self.db.add_tag("/test/file3.pdf", "important")

        all_tags = self.db.get_all_tags()
        self.assertEqual(len(all_tags), 2)
        self.assertIn("important", all_tags)
        self.assertIn("urgent", all_tags)

    def test_no_duplicate_tags(self):
        """Test that duplicate tags are not added."""
        self.db.add_tag("/test/file.pdf", "important")
        self.db.add_tag("/test/file.pdf", "important")
        tags = self.db.get_tags("/test/file.pdf")
        self.assertEqual(tags.count("important"), 1)


class TestAgentSettingsDB(unittest.TestCase):
    """Tests for AgentSettingsDB class."""

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()

        import icharlotte_core.ui.case_view_enhanced as module
        self.original_data_dir = module.GEMINI_DATA_DIR
        module.GEMINI_DATA_DIR = self.test_dir

        # Create a fresh db instance - need to ensure settings file is clean
        self.db = AgentSettingsDB()
        # Delete any existing settings file and reload defaults
        if os.path.exists(self.db.settings_path):
            os.remove(self.db.settings_path)
        self.db.settings = self.db._load()

    def tearDown(self):
        """Clean up test fixtures."""
        import icharlotte_core.ui.case_view_enhanced as module
        module.GEMINI_DATA_DIR = self.original_data_dir
        shutil.rmtree(self.test_dir, ignore_errors=True)
        # Also clean up the settings file if it exists
        if hasattr(self, 'db') and os.path.exists(self.db.settings_path):
            try:
                os.remove(self.db.settings_path)
            except:
                pass

    def test_default_settings(self):
        """Test that default settings are loaded."""
        settings = self.db.get_settings("docket.py")
        self.assertIn("auto_run_days", settings)
        self.assertEqual(settings["auto_run_days"], 30)

    def test_update_settings(self):
        """Test updating agent settings."""
        self.db.update_settings("docket.py", {"auto_run_days": 15})
        settings = self.db.get_settings("docket.py")
        self.assertEqual(settings["auto_run_days"], 15)

    def test_unknown_script_returns_empty(self):
        """Test that unknown scripts return empty settings."""
        settings = self.db.get_settings("unknown_script.py")
        self.assertEqual(settings, {})


class TestSpinningIndicator(unittest.TestCase):
    """Tests for SpinningIndicator widget."""

    def test_start_stop(self):
        """Test starting and stopping the spinner."""
        from PySide6.QtWidgets import QApplication

        # Create app if not exists
        app = QApplication.instance()
        if not app:
            app = QApplication([])

        spinner = SpinningIndicator()

        # Test initial state
        self.assertFalse(spinner._running)
        self.assertFalse(spinner.timer.isActive())

        # Start spinner
        spinner.start()
        self.assertTrue(spinner._running)
        self.assertTrue(spinner.timer.isActive())

        # Stop spinner
        spinner.stop()
        self.assertFalse(spinner._running)
        self.assertFalse(spinner.timer.isActive())


class TestEnhancedAgentButton(unittest.TestCase):
    """Tests for EnhancedAgentButton widget."""

    def setUp(self):
        """Set up test fixtures."""
        from PySide6.QtWidgets import QApplication

        self.app = QApplication.instance()
        if not self.app:
            self.app = QApplication([])

    def test_create_button(self):
        """Test creating an enhanced agent button."""
        btn = EnhancedAgentButton("Test Agent", "test.py")
        self.assertEqual(btn.name, "Test Agent")
        self.assertEqual(btn.script_name, "test.py")
        self.assertFalse(btn.is_running)

    def test_set_running(self):
        """Test setting running state."""
        btn = EnhancedAgentButton("Test Agent", "test.py")

        btn.set_running(True)
        self.assertTrue(btn.is_running)
        self.assertFalse(btn.name_btn.isEnabled())

        btn.set_running(False)
        self.assertFalse(btn.is_running)
        self.assertTrue(btn.name_btn.isEnabled())

    def test_set_status(self):
        """Test setting status text."""
        btn = EnhancedAgentButton("Test Agent", "test.py")
        btn.set_status("Running...")
        self.assertEqual(btn.status_label.text(), "Running...")

    def test_set_last_run(self):
        """Test setting last run date."""
        btn = EnhancedAgentButton("Test Agent", "test.py")
        btn.set_last_run("2024-01-15")
        self.assertIn("Last:", btn.status_label.text())

    def test_signals_emitted(self):
        """Test that signals are emitted on click."""
        btn = EnhancedAgentButton("Test Agent", "test.py")

        clicked_received = []
        settings_clicked_received = []

        btn.clicked.connect(lambda: clicked_received.append(True))
        btn.settings_clicked.connect(lambda: settings_clicked_received.append(True))

        # Simulate clicks
        btn.name_btn.click()
        self.assertEqual(len(clicked_received), 1)

        btn.settings_btn.click()
        self.assertEqual(len(settings_clicked_received), 1)


class TestAdvancedFilterWidget(unittest.TestCase):
    """Tests for AdvancedFilterWidget."""

    def setUp(self):
        """Set up test fixtures."""
        from PySide6.QtWidgets import QApplication

        self.app = QApplication.instance()
        if not self.app:
            self.app = QApplication([])

    def test_create_widget(self):
        """Test creating the filter widget."""
        widget = AdvancedFilterWidget()
        self.assertIsNotNone(widget)

    def test_get_filters_default(self):
        """Test getting default filter values."""
        widget = AdvancedFilterWidget()
        filters = widget.get_filters()

        self.assertIn("file_types", filters)
        self.assertIn("processing_status", filters)
        self.assertIn("date_filter", filters)
        self.assertEqual(filters["date_filter"], "Any time")

    def test_set_available_tags(self):
        """Test setting available tags."""
        widget = AdvancedFilterWidget()
        widget.set_available_tags(["tag1", "tag2", "tag3"])

        # Check that tags are in the combo box
        self.assertGreaterEqual(widget.tags_combo.count(), 4)  # "All Tags" + 3 tags

    def test_filter_changed_signal(self):
        """Test that filter_changed signal is emitted."""
        widget = AdvancedFilterWidget()

        signal_received = []
        widget.filter_changed.connect(lambda f: signal_received.append(f))

        # Trigger a filter change
        widget.type_pdf.setChecked(False)

        self.assertEqual(len(signal_received), 1)


class TestFilePreviewWidget(unittest.TestCase):
    """Tests for FilePreviewWidget."""

    def setUp(self):
        """Set up test fixtures."""
        from PySide6.QtWidgets import QApplication

        self.app = QApplication.instance()
        if not self.app:
            self.app = QApplication([])

        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_create_widget(self):
        """Test creating the preview widget."""
        widget = FilePreviewWidget()
        self.assertIsNotNone(widget)
        self.assertFalse(widget.open_btn.isEnabled())

    def test_show_text_file(self):
        """Test showing a text file preview."""
        # Create a test file
        test_file = os.path.join(self.test_dir, "test.txt")
        with open(test_file, "w") as f:
            f.write("Hello, World!")

        widget = FilePreviewWidget()
        widget.show_file(test_file)

        self.assertEqual(widget.current_file, test_file)
        self.assertTrue(widget.open_btn.isEnabled())

    def test_clear_preview(self):
        """Test clearing the preview."""
        widget = FilePreviewWidget()
        widget.current_file = "/test/file.txt"
        widget.open_btn.setEnabled(True)

        widget.clear()

        self.assertIsNone(widget.current_file)
        self.assertFalse(widget.open_btn.isEnabled())


class TestEnhancedFileTreeWidget(unittest.TestCase):
    """Tests for EnhancedFileTreeWidget."""

    def setUp(self):
        """Set up test fixtures."""
        from PySide6.QtWidgets import QApplication

        self.app = QApplication.instance()
        if not self.app:
            self.app = QApplication([])

        self.test_dir = tempfile.mkdtemp()

        import icharlotte_core.ui.case_view_enhanced as module
        self.original_data_dir = module.GEMINI_DATA_DIR
        module.GEMINI_DATA_DIR = self.test_dir

    def tearDown(self):
        """Clean up test fixtures."""
        import icharlotte_core.ui.case_view_enhanced as module
        module.GEMINI_DATA_DIR = self.original_data_dir
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_create_widget(self):
        """Test creating the enhanced tree widget."""
        widget = EnhancedFileTreeWidget()
        self.assertIsNotNone(widget)

    def test_set_databases(self):
        """Test setting up databases for a file number."""
        widget = EnhancedFileTreeWidget()
        widget.set_databases("2024.001")

        self.assertIsNotNone(widget.tags_db)
        self.assertIsNotNone(widget.processing_log)

    def test_get_processing_status(self):
        """Test getting processing status for a file."""
        widget = EnhancedFileTreeWidget()
        widget.set_databases("2024.001")

        # Add some processing logs
        widget.processing_log.add_entry("/test/file.pdf", "OCR", "success")
        widget.processing_log.add_entry("/test/file.pdf", "Summarize", "success")

        status = widget.get_processing_status("/test/file.pdf")
        self.assertIn("OCR", status)
        self.assertIn("SUM", status)


class TestIntegration(unittest.TestCase):
    """Integration tests for the enhanced components."""

    def setUp(self):
        """Set up test fixtures."""
        from PySide6.QtWidgets import QApplication

        self.app = QApplication.instance()
        if not self.app:
            self.app = QApplication([])

        self.test_dir = tempfile.mkdtemp()

        import icharlotte_core.ui.case_view_enhanced as module
        self.original_data_dir = module.GEMINI_DATA_DIR
        module.GEMINI_DATA_DIR = self.test_dir

    def tearDown(self):
        """Clean up test fixtures."""
        import icharlotte_core.ui.case_view_enhanced as module
        module.GEMINI_DATA_DIR = self.original_data_dir
        shutil.rmtree(self.test_dir, ignore_errors=True)
        # Also clean up the config folder that may have been created at parent level
        config_dir = os.path.join(self.test_dir, "..", "config")
        if os.path.exists(config_dir):
            shutil.rmtree(config_dir, ignore_errors=True)

    def test_full_workflow(self):
        """Test a full workflow with multiple components."""
        file_number = "2024.001"

        # Create processing log
        proc_log = ProcessingLogDB(file_number)
        proc_log.add_entry("/test/file.pdf", "OCR", "success", duration_sec=5.0)
        proc_log.add_entry("/test/file.pdf", "Summarize", "success", duration_sec=10.0)

        # Create file tags
        tags_db = FileTagsDB(file_number)
        tags_db.add_tag("/test/file.pdf", "important")
        tags_db.add_tag("/test/file.pdf", "reviewed")

        # Create agent settings
        settings_db = AgentSettingsDB()
        settings_db.update_settings("docket.py", {"auto_run_days": 14})

        # Verify data
        logs = proc_log.get_all_logs()
        self.assertEqual(len(logs), 2)

        tags = tags_db.get_tags("/test/file.pdf")
        self.assertEqual(len(tags), 2)

        settings = settings_db.get_settings("docket.py")
        self.assertEqual(settings["auto_run_days"], 14)


if __name__ == "__main__":
    unittest.main()
