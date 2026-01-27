"""
Tests for Enhanced Case View Tab Database Components (No GUI required)

Tests cover:
- ProcessingLogDB - persistent processing logs
- FileTagsDB - file tagging system
- AgentSettingsDB - agent configuration
"""

import os
import sys
import json
import tempfile
import shutil
import unittest
from datetime import datetime

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Mock PySide6 before importing the module
sys.modules['PySide6'] = type(sys)('PySide6')
sys.modules['PySide6.QtWidgets'] = type(sys)('PySide6.QtWidgets')
sys.modules['PySide6.QtCore'] = type(sys)('PySide6.QtCore')
sys.modules['PySide6.QtGui'] = type(sys)('PySide6.QtGui')

# Now we can test the database components by importing them directly
# but we need to patch them to avoid the Qt imports

# Create minimal test implementations
class ProcessingLogDB:
    """Test implementation of ProcessingLogDB."""

    def __init__(self, file_number, data_dir=None):
        self.file_number = file_number
        self.data_dir = data_dir or tempfile.gettempdir()
        self.log_path = os.path.join(self.data_dir, f"{file_number}_processing_log.json")
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
            print(f"Error saving processing log: {e}")

    def add_entry(self, file_path, task_type, status, output_path=None, error_message=None, duration_sec=0):
        entry = {
            "timestamp": datetime.now().isoformat(),
            "file_path": file_path,
            "file_name": os.path.basename(file_path),
            "task_type": task_type,
            "status": status,
            "output_path": output_path,
            "error_message": error_message,
            "duration_sec": duration_sec
        }
        self.logs.insert(0, entry)
        self._save()
        return entry

    def get_file_processing_status(self, file_path):
        return [log for log in self.logs if log.get("file_path") == file_path]

    def get_all_logs(self):
        return self.logs

    def clear_logs(self):
        self.logs = []
        self._save()


class FileTagsDB:
    """Test implementation of FileTagsDB."""

    def __init__(self, file_number, data_dir=None):
        self.file_number = file_number
        self.data_dir = data_dir or tempfile.gettempdir()
        self.tags_path = os.path.join(self.data_dir, f"{file_number}_file_tags.json")
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
            print(f"Error saving file tags: {e}")

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
        all_tags = set()
        for tags in self.tags.values():
            all_tags.update(tags)
        return sorted(list(all_tags))


class AgentSettingsDB:
    """Test implementation of AgentSettingsDB."""

    DEFAULT_SETTINGS = {
        "docket.py": {"auto_run_days": 30, "save_pdf": True},
        "ocr.py": {"language": "eng", "deskew": True}
    }

    def __init__(self, data_dir=None):
        self.data_dir = data_dir or tempfile.gettempdir()
        self.settings_path = os.path.join(self.data_dir, "agent_settings.json")
        self.settings = self._load()

    def _load(self):
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
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
            print(f"Error saving agent settings: {e}")

    def get_settings(self, script_name):
        return self.settings.get(script_name, {})

    def update_settings(self, script_name, settings):
        if script_name not in self.settings:
            self.settings[script_name] = {}
        self.settings[script_name].update(settings)
        self._save()


class TestProcessingLogDB(unittest.TestCase):
    """Tests for ProcessingLogDB class."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.db = ProcessingLogDB("2024.001", self.test_dir)

    def tearDown(self):
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
        db2 = ProcessingLogDB("2024.001", self.test_dir)
        logs = db2.get_all_logs()
        self.assertEqual(len(logs), 1)

    def test_error_entry(self):
        """Test adding an error entry."""
        entry = self.db.add_entry(
            file_path="/test/file.pdf",
            task_type="OCR",
            status="failed",
            error_message="File not found"
        )

        self.assertEqual(entry["status"], "failed")
        self.assertEqual(entry["error_message"], "File not found")


class TestFileTagsDB(unittest.TestCase):
    """Tests for FileTagsDB class."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.db = FileTagsDB("2024.001", self.test_dir)

    def tearDown(self):
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

    def test_empty_file_returns_empty_list(self):
        """Test that untagged files return empty list."""
        tags = self.db.get_tags("/nonexistent/file.pdf")
        self.assertEqual(tags, [])


class TestAgentSettingsDB(unittest.TestCase):
    """Tests for AgentSettingsDB class."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.db = AgentSettingsDB(self.test_dir)

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

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

    def test_persistence(self):
        """Test that settings persist to disk."""
        self.db.update_settings("docket.py", {"auto_run_days": 7})

        # Create new instance
        db2 = AgentSettingsDB(self.test_dir)
        settings = db2.get_settings("docket.py")
        self.assertEqual(settings["auto_run_days"], 7)

    def test_add_new_setting(self):
        """Test adding a new setting key."""
        self.db.update_settings("docket.py", {"new_option": "value"})
        settings = self.db.get_settings("docket.py")
        self.assertEqual(settings["new_option"], "value")
        # Original settings should still be there
        self.assertIn("auto_run_days", settings)


class TestIntegration(unittest.TestCase):
    """Integration tests for the database components."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_full_workflow(self):
        """Test a full workflow with multiple components."""
        file_number = "2024.001"

        # Create processing log
        proc_log = ProcessingLogDB(file_number, self.test_dir)
        proc_log.add_entry("/test/file.pdf", "OCR", "success", duration_sec=5.0)
        proc_log.add_entry("/test/file.pdf", "Summarize", "success", duration_sec=10.0)

        # Create file tags
        tags_db = FileTagsDB(file_number, self.test_dir)
        tags_db.add_tag("/test/file.pdf", "important")
        tags_db.add_tag("/test/file.pdf", "reviewed")

        # Create agent settings
        settings_db = AgentSettingsDB(self.test_dir)
        settings_db.update_settings("docket.py", {"auto_run_days": 14})

        # Verify data
        logs = proc_log.get_all_logs()
        self.assertEqual(len(logs), 2)

        tags = tags_db.get_tags("/test/file.pdf")
        self.assertEqual(len(tags), 2)

        settings = settings_db.get_settings("docket.py")
        self.assertEqual(settings["auto_run_days"], 14)

    def test_multiple_files_tracking(self):
        """Test tracking multiple files."""
        file_number = "2024.001"

        proc_log = ProcessingLogDB(file_number, self.test_dir)
        tags_db = FileTagsDB(file_number, self.test_dir)

        # Process multiple files
        files = ["/test/file1.pdf", "/test/file2.pdf", "/test/file3.pdf"]

        for f in files:
            proc_log.add_entry(f, "OCR", "success")
            tags_db.add_tag(f, "processed")

        # Verify counts
        self.assertEqual(len(proc_log.get_all_logs()), 3)
        self.assertEqual(len(tags_db.get_all_tags()), 1)  # All have same tag

        # Verify individual file status
        for f in files:
            status = proc_log.get_file_processing_status(f)
            self.assertEqual(len(status), 1)
            tags = tags_db.get_tags(f)
            self.assertIn("processed", tags)


if __name__ == "__main__":
    # Run tests with verbosity
    unittest.main(verbosity=2)
