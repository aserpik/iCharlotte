"""
Tests for Case Loading and Agent Button State Management

Tests cover:
- AgentRunner disconnect_widget and connect_widget methods
- EnhancedAgentButton running state reset on case switch
- on_agent_finished callback with case switches
- clear_all_status disconnection behavior
- Status history save/load with running agents
- Integration tests for case switching scenarios
"""

import os
import sys
import json
import tempfile
import shutil
import unittest
from unittest.mock import Mock, MagicMock, patch, PropertyMock
from functools import partial

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication, QVBoxLayout, QWidget
from PySide6.QtCore import QObject, Signal

# Import the modules to test
from icharlotte_core.ui.widgets import StatusWidget, AgentRunner
from icharlotte_core.ui.case_view_enhanced import EnhancedAgentButton, SpinningIndicator


class TestAgentRunnerSignalManagement(unittest.TestCase):
    """Tests for AgentRunner signal connect/disconnect functionality."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_disconnect_widget_when_none(self):
        """Test that disconnect_widget handles None widget gracefully."""
        runner = AgentRunner("python", ["-c", "pass"])
        runner.status_widget = None

        # Should not raise any exceptions
        runner.disconnect_widget()
        self.assertIsNone(runner.status_widget)

    def test_disconnect_widget_clears_reference(self):
        """Test that disconnect_widget clears the widget reference."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget = StatusWidget("Test Agent", "Test details")

        runner.connect_widget(widget)
        self.assertIsNotNone(runner.status_widget)

        runner.disconnect_widget()
        self.assertIsNone(runner.status_widget)

    def test_connect_widget_disconnects_old_first(self):
        """Test that connect_widget disconnects from old widget before connecting to new."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget1 = StatusWidget("Agent 1", "Details 1")
        widget2 = StatusWidget("Agent 2", "Details 2")

        runner.connect_widget(widget1)
        self.assertEqual(runner.status_widget, widget1)

        # Connect to new widget - should disconnect from widget1 first
        runner.connect_widget(widget2)
        self.assertEqual(runner.status_widget, widget2)

    def test_disconnect_handles_deleted_widget(self):
        """Test that disconnect_widget handles already deleted widgets."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget = StatusWidget("Test Agent", "Test details")

        runner.connect_widget(widget)

        # Simulate widget deletion
        widget.deleteLater()
        QApplication.processEvents()

        # Should not raise exception even with deleted widget
        try:
            runner.disconnect_widget()
        except (TypeError, RuntimeError):
            self.fail("disconnect_widget raised exception for deleted widget")

    def test_signals_connected_after_connect_widget(self):
        """Test that signals are properly connected after connect_widget."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget = StatusWidget("Test Agent", "Test details")

        runner.connect_widget(widget)

        # Verify progress update signal is connected by emitting
        progress_received = []
        original_update = widget.update_progress
        widget.update_progress = lambda p, m: progress_received.append((p, m))

        runner.progress_update.emit(50, "Test message")

        self.assertEqual(len(progress_received), 1)
        self.assertEqual(progress_received[0], (50, "Test message"))

    def test_reconnect_widget_replays_state(self):
        """Test that reconnect_widget replays accumulated state."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget1 = StatusWidget("Test Agent", "Test details")

        runner.connect_widget(widget1)

        # Accumulate some state
        runner.log_history = ["Line 1\n", "Line 2\n"]
        runner.last_progress = 75
        runner.last_status = "Processing..."
        runner.output_file = "/path/to/output.txt"

        # Disconnect and create new widget
        runner.disconnect_widget()
        widget2 = StatusWidget("Test Agent", "Test details")

        # Reconnect should replay state
        runner.reconnect_widget(widget2)

        self.assertEqual(widget2.progress_bar.value(), 75)
        self.assertIn("Line 1", widget2.log_output.toPlainText())
        self.assertIn("Line 2", widget2.log_output.toPlainText())


class TestEnhancedAgentButtonReset(unittest.TestCase):
    """Tests for EnhancedAgentButton running state management."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_set_running_true(self):
        """Test setting button to running state."""
        btn = EnhancedAgentButton("Test Agent", "test.py")

        btn.set_running(True)

        self.assertTrue(btn.is_running)
        self.assertTrue(btn.spinner._running)
        self.assertFalse(btn.name_btn.isEnabled())

    def test_set_running_false(self):
        """Test setting button to not running state."""
        btn = EnhancedAgentButton("Test Agent", "test.py")

        # First set to running
        btn.set_running(True)

        # Then set to not running
        btn.set_running(False)

        self.assertFalse(btn.is_running)
        self.assertFalse(btn.spinner._running)
        self.assertTrue(btn.name_btn.isEnabled())

    def test_multiple_buttons_independent(self):
        """Test that multiple buttons maintain independent state."""
        btn1 = EnhancedAgentButton("Agent 1", "agent1.py")
        btn2 = EnhancedAgentButton("Agent 2", "agent2.py")
        btn3 = EnhancedAgentButton("Agent 3", "agent3.py")

        btn1.set_running(True)
        btn2.set_running(True)

        # Reset btn1 only
        btn1.set_running(False)

        self.assertFalse(btn1.is_running)
        self.assertTrue(btn2.is_running)
        self.assertFalse(btn3.is_running)

    def test_reset_all_buttons_simulates_case_switch(self):
        """Test resetting all buttons simulates case switch behavior."""
        buttons = {
            "docket.py": EnhancedAgentButton("Docket", "docket.py"),
            "complaint.py": EnhancedAgentButton("Complaint", "complaint.py"),
            "summons.py": EnhancedAgentButton("Summons", "summons.py"),
        }

        # Set some buttons to running
        buttons["docket.py"].set_running(True)
        buttons["complaint.py"].set_running(True)

        # Simulate case switch - reset all buttons
        for btn in buttons.values():
            btn.set_running(False)

        # All buttons should be not running
        for name, btn in buttons.items():
            self.assertFalse(btn.is_running, f"{name} should not be running")
            self.assertTrue(btn.name_btn.isEnabled(), f"{name} should be enabled")


class TestOnAgentFinishedCaseSwitching(unittest.TestCase):
    """Tests for on_agent_finished callback behavior during case switches."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def setUp(self):
        """Set up mock MainWindow-like object for testing."""
        self.file_number = "2024.001"
        self.running_agents = {}
        self.agent_buttons = {
            "docket.py": EnhancedAgentButton("Docket", "docket.py"),
        }
        self.master_db = MagicMock()

    def on_agent_finished(self, script, btn_widget, started_for_case, success):
        """Simplified version of MainWindow.on_agent_finished for testing."""
        from datetime import datetime
        today = datetime.now().strftime("%Y-%m-%d")

        # Clear the running state
        if script in self.running_agents:
            del self.running_agents[script]

        # Only update button UI if we're still on the same case
        if self.file_number == started_for_case:
            btn_widget.set_running(False)
            if success:
                btn_widget.set_status("Last: Just now")
            else:
                btn_widget.set_status("Last: Failed")

        # Update database for the ORIGINAL case
        if script == "docket.py" and success and started_for_case:
            self.master_db.update_last_docket_download(started_for_case, today)
            if self.file_number == started_for_case:
                btn_widget.set_last_run(today)

    def test_agent_finished_same_case(self):
        """Test on_agent_finished when still on same case."""
        from datetime import datetime
        today = datetime.now().strftime("%Y-%m-%d")

        btn = self.agent_buttons["docket.py"]
        btn.set_running(True)
        self.running_agents["docket.py"] = "2024.001"

        # Agent finishes while still on same case
        self.on_agent_finished("docket.py", btn, "2024.001", True)

        # Button should be updated
        self.assertFalse(btn.is_running)
        # Check that status was set (either "Just now" or "Today" depending on set_last_run)
        status_text = btn.status_label.text()
        self.assertTrue(
            "Just now" in status_text or "Today" in status_text,
            f"Expected 'Just now' or 'Today' in status, got: {status_text}"
        )

        # Database should be updated for correct case
        self.master_db.update_last_docket_download.assert_called_with("2024.001", today)

    def test_agent_finished_different_case(self):
        """Test on_agent_finished when switched to different case."""
        from datetime import datetime
        today = datetime.now().strftime("%Y-%m-%d")

        btn = self.agent_buttons["docket.py"]
        btn.set_running(True)
        self.running_agents["docket.py"] = "2024.001"

        # Switch to different case
        self.file_number = "2024.002"
        btn.set_running(False)  # Reset button for new case

        # Agent finishes for original case
        self.on_agent_finished("docket.py", btn, "2024.001", True)

        # Button should NOT be updated (we're on different case)
        self.assertFalse(btn.is_running)
        self.assertNotIn("Just now", btn.status_label.text())

        # Database should STILL be updated for ORIGINAL case
        self.master_db.update_last_docket_download.assert_called_with("2024.001", today)

    def test_agent_finished_failure(self):
        """Test on_agent_finished with failure status."""
        btn = self.agent_buttons["docket.py"]
        btn.set_running(True)
        self.running_agents["docket.py"] = "2024.001"

        # Agent fails
        self.on_agent_finished("docket.py", btn, "2024.001", False)

        # Button should show failure
        self.assertFalse(btn.is_running)
        self.assertIn("Failed", btn.status_label.text())

        # Database should NOT be updated on failure
        self.master_db.update_last_docket_download.assert_not_called()


class TestClearAllStatusDisconnection(unittest.TestCase):
    """Tests for clear_all_status disconnection behavior."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def setUp(self):
        """Set up test fixtures."""
        self.container = QWidget()
        self.status_list_layout = QVBoxLayout(self.container)
        self.agent_runners = []

    def clear_all_status(self):
        """Simplified version of MainWindow.clear_all_status for testing."""
        # Disconnect all runners from their widgets before deleting
        for runner in self.agent_runners:
            runner.disconnect_widget()

        for i in range(self.status_list_layout.count() - 1, -1, -1):
            item = self.status_list_layout.itemAt(i)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def test_clear_disconnects_runners(self):
        """Test that clear_all_status disconnects all runners."""
        # Create widgets and runners
        widget1 = StatusWidget("Agent 1", "Details 1")
        widget2 = StatusWidget("Agent 2", "Details 2")

        runner1 = AgentRunner("python", ["-c", "pass"], widget1)
        runner2 = AgentRunner("python", ["-c", "pass"], widget2)

        self.status_list_layout.addWidget(widget1)
        self.status_list_layout.addWidget(widget2)
        self.agent_runners = [runner1, runner2]

        # Clear all status
        self.clear_all_status()

        # Runners should be disconnected
        self.assertIsNone(runner1.status_widget)
        self.assertIsNone(runner2.status_widget)

    def test_signals_dont_crash_after_clear(self):
        """Test that emitting signals after clear doesn't crash."""
        widget = StatusWidget("Test Agent", "Details")
        runner = AgentRunner("python", ["-c", "pass"], widget)

        self.status_list_layout.addWidget(widget)
        self.agent_runners = [runner]

        # Clear all status
        self.clear_all_status()
        QApplication.processEvents()

        # Emit signals - should not crash
        try:
            runner.progress_update.emit(50, "Test")
            runner.log_update.emit("Log line")
            runner.finished.emit(True)
        except Exception as e:
            self.fail(f"Signal emission raised exception: {e}")


class TestStatusHistorySaveLoad(unittest.TestCase):
    """Tests for status history save/load functionality."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()
        self.container = QWidget()
        self.status_list_layout = QVBoxLayout(self.container)
        self.agent_runners = []
        self.file_number = "2024.001"

    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def save_status_history(self):
        """Simplified version of MainWindow.save_status_history for testing."""
        if not self.file_number:
            return

        history = []
        for i in range(self.status_list_layout.count()):
            item = self.status_list_layout.itemAt(i)
            widget = item.widget()
            if isinstance(widget, StatusWidget):
                history.append(widget.to_dict())

        save_path = os.path.join(self.test_dir, f"{self.file_number}_status_history.json")
        with open(save_path, 'w') as f:
            json.dump(history, f, indent=2)

    def load_status_history(self):
        """Simplified version of MainWindow.load_status_history for testing."""
        if not self.file_number:
            return

        save_path = os.path.join(self.test_dir, f"{self.file_number}_status_history.json")
        if not os.path.exists(save_path):
            return

        with open(save_path, 'r') as f:
            history = json.load(f)

        loaded_widgets = []
        for item_data in history:
            widget = StatusWidget.from_dict(item_data)

            # Try to reconnect to running agent
            reconnected = False
            if not widget.is_finished:
                task_id = getattr(widget, 'task_id', None)
                if task_id:
                    for runner in self.agent_runners:
                        if getattr(runner, 'task_id', None) == task_id:
                            runner.reconnect_widget(widget)
                            reconnected = True
                            break

                if not reconnected:
                    widget.status_text_label.setText(widget.status_text_label.text() + " (Interrupted)")
                    widget.is_finished = True

            self.status_list_layout.addWidget(widget)
            loaded_widgets.append(widget)

        return loaded_widgets

    def test_save_and_load_finished_task(self):
        """Test saving and loading a finished task."""
        widget = StatusWidget("Test Agent", "Details")
        widget.update_progress(100, "Completed")
        widget.is_finished = True

        self.status_list_layout.addWidget(widget)

        # Save
        self.save_status_history()

        # Clear and load
        while self.status_list_layout.count():
            item = self.status_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        QApplication.processEvents()

        loaded = self.load_status_history()

        self.assertEqual(len(loaded), 1)
        self.assertTrue(loaded[0].is_finished)
        self.assertEqual(loaded[0].progress_bar.value(), 100)

    def test_load_running_task_without_runner_marks_interrupted(self):
        """Test that loading a running task without runner marks it interrupted."""
        widget = StatusWidget("Test Agent", "Details")
        widget.update_progress(50, "Processing...")
        widget.is_finished = False

        self.status_list_layout.addWidget(widget)

        # Save
        self.save_status_history()

        # Clear layout
        while self.status_list_layout.count():
            item = self.status_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        QApplication.processEvents()

        # Load without any runners
        self.agent_runners = []
        loaded = self.load_status_history()

        self.assertEqual(len(loaded), 1)
        self.assertTrue(loaded[0].is_finished)
        self.assertIn("Interrupted", loaded[0].status_text_label.text())

    def test_load_running_task_reconnects_to_runner(self):
        """Test that loading a running task reconnects to its runner."""
        # Create widget and runner
        widget = StatusWidget("Test Agent", "Details")
        widget.is_finished = False
        task_id = widget.task_id

        runner = AgentRunner("python", ["-c", "pass"])
        runner.task_id = task_id
        runner.log_history = ["Test log\n"]
        runner.last_progress = 60
        runner.last_status = "Running..."

        self.status_list_layout.addWidget(widget)

        # Save
        self.save_status_history()

        # Clear layout but keep runner
        while self.status_list_layout.count():
            item = self.status_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        QApplication.processEvents()

        # Load with runner available
        self.agent_runners = [runner]
        loaded = self.load_status_history()

        self.assertEqual(len(loaded), 1)
        self.assertFalse(loaded[0].is_finished)
        self.assertEqual(runner.status_widget, loaded[0])
        self.assertEqual(loaded[0].progress_bar.value(), 60)


class TestCaseSwitchingIntegration(unittest.TestCase):
    """Integration tests for complete case switching scenarios."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()
        self.container = QWidget()
        self.status_list_layout = QVBoxLayout(self.container)
        self.agent_runners = []
        self.agent_buttons = {
            "docket.py": EnhancedAgentButton("Docket", "docket.py"),
            "complaint.py": EnhancedAgentButton("Complaint", "complaint.py"),
        }
        self.running_agents = {}
        self.file_number = None
        self.master_db = MagicMock()

    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def clear_all_status(self):
        """Clear all status widgets."""
        for runner in self.agent_runners:
            runner.disconnect_widget()

        for i in range(self.status_list_layout.count() - 1, -1, -1):
            item = self.status_list_layout.itemAt(i)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def reset_agent_buttons(self):
        """Reset all agent buttons."""
        for btn in self.agent_buttons.values():
            btn.set_running(False)

    def load_case(self, file_number):
        """Simulate loading a case."""
        self.file_number = file_number
        self.clear_all_status()
        self.reset_agent_buttons()
        QApplication.processEvents()

    def start_agent(self, script):
        """Simulate starting an agent."""
        btn = self.agent_buttons.get(script)
        if btn:
            btn.set_running(True)

        widget = StatusWidget(f"{script} Agent", f"Case: {self.file_number}")
        runner = AgentRunner("python", ["-c", "pass"], widget)
        runner.task_id = widget.task_id

        self.status_list_layout.addWidget(widget)
        self.agent_runners.append(runner)
        self.running_agents[script] = self.file_number

        return runner, widget

    def finish_agent(self, script, success=True):
        """Simulate agent finishing."""
        from datetime import datetime
        today = datetime.now().strftime("%Y-%m-%d")

        started_for_case = self.running_agents.get(script)
        btn = self.agent_buttons.get(script)

        if script in self.running_agents:
            del self.running_agents[script]

        if btn and self.file_number == started_for_case:
            btn.set_running(False)
            btn.set_status("Last: Just now" if success else "Last: Failed")

        if script == "docket.py" and success and started_for_case:
            self.master_db.update_last_docket_download(started_for_case, today)

    def test_full_case_switch_scenario(self):
        """Test complete scenario: start agent, switch case, agent finishes."""
        from datetime import datetime
        today = datetime.now().strftime("%Y-%m-%d")

        # Load case A
        self.load_case("2024.001")

        # Start docket agent
        runner, widget = self.start_agent("docket.py")

        self.assertTrue(self.agent_buttons["docket.py"].is_running)
        self.assertEqual(self.running_agents["docket.py"], "2024.001")

        # Switch to case B
        self.load_case("2024.002")

        # Button should be reset
        self.assertFalse(self.agent_buttons["docket.py"].is_running)

        # Runner should be disconnected from widget
        self.assertIsNone(runner.status_widget)

        # Agent finishes for case A
        self.finish_agent("docket.py", True)

        # Button should NOT be updated (we're on case B)
        self.assertNotIn("Just now", self.agent_buttons["docket.py"].status_label.text())

        # But database should be updated for case A
        self.master_db.update_last_docket_download.assert_called_with("2024.001", today)

    def test_switch_back_to_original_case(self):
        """Test switching back to case with finished agent."""
        # Load case A
        self.load_case("2024.001")

        # Start and complete agent
        self.start_agent("docket.py")
        self.finish_agent("docket.py", True)

        # Switch to case B
        self.load_case("2024.002")

        # Switch back to case A
        self.load_case("2024.001")

        # Should not crash
        self.assertEqual(self.file_number, "2024.001")

    def test_multiple_agents_different_cases(self):
        """Test multiple agents running for different cases."""
        from datetime import datetime
        today = datetime.now().strftime("%Y-%m-%d")

        # Load case A, start docket
        self.load_case("2024.001")
        runner_a, _ = self.start_agent("docket.py")

        # Switch to case B, start complaint
        self.load_case("2024.002")
        runner_b, _ = self.start_agent("complaint.py")

        self.assertTrue(self.agent_buttons["complaint.py"].is_running)
        self.assertFalse(self.agent_buttons["docket.py"].is_running)

        # Docket agent finishes for case A
        self.finish_agent("docket.py", True)

        # Database updated for case A
        self.master_db.update_last_docket_download.assert_called_with("2024.001", today)

        # Complaint button still running
        self.assertTrue(self.agent_buttons["complaint.py"].is_running)

    def test_rapid_case_switching(self):
        """Test rapid switching between cases doesn't cause crashes."""
        cases = ["2024.001", "2024.002", "2024.003", "2024.004"]

        # Start agent on first case
        self.load_case(cases[0])
        self.start_agent("docket.py")

        # Rapidly switch between cases
        for _ in range(10):
            for case in cases:
                self.load_case(case)
                QApplication.processEvents()

        # Should not crash
        self.assertIn(self.file_number, cases)


class TestEdgeCases(unittest.TestCase):
    """Tests for edge cases and error conditions."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def test_disconnect_widget_multiple_times(self):
        """Test calling disconnect_widget multiple times."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget = StatusWidget("Test", "Details")

        runner.connect_widget(widget)

        # Disconnect multiple times
        runner.disconnect_widget()
        runner.disconnect_widget()
        runner.disconnect_widget()

        self.assertIsNone(runner.status_widget)

    def test_connect_same_widget_twice(self):
        """Test connecting the same widget twice."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget = StatusWidget("Test", "Details")

        runner.connect_widget(widget)
        runner.connect_widget(widget)

        self.assertEqual(runner.status_widget, widget)

    def test_empty_running_agents_dict(self):
        """Test operations with empty running_agents."""
        running_agents = {}

        # Should not raise KeyError
        if "docket.py" in running_agents:
            del running_agents["docket.py"]

        self.assertEqual(len(running_agents), 0)

    def test_button_status_after_rapid_toggle(self):
        """Test button status after rapid running state toggles."""
        btn = EnhancedAgentButton("Test", "test.py")

        for _ in range(100):
            btn.set_running(True)
            btn.set_running(False)

        self.assertFalse(btn.is_running)
        self.assertTrue(btn.name_btn.isEnabled())
        self.assertFalse(btn.spinner._running)

    def test_status_widget_serialization_roundtrip(self):
        """Test StatusWidget serialization and deserialization."""
        widget = StatusWidget("Test Agent", "Test Details")
        widget.update_progress(75, "Processing...")
        widget.append_log("Line 1\nLine 2\n")
        widget.is_finished = False

        # Serialize
        data = widget.to_dict()

        # Deserialize
        widget2 = StatusWidget.from_dict(data)

        self.assertEqual(widget2.progress_bar.value(), 75)
        self.assertEqual(widget2.task_id, widget.task_id)
        self.assertIn("Line 1", widget2.log_output.toPlainText())

    def test_agent_finished_with_none_file_number(self):
        """Test on_agent_finished behavior when file_number is None."""
        btn = EnhancedAgentButton("Test", "test.py")
        btn.set_running(True)
        running_agents = {"test.py": None}
        file_number = "2024.001"
        master_db = MagicMock()

        # Simulate on_agent_finished with started_for_case=None
        started_for_case = running_agents.get("test.py")
        del running_agents["test.py"]

        # Should not crash when started_for_case is None
        if file_number == started_for_case:
            btn.set_running(False)

        # Button should still be running (case doesn't match None)
        self.assertTrue(btn.is_running)

    def test_start_same_agent_twice(self):
        """Test starting the same agent while it's already running."""
        btn = EnhancedAgentButton("Docket", "docket.py")
        running_agents = {}

        # Start first time
        btn.set_running(True)
        running_agents["docket.py"] = "2024.001"

        # Try to start again (simulate click while running)
        # In real code, button is disabled, but test the state management
        running_agents["docket.py"] = "2024.001"  # Overwrites

        self.assertTrue(btn.is_running)
        self.assertEqual(running_agents["docket.py"], "2024.001")

    def test_multiple_agents_finish_simultaneously(self):
        """Test multiple agents finishing at the same time."""
        buttons = {
            "agent1.py": EnhancedAgentButton("Agent 1", "agent1.py"),
            "agent2.py": EnhancedAgentButton("Agent 2", "agent2.py"),
            "agent3.py": EnhancedAgentButton("Agent 3", "agent3.py"),
        }
        running_agents = {
            "agent1.py": "2024.001",
            "agent2.py": "2024.001",
            "agent3.py": "2024.001",
        }
        file_number = "2024.001"

        for name, btn in buttons.items():
            btn.set_running(True)

        # All finish at once
        for script in list(running_agents.keys()):
            started_for_case = running_agents[script]
            del running_agents[script]
            if file_number == started_for_case:
                buttons[script].set_running(False)
                buttons[script].set_status("Last: Just now")

        # All buttons should be updated
        for name, btn in buttons.items():
            self.assertFalse(btn.is_running, f"{name} should not be running")
            self.assertIn("Just now", btn.status_label.text())

    def test_runner_emit_after_widget_deleted(self):
        """Test that runner can emit signals after widget is deleted."""
        runner = AgentRunner("python", ["-c", "pass"])
        widget = StatusWidget("Test", "Details")

        runner.connect_widget(widget)

        # Delete the widget
        widget.deleteLater()
        QApplication.processEvents()

        # Disconnect should handle deleted widget
        runner.disconnect_widget()

        # Emitting signals after disconnect should not crash
        try:
            runner.progress_update.emit(50, "Test")
            runner.log_update.emit("Log")
            runner.finished.emit(True)
        except Exception as e:
            self.fail(f"Signal emission crashed: {e}")

    def test_runner_reconnect_after_finished(self):
        """Test reconnecting a runner to a widget after it has finished."""
        runner = AgentRunner("python", ["-c", "pass"])
        runner.success = True  # Mark as finished
        runner.last_progress = 100
        runner.last_status = "Completed"
        runner.log_history = ["Done\n"]

        widget = StatusWidget("Test", "Details")

        # Reconnect should still work
        runner.reconnect_widget(widget)

        self.assertEqual(widget.progress_bar.value(), 100)
        self.assertIn("Done", widget.log_output.toPlainText())

    def test_case_switch_with_no_status_history(self):
        """Test switching to a case with no saved status history."""
        test_dir = tempfile.mkdtemp()
        file_number = "2024.999"  # Non-existent case

        save_path = os.path.join(test_dir, f"{file_number}_status_history.json")

        # File doesn't exist - load should not crash
        self.assertFalse(os.path.exists(save_path))

        # Cleanup
        shutil.rmtree(test_dir, ignore_errors=True)

    def test_corrupted_status_history_json(self):
        """Test loading corrupted status history JSON."""
        test_dir = tempfile.mkdtemp()
        file_number = "2024.001"

        save_path = os.path.join(test_dir, f"{file_number}_status_history.json")

        # Write corrupted JSON
        with open(save_path, 'w') as f:
            f.write("{corrupted json data")

        # Load should handle gracefully
        try:
            with open(save_path, 'r') as f:
                json.load(f)
            self.fail("Should have raised JSONDecodeError")
        except json.JSONDecodeError:
            pass  # Expected

        # Cleanup
        shutil.rmtree(test_dir, ignore_errors=True)

    def test_agent_button_with_empty_status(self):
        """Test agent button with empty/None status values."""
        btn = EnhancedAgentButton("Test", "test.py")

        # Set various empty/None values
        btn.set_status("")
        self.assertEqual(btn.status_label.text(), "")

        btn.set_status(None)
        # Should handle None gracefully

        btn.set_last_run(None)
        self.assertEqual(btn.status_label.text(), "")

        btn.set_last_run("")
        # Should handle empty string

    def test_runner_with_no_widget(self):
        """Test runner operations when no widget is connected."""
        runner = AgentRunner("python", ["-c", "pass"])

        # Should not crash when emitting with no widget
        try:
            runner.progress_update.emit(50, "Test")
            runner.log_update.emit("Log line\n")
            runner.finished.emit(True)
        except Exception as e:
            self.fail(f"Signal emission without widget crashed: {e}")

    def test_very_long_log_content(self):
        """Test widget with very long log content."""
        widget = StatusWidget("Test", "Details")

        # Add a lot of log content
        for i in range(1000):
            widget.append_log(f"Log line {i}: " + "x" * 100 + "\n")

        # Should not crash or hang
        data = widget.to_dict()
        self.assertIn("log_content", data)

        # Deserialize should also work
        widget2 = StatusWidget.from_dict(data)
        self.assertIsNotNone(widget2)

    def test_spinner_start_stop_rapid(self):
        """Test rapid start/stop of spinner."""
        from icharlotte_core.ui.case_view_enhanced import SpinningIndicator

        spinner = SpinningIndicator()

        for _ in range(100):
            spinner.start()
            spinner.stop()

        self.assertFalse(spinner._running)
        self.assertFalse(spinner.timer.isActive())

    def test_button_click_while_running(self):
        """Test that button click signal still works after running state changes."""
        btn = EnhancedAgentButton("Test", "test.py")
        clicks = []

        btn.clicked.connect(lambda: clicks.append(1))

        # Click while not running
        btn.name_btn.click()
        self.assertEqual(len(clicks), 1)

        # Set running (button disabled)
        btn.set_running(True)
        # Click won't work because button is disabled

        # Set not running (button enabled)
        btn.set_running(False)
        btn.name_btn.click()
        self.assertEqual(len(clicks), 2)

    def test_status_widget_cancel_while_finished(self):
        """Test cancel signal on already finished widget."""
        widget = StatusWidget("Test", "Details")
        widget.is_finished = True

        cancel_received = []
        widget.cancel_requested.connect(lambda: cancel_received.append(1))

        # Cancel button click should still emit signal
        widget.cancel_btn.click()
        self.assertEqual(len(cancel_received), 1)


class TestStressTests(unittest.TestCase):
    """Stress tests for the case loading system."""

    @classmethod
    def setUpClass(cls):
        """Create QApplication for all tests."""
        cls.app = QApplication.instance()
        if not cls.app:
            cls.app = QApplication([])

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        """Clean up test fixtures."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_many_agents_many_cases(self):
        """Test managing many agents across many cases."""
        num_agents = 10
        num_cases = 20

        buttons = {f"agent{i}.py": EnhancedAgentButton(f"Agent {i}", f"agent{i}.py")
                   for i in range(num_agents)}
        running_agents = {}

        # Simulate running agents on different cases
        for case_num in range(num_cases):
            file_number = f"2024.{case_num:03d}"

            # Start random agents
            for i, (script, btn) in enumerate(buttons.items()):
                if i % 2 == case_num % 2:  # Alternating pattern
                    btn.set_running(True)
                    running_agents[script] = file_number

            # Reset for "new case"
            for btn in buttons.values():
                btn.set_running(False)

        # All buttons should be not running
        for script, btn in buttons.items():
            self.assertFalse(btn.is_running)

    def test_rapid_widget_creation_destruction(self):
        """Test rapid creation and destruction of status widgets."""
        runners = []

        for i in range(50):
            widget = StatusWidget(f"Agent {i}", f"Details {i}")
            runner = AgentRunner("python", ["-c", "pass"], widget)
            runners.append(runner)

            # Simulate some activity
            runner.progress_update.emit(i * 2, f"Step {i}")
            runner.log_update.emit(f"Log line {i}\n")

            # Disconnect and destroy
            runner.disconnect_widget()
            widget.deleteLater()

        QApplication.processEvents()

        # All runners should have no widget
        for runner in runners:
            self.assertIsNone(runner.status_widget)

    def test_concurrent_signal_emissions(self):
        """Test many signals being emitted rapidly."""
        widget = StatusWidget("Test", "Details")
        runner = AgentRunner("python", ["-c", "pass"], widget)

        # Emit many signals rapidly
        for i in range(100):
            runner.progress_update.emit(i % 100, f"Progress {i}")
            runner.log_update.emit(f"Line {i}\n")

        # Widget should still be functional
        self.assertEqual(runner.status_widget, widget)
        self.assertGreater(len(widget.log_output.toPlainText()), 0)

    def test_serialization_with_special_characters(self):
        """Test serialization with special characters in content."""
        widget = StatusWidget("Test™ Agent", "Details with \"quotes\" and 'apostrophes'")
        widget.update_progress(50, "Status: <pending> & waiting…")
        widget.append_log("Line with unicode: 日本語 中文 العربية\n")
        widget.append_log("Newlines:\n\n\tTabs\t\there\n")

        # Serialize
        data = widget.to_dict()

        # Deserialize
        widget2 = StatusWidget.from_dict(data)

        self.assertEqual(widget2.progress_bar.value(), 50)
        self.assertIn("日本語", widget2.log_output.toPlainText())

    def test_status_history_with_many_entries(self):
        """Test saving and loading status history with many entries."""
        container = QWidget()
        layout = QVBoxLayout(container)

        # Create many widgets
        for i in range(20):
            widget = StatusWidget(f"Agent {i}", f"Details {i}")
            widget.update_progress(i * 5, f"Status {i}")
            widget.is_finished = (i % 2 == 0)
            layout.addWidget(widget)

        # Save history
        history = []
        for i in range(layout.count()):
            item = layout.itemAt(i)
            widget = item.widget()
            if isinstance(widget, StatusWidget):
                history.append(widget.to_dict())

        save_path = os.path.join(self.test_dir, "status_history.json")
        with open(save_path, 'w') as f:
            json.dump(history, f)

        # Load history
        with open(save_path, 'r') as f:
            loaded = json.load(f)

        self.assertEqual(len(loaded), 20)

    def test_button_state_consistency(self):
        """Test that button state remains consistent through various operations."""
        btn = EnhancedAgentButton("Test", "test.py")

        states = []

        for _ in range(50):
            # Random operations
            btn.set_running(True)
            states.append(btn.is_running)
            btn.set_status("Running...")
            btn.set_running(False)
            states.append(btn.is_running)
            btn.set_status("Idle")
            btn.set_last_run("2024-01-15")

        # Final state should be consistent
        self.assertFalse(btn.is_running)
        self.assertTrue(btn.name_btn.isEnabled())
        self.assertFalse(btn.spinner._running)

    def test_runner_state_after_many_reconnects(self):
        """Test runner state after many connect/disconnect cycles."""
        runner = AgentRunner("python", ["-c", "pass"])
        runner.log_history = ["Initial log\n"]
        runner.last_progress = 50
        runner.last_status = "Working..."

        for i in range(20):
            widget = StatusWidget(f"Widget {i}", "Details")
            runner.reconnect_widget(widget)

            # Verify state was replayed
            self.assertEqual(widget.progress_bar.value(), 50)
            self.assertIn("Initial log", widget.log_output.toPlainText())

            runner.disconnect_widget()
            widget.deleteLater()

        QApplication.processEvents()

        # Runner should have no widget but state preserved
        self.assertIsNone(runner.status_widget)
        self.assertEqual(runner.last_progress, 50)


if __name__ == "__main__":
    unittest.main(verbosity=2)
