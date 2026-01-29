"""
Comprehensive UI Crash Safety Tests for iCharlotte

This test suite validates that the UI components handle edge cases gracefully
without crashing. Focuses on:
- Master Case Tab (table operations, database interactions)
- Case View Enhanced (file tree, tags, processing logs)
- Status Widgets (agent runner, progress tracking)
- Index Tab (chat, OCR operations)

Tests include:
- Null/empty data handling
- Rapid user interactions
- Database failures
- Threading edge cases
- Widget lifecycle issues
"""

import unittest
import sys
import os
import json
import tempfile
import shutil
from unittest.mock import MagicMock, patch, PropertyMock

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import Qt first to ensure proper initialization
from PySide6.QtWidgets import QApplication, QTableWidget, QTableWidgetItem, QListWidget
from PySide6.QtCore import Qt, QDate

# Create QApplication instance for tests
app = QApplication.instance()
if not app:
    app = QApplication([])


class TestTableItemSafety(unittest.TestCase):
    """Test table item null safety patterns."""

    def setUp(self):
        """Create a test table widget."""
        self.table = QTableWidget(5, 7)

    def tearDown(self):
        self.table.deleteLater()

    def test_empty_table_item_access(self):
        """Test accessing items in empty table doesn't crash."""
        # Table has rows but no items set
        for row in range(5):
            for col in range(7):
                item = self.table.item(row, col)
                # Should return None, not crash
                self.assertIsNone(item)

    def test_item_data_with_null_item(self):
        """Test safe pattern for accessing item data."""
        # This is what crashes: self.table.item(0, 0).data(...)
        # Safe pattern:
        item = self.table.item(0, 0)
        if item:
            data = item.data(Qt.ItemDataRole.UserRole)
        else:
            data = None
        self.assertIsNone(data)

    def test_mixed_populated_table(self):
        """Test table with some cells populated, some empty."""
        # Only set some items
        self.table.setItem(0, 0, QTableWidgetItem("test"))
        self.table.setItem(2, 3, QTableWidgetItem("another"))

        # Iterate all cells safely
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                text = item.text() if item else ""
                # Should not crash

    def test_row_selection_with_no_items(self):
        """Test selecting rows when table has no items."""
        self.table.selectRow(0)
        selected = self.table.selectionModel().selectedRows()
        self.assertEqual(len(selected), 1)
        # But the item at that row is None
        row = selected[0].row()
        item = self.table.item(row, 0)
        self.assertIsNone(item)


class TestMasterCaseTabSafety(unittest.TestCase):
    """Test MasterCaseTab crash safety."""

    @classmethod
    def setUpClass(cls):
        """Set up test fixtures."""
        cls.temp_dir = tempfile.mkdtemp()

    @classmethod
    def tearDownClass(cls):
        """Clean up test fixtures."""
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def test_delete_table_cell_with_null_item(self):
        """Test delete_table_cell handles null items gracefully."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []

            tab = MasterCaseTab()
            # Set current cell to one that doesn't exist
            tab.table.setCurrentCell(0, 0)

            # Should not crash even if item is None
            try:
                tab.delete_table_cell()
            except AttributeError:
                self.fail("delete_table_cell crashed on null item")
            finally:
                tab.deleteLater()

    def test_on_case_selected_with_empty_table(self):
        """Test case selection when table is empty."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []

            tab = MasterCaseTab()
            # Try to select with empty table
            tab.table.selectRow(0)

            try:
                tab.on_case_selected()
            except (AttributeError, TypeError):
                self.fail("on_case_selected crashed on empty table")
            finally:
                tab.deleteLater()

    def test_on_cell_clicked_with_null_item(self):
        """Test cell click handlers with null items."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []

            tab = MasterCaseTab()
            # Simulate click on empty cell
            try:
                tab.on_cell_clicked(0, 0)
            except AttributeError:
                self.fail("on_cell_clicked crashed on null item")
            finally:
                tab.deleteLater()

    def test_refresh_data_with_invalid_json(self):
        """Test refresh_data handles corrupted JSON files."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = [
                {
                    'file_number': '1234.567',
                    'plaintiff_last_name': 'Test',
                    'next_hearing_date': '',
                    'trial_date': '',
                    'case_path': '',
                    'assigned_attorney': '',
                    'last_report_text': '',
                    'plaintiff_override': 0
                }
            ]
            mock_db.return_value.get_todos.return_value = []
            mock_db.return_value.get_last_status_update_date.return_value = None

            # Create corrupted JSON - patch at config module level
            with patch('icharlotte_core.config.GEMINI_DATA_DIR', self.temp_dir):
                json_path = os.path.join(self.temp_dir, '1234.567.json')
                with open(json_path, 'w') as f:
                    f.write('{invalid json')

                tab = MasterCaseTab()
                try:
                    tab.refresh_data()
                except json.JSONDecodeError:
                    self.fail("refresh_data crashed on invalid JSON")
                finally:
                    tab.deleteLater()

    def test_save_dates_without_current_case(self):
        """Test save_dates when no case is selected."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []

            tab = MasterCaseTab()
            # Remove current_file_number attribute
            if hasattr(tab, 'current_file_number'):
                delattr(tab, 'current_file_number')

            try:
                tab.save_dates()
            except AttributeError:
                self.fail("save_dates crashed without current case")
            finally:
                tab.deleteLater()

    def test_restore_selection_without_current_case(self):
        """Test restore_selection when no case was selected."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []

            tab = MasterCaseTab()
            # Remove current_file_number attribute
            if hasattr(tab, 'current_file_number'):
                delattr(tab, 'current_file_number')

            try:
                tab.restore_selection()
            except AttributeError:
                self.fail("restore_selection crashed without current case")
            finally:
                tab.deleteLater()


class TestDatabaseReturnValidation(unittest.TestCase):
    """Test handling of None returns from database methods."""

    def test_get_case_returns_none(self):
        """Test handling when get_case returns None."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []
            mock_db.return_value.get_case.return_value = None

            tab = MasterCaseTab()
            tab.current_file_number = 'nonexistent'

            try:
                tab.refresh_details()
            except (TypeError, AttributeError):
                self.fail("refresh_details crashed on None case")
            finally:
                tab.deleteLater()

    def test_get_todos_returns_none(self):
        """Test handling when get_todos returns None instead of empty list."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []
            mock_db.return_value.get_case.return_value = {'assigned_attorney': 'HP'}
            mock_db.return_value.get_todos.return_value = None  # Edge case
            mock_db.return_value.get_history.return_value = []

            tab = MasterCaseTab()
            tab.current_file_number = 'test123'

            try:
                tab.refresh_details()
            except (TypeError, AttributeError):
                self.fail("refresh_details crashed on None todos")
            finally:
                tab.deleteLater()


class TestFileTreeWidgetSafety(unittest.TestCase):
    """Test FileTreeWidget crash safety."""

    def test_drop_on_empty_tree(self):
        """Test drop event when tree is empty."""
        from icharlotte_core.ui.widgets import FileTreeWidget

        tree = FileTreeWidget()
        # itemAt should return None for empty tree
        item = tree.itemAt(0, 0)
        self.assertIsNone(item)
        tree.deleteLater()

    def test_context_menu_on_empty_tree(self):
        """Test context menu on empty tree doesn't crash."""
        from icharlotte_core.ui.widgets import FileTreeWidget
        from PySide6.QtCore import QPoint

        tree = FileTreeWidget()
        try:
            tree.open_context_menu(QPoint(0, 0))
        except AttributeError:
            self.fail("Context menu crashed on empty tree")
        finally:
            tree.deleteLater()


class TestStatusWidgetSafety(unittest.TestCase):
    """Test StatusWidget and AgentRunner crash safety."""

    def test_status_widget_update_after_deletion(self):
        """Test that signals don't crash after widget deletion."""
        from icharlotte_core.ui.widgets import StatusWidget, AgentRunner

        widget = StatusWidget("Test Agent", "Testing")
        runner = AgentRunner("python", ["-c", "print('test')"])
        runner.connect_widget(widget)

        # Simulate widget deletion before process finishes
        widget.deleteLater()
        app.processEvents()

        # Runner should handle disconnected widget gracefully
        try:
            runner.progress_update.emit(50, "test")
            runner.log_update.emit("test log")
        except RuntimeError:
            pass  # Expected - widget is deleted

    def test_agent_runner_cancel_when_not_running(self):
        """Test canceling runner that's not running."""
        from icharlotte_core.ui.widgets import AgentRunner

        runner = AgentRunner("python", ["-c", "pass"])
        try:
            runner.cancel()
        except Exception as e:
            self.fail(f"Cancel crashed: {e}")


class TestChatTabSafety(unittest.TestCase):
    """Test ChatTab crash safety."""

    def test_load_case_without_persistence(self):
        """Test loading case when file doesn't exist."""
        from icharlotte_core.ui.tabs import ChatTab

        tab = ChatTab()
        try:
            tab.load_case("nonexistent_case_12345")
        except FileNotFoundError:
            self.fail("load_case crashed on missing case")
        finally:
            tab.deleteLater()

    def test_send_message_without_conversation(self):
        """Test sending message when no conversation exists."""
        from icharlotte_core.ui.tabs import ChatTab

        tab = ChatTab()
        tab.chat_input.setPlainText("Test message")

        try:
            # This should handle no conversation gracefully
            tab.send_message()
        except (AttributeError, TypeError):
            self.fail("send_message crashed without conversation")
        finally:
            tab.deleteLater()

    def test_clear_files_with_none_list(self):
        """Test clearing files when list is not initialized."""
        from icharlotte_core.ui.tabs import ChatTab

        tab = ChatTab()
        tab.attached_files = None

        try:
            tab.clear_files()
        except (TypeError, AttributeError):
            self.fail("clear_files crashed with None list")
        finally:
            tab.deleteLater()


class TestProcessingLogDBSafety(unittest.TestCase):
    """Test ProcessingLogDB crash safety."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_load_corrupted_file(self):
        """Test loading corrupted JSON log file."""
        from icharlotte_core.ui.case_view_enhanced import ProcessingLogDB

        with patch('icharlotte_core.ui.case_view_enhanced.GEMINI_DATA_DIR', self.temp_dir):
            # Create corrupted file
            log_path = os.path.join(self.temp_dir, "test123_processing_log.json")
            with open(log_path, 'w') as f:
                f.write("not valid json {{{")

            try:
                db = ProcessingLogDB("test123")
                # Should return empty list, not crash
                self.assertEqual(db.logs, [])
            except json.JSONDecodeError:
                self.fail("ProcessingLogDB crashed on corrupted file")

    def test_save_with_readonly_directory(self):
        """Test saving when directory is not writable."""
        from icharlotte_core.ui.case_view_enhanced import ProcessingLogDB

        with patch('icharlotte_core.ui.case_view_enhanced.GEMINI_DATA_DIR', self.temp_dir):
            db = ProcessingLogDB("test456")

            # Make directory readonly (Windows doesn't fully support this, skip on Windows)
            if os.name != 'nt':
                os.chmod(self.temp_dir, 0o444)
                try:
                    db.add_entry("/test/path", "summary", "success")
                except PermissionError:
                    self.fail("ProcessingLogDB crashed on readonly directory")
                finally:
                    os.chmod(self.temp_dir, 0o755)


class TestEdgeCases(unittest.TestCase):
    """Test various edge cases that could cause crashes."""

    def test_parse_hearing_data_with_invalid_input(self):
        """Test hearing data parsing with invalid formats."""
        from icharlotte_core.utils import parse_hearing_data

        test_cases = [
            "",  # Empty string
            None,  # None (should handle gracefully)
            "invalid date format",
            "CMC no date",
            "12/invalid/2024",
            "\n\n\n",  # Only newlines
            "    ",  # Only whitespace
        ]

        for test in test_cases:
            try:
                result = parse_hearing_data(test)
                # Should return a list (possibly empty)
                self.assertIsInstance(result, list)
            except Exception as e:
                self.fail(f"parse_hearing_data crashed on '{test}': {e}")

    def test_date_table_widget_item_comparison_with_invalid_dates(self):
        """Test DateTableWidgetItem sorting with invalid dates."""
        from icharlotte_core.ui.master_case_tab import DateTableWidgetItem

        item1 = DateTableWidgetItem("")
        item2 = DateTableWidgetItem("invalid")
        item3 = DateTableWidgetItem("2024-01-15")

        try:
            # These comparisons should not crash
            _ = item1 < item2
            _ = item2 < item3
            _ = item1 < item3
        except Exception as e:
            self.fail(f"DateTableWidgetItem comparison crashed: {e}")


class TestThreadingSafety(unittest.TestCase):
    """Test threading and worker safety."""

    def test_worker_signals_after_thread_finished(self):
        """Test that emitting signals after thread finish doesn't crash."""
        from icharlotte_core.ui.widgets import AgentRunner

        runner = AgentRunner("python", ["-c", "pass"])

        # Simulate thread finishing
        runner.success = True

        # Emitting signals should be safe
        try:
            runner.progress_update.emit(100, "Done")
        except RuntimeError:
            pass  # May be expected if no slots connected

    def test_disconnect_widget_when_not_connected(self):
        """Test disconnecting widget that was never connected."""
        from icharlotte_core.ui.widgets import AgentRunner

        runner = AgentRunner("python", ["-c", "pass"])

        try:
            runner.disconnect_widget()
        except Exception as e:
            self.fail(f"disconnect_widget crashed: {e}")


class TestTodoItemWidgetSafety(unittest.TestCase):
    """Test TodoItemWidget crash safety."""

    def test_todo_with_none_values(self):
        """Test creating TodoItemWidget with None values."""
        from icharlotte_core.ui.master_case_tab import TodoItemWidget

        try:
            widget = TodoItemWidget(
                text="Test",
                status=None,
                color=None,
                created_date=None,
                assigned_to=None,
                assigned_date=None,
                case_assigned_attorney=None
            )
            widget.deleteLater()
        except (TypeError, AttributeError):
            self.fail("TodoItemWidget crashed with None values")

    def test_cycle_color_with_invalid_current(self):
        """Test color cycling with invalid current color."""
        from icharlotte_core.ui.master_case_tab import TodoItemWidget

        widget = TodoItemWidget(
            text="Test",
            status="pending",
            color="invalid_color",  # Invalid color
            created_date="2024-01-15",
            assigned_to="",
            assigned_date="",
            case_assigned_attorney=""
        )

        try:
            widget.cycle_color()
        except Exception as e:
            self.fail(f"cycle_color crashed with invalid color: {e}")
        finally:
            widget.deleteLater()


class TestHearingCellWidgetSafety(unittest.TestCase):
    """Test HearingCellWidget crash safety."""

    def test_empty_hearings_list(self):
        """Test HearingCellWidget with empty hearings list."""
        from icharlotte_core.ui.master_case_tab import HearingCellWidget

        try:
            widget = HearingCellWidget([])
            self.assertIsNone(widget.display_hearing)
            widget.deleteLater()
        except Exception as e:
            self.fail(f"HearingCellWidget crashed with empty list: {e}")

    def test_hearings_with_invalid_dates(self):
        """Test HearingCellWidget with invalid date objects."""
        from icharlotte_core.ui.master_case_tab import HearingCellWidget
        import datetime

        hearings = [
            {
                'date_obj': datetime.datetime(9999, 12, 31),  # "No date" marker
                'date_sort': '9999-99-99',
                'display': 'CMC (no date)'
            }
        ]

        try:
            widget = HearingCellWidget(hearings)
            widget.deleteLater()
        except Exception as e:
            self.fail(f"HearingCellWidget crashed with special dates: {e}")


class TestRapidInteractions(unittest.TestCase):
    """Test rapid user interactions that could cause race conditions."""

    def test_rapid_refresh_data(self):
        """Test multiple rapid refresh_data calls don't crash."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = [
                {
                    'file_number': f'1234.{i:03d}',
                    'plaintiff_last_name': f'Test{i}',
                    'next_hearing_date': '',
                    'trial_date': '',
                    'case_path': '',
                    'assigned_attorney': 'HP',
                    'last_report_text': '',
                    'plaintiff_override': 0
                } for i in range(50)
            ]
            mock_db.return_value.get_todos.return_value = []
            mock_db.return_value.get_last_status_update_date.return_value = None

            tab = MasterCaseTab()
            try:
                # Rapid refresh calls
                for _ in range(10):
                    tab.refresh_data()
                    app.processEvents()
            except Exception as e:
                self.fail(f"Rapid refresh crashed: {e}")
            finally:
                tab.deleteLater()

    def test_rapid_case_selection(self):
        """Test rapidly selecting different cases doesn't crash."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = [
                {
                    'file_number': f'1234.{i:03d}',
                    'plaintiff_last_name': f'Test{i}',
                    'next_hearing_date': '2024-06-15',
                    'trial_date': '',
                    'case_path': '',
                    'assigned_attorney': 'HP',
                    'last_report_text': '',
                    'plaintiff_override': 0
                } for i in range(10)
            ]
            mock_db.return_value.get_todos.return_value = []
            mock_db.return_value.get_history.return_value = []
            mock_db.return_value.get_case.return_value = {
                'assigned_attorney': 'HP',
                'case_summary': 'Test',
                'next_hearing_date': '2024-06-15',
                'trial_date': ''
            }
            mock_db.return_value.get_last_status_update_date.return_value = None

            tab = MasterCaseTab()
            try:
                # Rapidly select different rows
                for i in range(10):
                    tab.table.selectRow(i % 10)
                    app.processEvents()
            except Exception as e:
                self.fail(f"Rapid selection crashed: {e}")
            finally:
                tab.deleteLater()


class TestWorkerThreadSafety(unittest.TestCase):
    """Test worker thread lifecycle safety."""

    def test_case_scanner_worker_stop(self):
        """Test stopping CaseScannerWorker gracefully."""
        from icharlotte_core.ui.master_case_tab import CaseScannerWorker

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []
            worker = CaseScannerWorker(mock_db.return_value)

            try:
                worker.stop()
                self.assertFalse(worker.running)
            except Exception as e:
                self.fail(f"CaseScannerWorker stop crashed: {e}")

    def test_multiple_agent_runners(self):
        """Test creating and managing multiple AgentRunners."""
        from icharlotte_core.ui.widgets import AgentRunner, StatusWidget

        runners = []
        widgets = []

        try:
            for i in range(5):
                widget = StatusWidget(f"Test Agent {i}", f"Testing {i}")
                runner = AgentRunner("python", ["-c", "pass"], widget)
                runners.append(runner)
                widgets.append(widget)

            # Disconnect all widgets
            for runner in runners:
                runner.disconnect_widget()

        except Exception as e:
            self.fail(f"Multiple runners crashed: {e}")
        finally:
            for widget in widgets:
                widget.deleteLater()


class TestFilterTableEdgeCases(unittest.TestCase):
    """Test filter_table edge cases."""

    def test_filter_with_none_items(self):
        """Test filtering table with some None items."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = []

            tab = MasterCaseTab()
            # Set up table with some rows but no items
            tab.table.setRowCount(5)

            try:
                tab.filter_table("test")
            except Exception as e:
                self.fail(f"filter_table crashed with None items: {e}")
            finally:
                tab.deleteLater()

    def test_filter_with_unicode(self):
        """Test filtering with unicode characters."""
        from icharlotte_core.ui.master_case_tab import MasterCaseTab

        with patch('icharlotte_core.ui.master_case_tab.MasterCaseDatabase') as mock_db:
            mock_db.return_value.get_all_cases.return_value = [
                {
                    'file_number': '1234.567',
                    'plaintiff_last_name': 'Müller',
                    'next_hearing_date': '',
                    'trial_date': '',
                    'case_path': '',
                    'assigned_attorney': '',
                    'last_report_text': '',
                    'plaintiff_override': 0
                }
            ]
            mock_db.return_value.get_todos.return_value = []
            mock_db.return_value.get_last_status_update_date.return_value = None

            tab = MasterCaseTab()
            try:
                tab.filter_table("Müller")
                tab.filter_table("日本語")
                tab.filter_table("emoji🎉")
            except Exception as e:
                self.fail(f"Unicode filter crashed: {e}")
            finally:
                tab.deleteLater()


class TestCalendarDialogEdgeCases(unittest.TestCase):
    """Test CalendarDialog edge cases."""

    def test_calendar_dialog_with_null_date(self):
        """Test CalendarDialog with None current_date."""
        from icharlotte_core.ui.master_case_tab import CalendarDialog

        try:
            dialog = CalendarDialog(None, None)
            dialog.deleteLater()
        except Exception as e:
            self.fail(f"CalendarDialog with None date crashed: {e}")

    def test_calendar_dialog_clear(self):
        """Test CalendarDialog clear functionality."""
        from icharlotte_core.ui.master_case_tab import CalendarDialog
        from PySide6.QtCore import QDate

        try:
            dialog = CalendarDialog(None, QDate(2024, 6, 15))
            dialog.clear_date()
            self.assertTrue(dialog.cleared)
            dialog.deleteLater()
        except Exception as e:
            self.fail(f"CalendarDialog clear crashed: {e}")


class TestPassProgressWidgetEdgeCases(unittest.TestCase):
    """Test PassProgressWidget edge cases."""

    def test_pass_started_without_setup(self):
        """Test pass_started without calling set_total_passes first."""
        from icharlotte_core.ui.widgets import PassProgressWidget

        widget = PassProgressWidget()
        try:
            # Should dynamically create pass info
            widget.pass_started("Dynamic Pass", 1, 3)
            self.assertIn("Dynamic Pass", widget.passes)
        except Exception as e:
            self.fail(f"Dynamic pass creation crashed: {e}")
        finally:
            widget.deleteLater()

    def test_pass_completed_for_unknown_pass(self):
        """Test pass_completed for a pass that was never started."""
        from icharlotte_core.ui.widgets import PassProgressWidget

        widget = PassProgressWidget()
        try:
            widget.pass_completed("Unknown Pass", 5.0)
            # Should not crash, just do nothing
        except Exception as e:
            self.fail(f"Unknown pass completed crashed: {e}")
        finally:
            widget.deleteLater()

    def test_pass_failed_with_empty_error(self):
        """Test pass_failed with empty error message."""
        from icharlotte_core.ui.widgets import PassProgressWidget

        widget = PassProgressWidget()
        widget.set_total_passes(2, ["Pass 1", "Pass 2"])

        try:
            widget.pass_failed("Pass 1", "", True)
        except Exception as e:
            self.fail(f"pass_failed with empty error crashed: {e}")
        finally:
            widget.deleteLater()


class TestAgentRunnerParseProgress(unittest.TestCase):
    """Test AgentRunner progress parsing edge cases."""

    def test_parse_malformed_progress(self):
        """Test parsing malformed progress strings."""
        from icharlotte_core.ui.widgets import AgentRunner

        runner = AgentRunner("python", ["-c", "pass"])

        malformed_inputs = [
            "PROGRESS:",
            "PROGRESS:abc:def",
            "PROGRESS:100:",
            "PASS_START:",
            "PASS_START:name:",
            "PASS_START:name:abc:def",
            "PASS_COMPLETE:",
            "PASS_FAILED:",
            "OUTPUT_FILE:",
            "PROGRESS:9999999:test",  # Large but valid integer (will be clamped)
            "PROGRESS:-100:negative test",  # Negative (will be clamped to 0)
        ]

        try:
            for text in malformed_inputs:
                runner.parse_progress(text)
        except Exception as e:
            self.fail(f"parse_progress crashed on malformed input: {e}")


class TestEnhancedFileTreeWidget(unittest.TestCase):
    """Test EnhancedFileTreeWidget from case_view_enhanced."""

    def test_processing_log_db_with_missing_file(self):
        """Test ProcessingLogDB when log file doesn't exist."""
        from icharlotte_core.ui.case_view_enhanced import ProcessingLogDB
        import tempfile

        with tempfile.TemporaryDirectory() as temp_dir:
            with patch('icharlotte_core.ui.case_view_enhanced.GEMINI_DATA_DIR', temp_dir):
                try:
                    db = ProcessingLogDB("nonexistent_case")
                    self.assertEqual(db.logs, [])
                except Exception as e:
                    self.fail(f"ProcessingLogDB with missing file crashed: {e}")

    def test_file_tags_db_with_missing_file(self):
        """Test FileTagsDB when tags file doesn't exist."""
        from icharlotte_core.ui.case_view_enhanced import FileTagsDB
        import tempfile

        with tempfile.TemporaryDirectory() as temp_dir:
            with patch('icharlotte_core.ui.case_view_enhanced.GEMINI_DATA_DIR', temp_dir):
                try:
                    db = FileTagsDB("nonexistent_case")
                    self.assertEqual(db.tags, {})
                    # Test adding tags
                    db.add_tag("/path/to/file.pdf", "important")
                    tags = db.get_tags("/path/to/file.pdf")
                    self.assertIn("important", tags)
                except Exception as e:
                    self.fail(f"FileTagsDB with missing file crashed: {e}")


if __name__ == '__main__':
    # Run tests
    unittest.main(verbosity=2)
