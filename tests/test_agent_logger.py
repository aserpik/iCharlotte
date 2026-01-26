"""
Tests for Agent Logger Module

Tests structured logging and pass progress reporting.
"""

import os
import sys
import pytest
from unittest.mock import Mock, patch
from io import StringIO

# Add parent to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core.agent_logger import AgentLogger, create_legacy_log_event


class TestAgentLogger:
    """Tests for AgentLogger class."""

    def test_initialization(self):
        logger = AgentLogger("TestAgent")
        assert logger.agent_name == "TestAgent"
        assert logger.file_number is None

    def test_initialization_with_file_number(self):
        logger = AgentLogger("TestAgent", file_number="2024.123")
        assert logger.file_number == "2024.123"

    def test_info_message(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.info("Test message")
        captured = capsys.readouterr()
        assert "Test message" in captured.out

    def test_error_message(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.error("Error occurred")
        captured = capsys.readouterr()
        assert "Error occurred" in captured.out

    def test_warning_message(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.warning("Warning message")
        captured = capsys.readouterr()
        assert "Warning message" in captured.out

    def test_pass_start(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.pass_start("Extraction", 1, 3)
        captured = capsys.readouterr()
        assert "PASS_START:Extraction:1:3" in captured.out

    def test_pass_complete(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.pass_start("Extraction", 1, 3)  # Start first
        logger.pass_complete("Extraction", success=True)
        captured = capsys.readouterr()
        assert "PASS_COMPLETE:Extraction:success" in captured.out

    def test_pass_complete_with_duration(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.pass_start("Extraction", 1, 3)
        logger.pass_complete("Extraction", success=True, duration_sec=5.5)
        captured = capsys.readouterr()
        assert "5.5" in captured.out

    def test_pass_failed(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.pass_failed("Extraction", "Network error", recoverable=True)
        captured = capsys.readouterr()
        assert "PASS_FAILED:Extraction:Network error:recoverable" in captured.out

    def test_pass_failed_not_recoverable(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.pass_failed("Extraction", "Fatal error", recoverable=False)
        captured = capsys.readouterr()
        assert "not_recoverable" in captured.out

    def test_progress(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.progress(50, "Halfway done")
        captured = capsys.readouterr()
        assert "PROGRESS:50:Halfway done" in captured.out

    def test_output_file(self, capsys):
        logger = AgentLogger("TestAgent")
        logger.output_file("/path/to/output.docx")
        captured = capsys.readouterr()
        assert "OUTPUT_FILE:/path/to/output.docx" in captured.out


class TestLegacyLogEvent:
    """Tests for legacy log event function."""

    def test_create_legacy_log_event(self):
        log_event = create_legacy_log_event("TestAgent", "/tmp/test.log")
        assert callable(log_event)

    def test_legacy_log_event_output(self, capsys):
        log_event = create_legacy_log_event("TestAgent", "/tmp/test.log")
        log_event("Test message")
        captured = capsys.readouterr()
        assert "Test message" in captured.out

    def test_legacy_log_event_error_level(self, capsys):
        log_event = create_legacy_log_event("TestAgent", "/tmp/test.log")
        log_event("Error message", level="error")
        captured = capsys.readouterr()
        assert "Error message" in captured.out


class TestLoggerPassTracking:
    """Tests for pass timing and tracking."""

    def test_pass_timing(self, capsys):
        import time
        logger = AgentLogger("TestAgent")

        logger.pass_start("Slow Pass", 1, 1)
        time.sleep(0.1)  # 100ms
        logger.pass_complete("Slow Pass", success=True)

        captured = capsys.readouterr()
        # Should have recorded some duration
        assert "PASS_COMPLETE" in captured.out

    def test_multiple_passes(self, capsys):
        logger = AgentLogger("TestAgent")

        logger.pass_start("Pass1", 1, 3)
        logger.pass_complete("Pass1", success=True)

        logger.pass_start("Pass2", 2, 3)
        logger.pass_complete("Pass2", success=True)

        logger.pass_start("Pass3", 3, 3)
        logger.pass_complete("Pass3", success=True)

        captured = capsys.readouterr()
        assert "Pass1" in captured.out
        assert "Pass2" in captured.out
        assert "Pass3" in captured.out


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
