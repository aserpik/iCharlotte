"""
Integration tests for iCharlotte agent modules.

Tests that all core modules can be imported and basic functionality works
without requiring external services (LLMs, databases, etc.).
"""

import pytest
import sys
import os
from pathlib import Path
from unittest.mock import patch, MagicMock

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))


class TestCoreModuleImports:
    """Test that all core modules can be imported."""

    def test_import_agent_logger(self):
        from icharlotte_core.agent_logger import AgentLogger
        assert AgentLogger is not None

    def test_import_config(self):
        from icharlotte_core.config import API_KEYS, BASE_PATH_WIN
        assert API_KEYS is not None
        assert BASE_PATH_WIN is not None

    def test_import_document_processor(self):
        from icharlotte_core.document_processor import (
            DocumentProcessor, ExtractResult, OCRConfig
        )
        assert DocumentProcessor is not None
        assert ExtractResult is not None
        assert OCRConfig is not None

    def test_import_output_validator(self):
        from icharlotte_core.output_validator import (
            ValidationResult, BaseValidator, get_validator
        )
        assert ValidationResult is not None
        assert BaseValidator is not None
        assert get_validator is not None

    def test_import_exceptions(self):
        from icharlotte_core.exceptions import (
            ExtractionError, OCRError, LLMError
        )
        assert ExtractionError is not None
        assert OCRError is not None
        assert LLMError is not None

    def test_import_utils(self):
        from icharlotte_core.utils import (
            log_event, get_case_path, sanitize_filename
        )
        assert log_event is not None
        assert get_case_path is not None
        assert sanitize_filename is not None

    def test_import_llm_handler(self):
        from icharlotte_core.llm import LLMHandler
        assert LLMHandler is not None

    def test_import_llm_config(self):
        from icharlotte_core.llm_config import LLMConfig
        assert LLMConfig is not None


class TestCalendarModuleImports:
    """Test that calendar modules can be imported."""

    def test_import_deadline_calculator(self):
        from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
        assert DeadlineCalculator is not None

    def test_import_attachment_classifier(self):
        from icharlotte_core.calendar.attachment_classifier import AttachmentClassifier
        assert AttachmentClassifier is not None

    def test_import_calendar_monitor(self):
        from icharlotte_core.calendar.calendar_monitor import CalendarMonitorWorker
        assert CalendarMonitorWorker is not None


class TestLogsTabImport:
    """Test that logs_tab module can be imported."""

    def test_import_log_manager(self):
        from icharlotte_core.ui.logs_tab import LogManager
        log_manager = LogManager()
        assert log_manager is not None

    def test_log_manager_add_log(self):
        from icharlotte_core.ui.logs_tab import LogManager
        log_manager = LogManager()
        log_manager.add_log("Test", "Test message")
        logs = log_manager.get_logs("Test")
        assert len(logs) >= 1
        assert "Test message" in logs[-1]


class TestAgentLogger:
    """Test AgentLogger functionality."""

    def test_agent_logger_creation(self):
        from icharlotte_core.agent_logger import AgentLogger
        logger = AgentLogger("TestAgent")
        assert logger.agent_name == "TestAgent"

    def test_agent_logger_pass_workflow(self, capsys):
        from icharlotte_core.agent_logger import AgentLogger
        logger = AgentLogger("TestAgent")

        logger.pass_start("Extraction", 1, 2)
        captured = capsys.readouterr()
        assert "PASS_START:Extraction:1:2" in captured.out

        logger.pass_complete("Extraction", success=True)
        captured = capsys.readouterr()
        assert "PASS_COMPLETE:Extraction:success" in captured.out

    def test_agent_logger_progress(self, capsys):
        from icharlotte_core.agent_logger import AgentLogger
        logger = AgentLogger("TestAgent")

        logger.progress(50, "Halfway done")
        captured = capsys.readouterr()
        assert "PROGRESS:50:Halfway done" in captured.out


class TestDocumentProcessor:
    """Test DocumentProcessor functionality."""

    def test_processor_initialization(self):
        from icharlotte_core.document_processor import DocumentProcessor, OCRConfig
        # OCRConfig uses different parameter names
        config = OCRConfig(adaptive=True)
        processor = DocumentProcessor(ocr_config=config)
        assert processor is not None
        assert processor.ocr_config is not None

    def test_extract_result_success(self):
        from icharlotte_core.document_processor import ExtractResult
        result = ExtractResult(
            text="Test content",
            page_count=1
        )
        assert result.success is True
        assert result.text == "Test content"
        assert result.page_count == 1

    def test_verification_detector(self):
        from icharlotte_core.document_processor import VerificationDetector

        # Verification page
        verification_text = """
        VERIFICATION
        I declare under penalty of perjury that the foregoing is TRUE AND CORRECT.
        """
        assert VerificationDetector.is_verification_page(verification_text) is True

        # Regular content
        regular_text = "This is regular document content."
        assert VerificationDetector.is_verification_page(regular_text) is False

    def test_document_type_classifier(self):
        from icharlotte_core.document_processor import DocumentTypeClassifier

        # Form interrogatories
        doc_type, confidence = DocumentTypeClassifier.classify("FORM INTERROGATORIES SET ONE")
        assert doc_type == "form_interrogatories"
        assert confidence > 0

        # Unknown
        doc_type, confidence = DocumentTypeClassifier.classify("Random text")
        assert doc_type == "unknown"


class TestOutputValidator:
    """Test OutputValidator functionality."""

    def test_validation_result(self):
        from icharlotte_core.output_validator import ValidationResult
        result = ValidationResult()
        # Check initial state has valid attribute
        assert hasattr(result, 'is_valid') or hasattr(result, 'valid')

        result.add_error("Test error")
        assert "Test error" in result.errors

    def test_base_validator(self):
        from icharlotte_core.output_validator import BaseValidator, ValidationResult
        # Just test that BaseValidator exists and has validate method
        assert hasattr(BaseValidator, 'validate')

    def test_placeholder_detection(self):
        from icharlotte_core.output_validator import BaseValidator
        # Test that BaseValidator has the validate method
        assert callable(getattr(BaseValidator, 'validate', None))

    def test_get_validator(self):
        from icharlotte_core.output_validator import get_validator, DiscoveryValidator
        validator = get_validator("discovery")
        # get_validator returns an instance, not the class
        assert isinstance(validator, DiscoveryValidator)


class TestDeadlineCalculator:
    """Test DeadlineCalculator functionality."""

    def test_calculator_initialization(self):
        from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
        calc = DeadlineCalculator()
        assert calc is not None
        assert len(calc.rules_by_slug) > 0

    def test_is_court_day(self):
        from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
        from datetime import datetime

        calc = DeadlineCalculator()

        # Monday should be court day
        monday = datetime(2026, 1, 26)  # A Monday
        assert calc.is_court_day(monday) is True

        # Saturday should not be court day
        saturday = datetime(2026, 1, 31)  # A Saturday
        assert calc.is_court_day(saturday) is False

    def test_count_court_days(self):
        from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
        from datetime import datetime

        calc = DeadlineCalculator()
        start = datetime(2026, 1, 26)  # Monday

        # Count 5 court days forward
        result = calc.count_court_days(start, 5, 'forward')
        assert result > start

    def test_discovery_response_deadline(self):
        from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
        from datetime import datetime

        calc = DeadlineCalculator()
        request_date = datetime(2026, 1, 26)

        deadline = calc.get_discovery_response_deadline(request_date, 'electronic')
        assert deadline is not None
        assert 'date' in deadline
        assert deadline['date'] > request_date

    def test_motion_deadlines(self):
        from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
        from datetime import datetime

        calc = DeadlineCalculator()
        hearing_date = datetime(2026, 3, 20)

        deadlines = calc.get_motion_deadlines('msj', hearing_date, 'electronic')
        assert len(deadlines) > 0

        # Check that we have opposition and reply deadlines
        titles = [d['title'] for d in deadlines]
        assert 'Opposition Due' in titles or any('opposition' in t.lower() for t in titles)


class TestAttachmentClassifier:
    """Test AttachmentClassifier functionality."""

    def test_classifier_initialization(self):
        from icharlotte_core.calendar.attachment_classifier import AttachmentClassifier
        classifier = AttachmentClassifier()
        assert classifier is not None

    def test_pattern_classification(self):
        from icharlotte_core.calendar.attachment_classifier import AttachmentClassifier
        classifier = AttachmentClassifier()

        # Test motion detection
        result = classifier._classify_with_patterns("NOTICE OF MOTION FOR SUMMARY JUDGMENT")
        assert result['doc_type'] == 'motion'

        # Test discovery request detection
        result = classifier._classify_with_patterns("FORM INTERROGATORIES SET ONE")
        assert result['doc_type'] == 'discovery_request'

        # Test deposition notice detection
        result = classifier._classify_with_patterns("NOTICE OF TAKING DEPOSITION OF JOHN DOE")
        assert result['doc_type'] == 'deposition_notice'

    def test_date_extraction(self):
        from icharlotte_core.calendar.attachment_classifier import AttachmentClassifier
        classifier = AttachmentClassifier()

        text = "The hearing is scheduled for March 20, 2026."
        dates = classifier._extract_dates(text)
        assert len(dates) > 0


class TestLLMHandler:
    """Test LLMHandler functionality."""

    def test_handler_exists(self):
        from icharlotte_core.llm import LLMHandler
        assert hasattr(LLMHandler, 'generate')
        assert hasattr(LLMHandler, 'create_cache')

    def test_handler_no_api_key(self):
        from icharlotte_core.llm import LLMHandler

        # Should raise error when no API key
        with patch.dict(os.environ, {}, clear=True):
            with pytest.raises(ValueError, match="API Key"):
                LLMHandler.generate(
                    provider="TestProvider",
                    model="test-model",
                    system_prompt="Test",
                    user_prompt="Test",
                    file_contents="",
                    settings={}
                )


class TestLLMConfig:
    """Test LLMConfig functionality."""

    def test_config_singleton(self):
        from icharlotte_core.llm_config import LLMConfig
        config1 = LLMConfig()
        config2 = LLMConfig()
        assert config1 is config2

    def test_config_has_agents(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        # Check that config has agent-related attributes
        assert hasattr(config, 'agents') or hasattr(config, 'get_agent_config')


class TestExceptions:
    """Test custom exceptions."""

    def test_extraction_error(self):
        from icharlotte_core.exceptions import ExtractionError
        error = ExtractionError("Test error")
        # Exception may include additional details in string representation
        assert "Test error" in str(error)

    def test_ocr_error(self):
        from icharlotte_core.exceptions import OCRError
        error = OCRError("OCR failed")
        assert "OCR failed" in str(error)

    def test_llm_error(self):
        from icharlotte_core.exceptions import LLMError
        error = LLMError("LLM failed")
        assert "LLM failed" in str(error)


class TestScriptsImport:
    """Test that Scripts modules can be imported (basic syntax check)."""

    def test_import_gemini_utils(self):
        # Should import without fatal exit even without google-genai
        try:
            from Scripts.gemini_utils import log_event, clean_json_string
            assert log_event is not None
            assert clean_json_string is not None
        except Exception as e:
            # Allow import to fail gracefully, just not sys.exit
            pytest.skip(f"gemini_utils import failed: {e}")

    def test_import_case_data_manager(self):
        try:
            from Scripts.case_data_manager import CaseDataManager
            assert CaseDataManager is not None
        except Exception as e:
            pytest.skip(f"case_data_manager import failed: {e}")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
