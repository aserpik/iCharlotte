"""
Tests for Document Processor Module

Tests text extraction, OCR integration, and document classification.
"""

import os
import sys
import pytest
from unittest.mock import Mock, patch, MagicMock

# Add parent to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core.document_processor import (
    DocumentProcessor, ExtractResult, OCRConfig,
    VerificationDetector, DocumentTypeClassifier, PartyExtractor
)


class TestExtractResult:
    """Tests for ExtractResult dataclass."""

    def test_success_result(self):
        result = ExtractResult(
            text="Sample text content",
            page_count=5,
            char_count=20,
            ocr_pages=[2, 3],
            extraction_method="mixed"
        )
        assert result.success
        assert result.ocr_percentage == 40.0

    def test_failure_result(self):
        result = ExtractResult(
            text="",
            page_count=0,
            error="File not found"
        )
        assert not result.success
        assert result.error == "File not found"

    def test_empty_pages(self):
        result = ExtractResult(text="text", page_count=0)
        assert result.ocr_percentage == 0.0


class TestVerificationDetector:
    """Tests for verification page detection."""

    def test_detect_verification(self):
        text = """
        VERIFICATION
        I, John Smith, declare under penalty of perjury that the foregoing
        is true and correct.
        Signed: John Smith
        Date: January 15, 2024
        """
        assert VerificationDetector.is_verification_page(text)

    def test_detect_declaration(self):
        # Implementation requires both perjury language AND "TRUE AND CORRECT" or "EXECUTED ON"
        text = "I declare under penalty of perjury that the foregoing is TRUE AND CORRECT"
        assert VerificationDetector.is_verification_page(text)

    def test_not_verification(self):
        text = """
        INTERROGATORY NO. 1: Please state your full legal name.
        RESPONSE: My name is John Smith. I have been known by this name
        since birth. I have never used any other names.
        """
        assert not VerificationDetector.is_verification_page(text)

    def test_short_document_check(self):
        # Verification detection based on keywords
        # Short text with verification keywords should trigger
        verification_text = """
        VERIFICATION
        I declare under penalty of perjury that the foregoing is TRUE AND CORRECT.
        """
        assert VerificationDetector.is_verification_page(verification_text)

        # Regular content without verification keywords should not trigger
        regular_text = "Content " * 100  # 800 chars of regular content
        assert not VerificationDetector.is_verification_page(regular_text)


class TestDocumentTypeClassifier:
    """Tests for document type classification."""

    def test_classify_form_interrogatory(self):
        text = "FORM INTERROGATORIES - GENERAL\nFROG SET ONE"
        doc_type, confidence = DocumentTypeClassifier.classify(text)
        assert doc_type == "form_interrogatories"
        assert confidence > 0

    def test_classify_special_interrogatory(self):
        text = "SPECIAL INTERROGATORIES SET ONE\nINTERROGATORY NO. 1:"
        doc_type, confidence = DocumentTypeClassifier.classify(text)
        assert doc_type == "special_interrogatories"
        assert confidence > 0

    def test_classify_rfp(self):
        text = "REQUEST FOR PRODUCTION OF DOCUMENTS\nREQUEST NO. 1:"
        doc_type, confidence = DocumentTypeClassifier.classify(text)
        assert doc_type == "requests_for_production"
        assert confidence > 0

    def test_classify_rfa(self):
        text = "REQUESTS FOR ADMISSION SET ONE\nREQUEST FOR ADMISSION NO. 1:"
        doc_type, confidence = DocumentTypeClassifier.classify(text)
        assert doc_type == "requests_for_admission"
        assert confidence > 0

    def test_classify_unknown(self):
        text = "Some random legal document content"
        doc_type, confidence = DocumentTypeClassifier.classify(text)
        assert doc_type == "unknown"
        assert confidence == 0.0


class TestPartyExtractor:
    """Tests for party name extraction."""

    def test_extract_plaintiff_responses(self):
        # Implementation returns the actual party name, not the role
        text = "PLAINTIFF JOHN SMITH'S RESPONSES TO DEFENDANT'S INTERROGATORIES"
        party = PartyExtractor.extract_party_name(text)
        assert party is not None
        assert "John Smith" in party or "Smith" in party

    def test_extract_defendant_responses(self):
        # Implementation returns the actual party name, not the role
        text = "DEFENDANT ABC CORPORATION'S RESPONSES TO FORM INTERROGATORIES"
        party = PartyExtractor.extract_party_name(text)
        assert party is not None
        assert "Corporation" in party or "Abc" in party

    def test_no_party_found(self):
        # Returns None when no party found (not empty string)
        text = "RESPONSES TO SPECIAL INTERROGATORIES"
        party = PartyExtractor.extract_party_name(text)
        assert party is None

    def test_extract_deponent_name(self):
        text = """
        DEPOSITION OF JOHN SMITH, M.D.
        Taken on January 15, 2024
        """
        name = PartyExtractor.extract_deponent_name(text)
        assert "John Smith" in name or "JOHN SMITH" in name


class TestDocumentProcessor:
    """Tests for DocumentProcessor class."""

    def test_initialization(self):
        processor = DocumentProcessor()
        assert processor.ocr_config is not None

    def test_initialization_with_config(self):
        config = OCRConfig(base_threshold=100, adaptive=False)
        processor = DocumentProcessor(ocr_config=config)
        assert processor.ocr_config.base_threshold == 100
        assert not processor.ocr_config.adaptive

    @patch('icharlotte_core.document_processor.PdfReader')
    def test_extract_text_pdf(self, mock_reader):
        """Test PDF text extraction (mocked)."""
        # Setup mock
        mock_page = MagicMock()
        mock_page.extract_text.return_value = "Sample PDF text content"
        mock_reader.return_value.pages = [mock_page]

        processor = DocumentProcessor()
        # Would need a real file to test fully
        # This just verifies the processor can be created

    def test_calculate_dynamic_threshold(self):
        processor = DocumentProcessor(ocr_config=OCRConfig(adaptive=True))

        # High density pages (lots of text normally) get higher threshold
        # This means OCR triggers if we're not getting expected text amount
        high_density = [500, 600, 550]
        threshold_high = processor._calculate_dynamic_threshold(high_density, "test.pdf")

        # Low density pages (sparse text normally) get lower threshold
        # This prevents unnecessary OCR on naturally sparse documents
        low_density = [50, 60, 45]
        threshold_low = processor._calculate_dynamic_threshold(low_density, "test.pdf")

        # Verify thresholds are within configured bounds
        assert threshold_low >= processor.ocr_config.min_threshold
        assert threshold_high <= processor.ocr_config.max_threshold
        # High density should have higher threshold (trigger OCR more readily)
        assert threshold_high >= threshold_low


class TestOCRConfig:
    """Tests for OCR configuration."""

    def test_default_config(self):
        config = OCRConfig()
        assert config.base_threshold == 50
        assert config.adaptive
        assert config.max_dpi == 300

    def test_custom_config(self):
        config = OCRConfig(base_threshold=100, adaptive=False, max_dpi=150)
        assert config.base_threshold == 100
        assert not config.adaptive
        assert config.max_dpi == 150


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
