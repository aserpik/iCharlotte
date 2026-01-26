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
        text = "I declare under penalty of perjury under the laws of the State of California"
        assert VerificationDetector.is_verification_page(text)

    def test_not_verification(self):
        text = """
        INTERROGATORY NO. 1: Please state your full legal name.
        RESPONSE: My name is John Smith. I have been known by this name
        since birth. I have never used any other names.
        """
        assert not VerificationDetector.is_verification_page(text)

    def test_short_document_check(self):
        # Verification should only apply to short documents
        long_text = "Content " * 1000  # 8000 chars
        short_text = "Content " * 100  # 800 chars

        # Even with verification keywords, long docs shouldn't trigger
        result = VerificationDetector.has_response_content(long_text)
        assert result  # Long doc with content


class TestDocumentTypeClassifier:
    """Tests for document type classification."""

    def test_classify_form_interrogatory(self):
        text = "FORM INTERROGATORIES - GENERAL\nFROG SET ONE"
        doc_type = DocumentTypeClassifier.classify(text)
        assert doc_type == "Form Interrogatories"

    def test_classify_special_interrogatory(self):
        text = "SPECIAL INTERROGATORIES SET ONE\nINTERROGATORY NO. 1:"
        doc_type = DocumentTypeClassifier.classify(text)
        assert doc_type == "Special Interrogatories"

    def test_classify_rfp(self):
        text = "REQUEST FOR PRODUCTION OF DOCUMENTS\nREQUEST NO. 1:"
        doc_type = DocumentTypeClassifier.classify(text)
        assert doc_type == "Request for Production"

    def test_classify_rfa(self):
        text = "REQUESTS FOR ADMISSION SET ONE\nREQUEST FOR ADMISSION NO. 1:"
        doc_type = DocumentTypeClassifier.classify(text)
        assert doc_type == "Requests for Admission"

    def test_classify_unknown(self):
        text = "Some random legal document content"
        doc_type = DocumentTypeClassifier.classify(text)
        assert doc_type == "Unknown"


class TestPartyExtractor:
    """Tests for party name extraction."""

    def test_extract_plaintiff_responses(self):
        text = "PLAINTIFF JOHN SMITH'S RESPONSES TO DEFENDANT'S INTERROGATORIES"
        party = PartyExtractor.extract_party_name(text)
        assert party == "Plaintiff"

    def test_extract_defendant_responses(self):
        text = "DEFENDANT ABC CORPORATION'S RESPONSES TO FORM INTERROGATORIES"
        party = PartyExtractor.extract_party_name(text)
        assert party == "Defendant"

    def test_no_party_found(self):
        text = "RESPONSES TO SPECIAL INTERROGATORIES"
        party = PartyExtractor.extract_party_name(text)
        assert party == ""

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

        # High density pages should have lower threshold
        high_density = [500, 600, 550]
        threshold_high = processor._calculate_dynamic_threshold(high_density, "test.pdf")

        # Low density pages should have higher threshold
        low_density = [50, 60, 45]
        threshold_low = processor._calculate_dynamic_threshold(low_density, "test.pdf")

        # Higher threshold for low density (more likely to need OCR)
        assert threshold_low >= threshold_high


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
