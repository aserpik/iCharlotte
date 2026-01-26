"""
Tests for Output Validator Module

Tests validation of agent outputs.
"""

import os
import sys
import pytest

# Add parent to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core.output_validator import (
    ValidationResult, BaseValidator, DiscoveryValidator,
    DepositionValidator, SummaryValidator, TimelineValidator,
    get_validator, validate_output
)


class TestValidationResult:
    """Tests for ValidationResult class."""

    def test_initial_state(self):
        result = ValidationResult()
        assert result.valid
        assert result.score == 1.0
        assert len(result.errors) == 0
        assert len(result.warnings) == 0

    def test_add_error(self):
        result = ValidationResult()
        result.add_error("Test error")
        assert not result.valid
        assert len(result.errors) == 1
        assert result.score < 1.0

    def test_add_warning(self):
        result = ValidationResult()
        result.add_warning("Test warning")
        assert result.valid  # Still valid with warnings
        assert len(result.warnings) == 1
        assert result.score < 1.0

    def test_add_suggestion(self):
        result = ValidationResult()
        result.add_suggestion("Consider this")
        assert result.valid
        assert result.score == 1.0  # Suggestions don't affect score
        assert len(result.suggestions) == 1

    def test_merge_results(self):
        result1 = ValidationResult()
        result1.add_error("Error 1")

        result2 = ValidationResult()
        result2.add_warning("Warning 1")

        result1.merge(result2)
        assert not result1.valid
        assert len(result1.errors) == 1
        assert len(result1.warnings) == 1


class TestBaseValidator:
    """Tests for BaseValidator class."""

    def test_minimum_length_pass(self):
        validator = BaseValidator()
        result = validator.validate("A" * 200)
        assert result.valid

    def test_minimum_length_fail(self):
        validator = BaseValidator()
        result = validator.validate("Short")
        assert not result.valid
        assert any("too short" in e.lower() for e in result.errors)

    def test_placeholder_detection(self):
        validator = BaseValidator()
        result = validator.validate("This has a [PLACEHOLDER] in it. " * 10)
        assert len(result.warnings) > 0

    def test_todo_detection(self):
        validator = BaseValidator()
        result = validator.validate("This needs TODO: complete this section. " * 10)
        assert len(result.warnings) > 0

    def test_filler_detection(self):
        validator = BaseValidator()
        result = validator.validate("The parties include John, Jane, etc. " * 20)
        assert len(result.warnings) > 0


class TestDiscoveryValidator:
    """Tests for DiscoveryValidator class."""

    def test_valid_discovery(self):
        text = """
        ## General Information
        Plaintiff is John Smith, age 45.

        ## Incident Details
        The accident occurred on 01/15/2024 at approximately 3:00 PM.

        FROG No. 6.1: Plaintiff was traveling northbound when Defendant
        rear-ended his vehicle.
        """
        validator = DiscoveryValidator()
        result = validator.validate(text)
        # May have some warnings but should be generally valid
        assert result.score > 0.5

    def test_missing_dates(self):
        text = """
        ## General Information
        Plaintiff is John Smith.

        ## Incident Details
        The accident occurred in the morning.
        """ + " Additional content. " * 20

        validator = DiscoveryValidator()
        result = validator.validate(text)
        assert any("date" in s.lower() for s in result.suggestions)

    def test_missing_response_numbers(self):
        text = """
        ## General Information
        Plaintiff provided responses to discovery.

        ## Incident Details
        Details about the incident.
        """ + " Content. " * 30

        validator = DiscoveryValidator()
        result = validator.validate(text)
        assert any("response numbers" in w.lower() for w in result.warnings)


class TestDepositionValidator:
    """Tests for DepositionValidator class."""

    def test_valid_deposition(self):
        text = """
        Deposition of John Smith

        ## Personal Background
        - The witness testified that he has lived in California for 20 years
        - He stated that he works as an accountant
        - He indicated his education includes a bachelor's degree

        ## Incident Knowledge
        - John stated he saw the accident occur
        - He testified the defendant was speeding
        - The witness indicated he called 911

        ## Medical Treatment
        - He stated he received treatment at County Hospital
        - The witness testified about ongoing therapy
        - He indicated pain persists
        """
        validator = DepositionValidator()
        result = validator.validate(text)
        assert result.valid

    def test_missing_deponent(self):
        text = """
        ## Personal Background
        - The witness has lived in California for 20 years

        ## Incident Details
        - Details about the incident
        """ + " Content. " * 30

        validator = DepositionValidator()
        result = validator.validate(text)
        assert any("deponent" in e.lower() for e in result.errors)

    def test_few_topics(self):
        text = """
        Deposition of John Smith

        ## Background
        - Some testimony
        - More testimony
        """ + " Content. " * 30

        validator = DepositionValidator()
        result = validator.validate(text)
        assert any("topic" in w.lower() for w in result.warnings)


class TestSummaryValidator:
    """Tests for SummaryValidator class."""

    def test_valid_summary(self):
        text = """
        ## Overview
        This document summarizes the key facts from the police report dated 01/15/2024.

        ## Details
        The incident involved 2 vehicles at the intersection of Main St and 5th Ave.
        Damage was estimated at $15,000.
        """
        validator = SummaryValidator()
        result = validator.validate(text)
        assert result.valid

    def test_no_numbers(self):
        text = """
        ## Overview
        This is a summary of the document.

        ## Details
        Various things happened in this case. More details follow.
        """ + " Additional content here. " * 20

        validator = SummaryValidator()
        result = validator.validate(text)
        assert any("number" in w.lower() for w in result.warnings)


class TestTimelineValidator:
    """Tests for TimelineValidator class."""

    def test_valid_timeline(self):
        text = """
        {"events": [
            {"date": "2024-01-15", "event": "Incident occurred"},
            {"date": "2024-01-16", "event": "ER visit"},
            {"date": "2024-02-01", "event": "Follow-up appointment"}
        ]}
        """
        validator = TimelineValidator()
        result = validator.validate(text)
        assert result.valid

    def test_few_dates(self):
        text = """
        2024-01-15: Something happened
        """ + " More content. " * 20

        validator = TimelineValidator()
        result = validator.validate(text)
        assert any("date" in w.lower() for w in result.warnings)


class TestGetValidator:
    """Tests for validator factory function."""

    def test_get_discovery_validator(self):
        validator = get_validator("discovery")
        assert isinstance(validator, DiscoveryValidator)

    def test_get_deposition_validator(self):
        validator = get_validator("deposition")
        assert isinstance(validator, DepositionValidator)

    def test_get_summary_validator(self):
        validator = get_validator("summary")
        assert isinstance(validator, SummaryValidator)

    def test_get_unknown_type(self):
        validator = get_validator("unknown_type")
        assert isinstance(validator, BaseValidator)

    def test_case_insensitive(self):
        validator = get_validator("DISCOVERY")
        assert isinstance(validator, DiscoveryValidator)


class TestValidateOutput:
    """Tests for validate_output convenience function."""

    def test_validate_output(self):
        result = validate_output("A" * 200, "summary")
        assert isinstance(result, ValidationResult)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
