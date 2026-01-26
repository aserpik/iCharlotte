"""
Output Validator for iCharlotte

Validates agent outputs to ensure quality and completeness.

Features:
- Schema validation for different output types
- Required section checking
- Minimum content length validation
- Placeholder detection
- Quality scoring
"""

import re
from typing import List, Optional, Dict, Any
from dataclasses import dataclass, field


# =============================================================================
# Validation Results
# =============================================================================

@dataclass
class ValidationResult:
    """Result of validating an output."""
    valid: bool = True
    score: float = 1.0  # 0.0 to 1.0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    suggestions: List[str] = field(default_factory=list)

    def add_error(self, message: str):
        """Add an error and mark as invalid."""
        self.errors.append(message)
        self.valid = False
        self.score = max(0.0, self.score - 0.2)

    def add_warning(self, message: str):
        """Add a warning."""
        self.warnings.append(message)
        self.score = max(0.0, self.score - 0.1)

    def add_suggestion(self, message: str):
        """Add a suggestion for improvement."""
        self.suggestions.append(message)

    def merge(self, other: 'ValidationResult'):
        """Merge another result into this one."""
        self.valid = self.valid and other.valid
        self.errors.extend(other.errors)
        self.warnings.extend(other.warnings)
        self.suggestions.extend(other.suggestions)
        self.score = min(self.score, other.score)


# =============================================================================
# Base Validator
# =============================================================================

class BaseValidator:
    """Base class for output validators."""

    # Placeholder patterns to detect
    PLACEHOLDER_PATTERNS = [
        r'\[.*?\]',  # [placeholder]
        r'\{.*?\}',  # {placeholder} (not JSON)
        r'<.*?>',    # <placeholder>
        r'INSERT\s+\w+\s+HERE',
        r'TODO:?\s*\w+',
        r'PLACEHOLDER',
        r'XXX+',
        r'\?\?\?+',
    ]

    # Common filler phrases that suggest incomplete content
    FILLER_PATTERNS = [
        r'etc\.?\s*$',
        r'and so on\s*\.?$',
        r'more details needed',
        r'to be determined',
        r'TBD',
        r'N/A' * 3,  # Multiple N/A in a row
    ]

    def __init__(self):
        self.result = ValidationResult()

    def validate(self, text: str) -> ValidationResult:
        """
        Validate the output text.

        Args:
            text: The output text to validate.

        Returns:
            ValidationResult with errors, warnings, and score.
        """
        self.result = ValidationResult()

        # Run all validation checks
        self._check_minimum_length(text)
        self._check_placeholders(text)
        self._check_filler_phrases(text)
        self._check_basic_structure(text)

        return self.result

    def _check_minimum_length(self, text: str, min_chars: int = 100):
        """Check minimum content length."""
        if len(text) < min_chars:
            self.result.add_error(f"Output too short ({len(text)} chars, minimum {min_chars})")

    def _check_placeholders(self, text: str):
        """Check for placeholder text."""
        for pattern in self.PLACEHOLDER_PATTERNS:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                # Ignore if it looks like actual content
                if len(match) > 50:  # Probably actual content in brackets
                    continue
                if match.startswith('{') and ':' in match:  # Might be JSON
                    continue
                self.result.add_warning(f"Possible placeholder detected: {match[:50]}")

    def _check_filler_phrases(self, text: str):
        """Check for filler phrases suggesting incomplete content."""
        for pattern in self.FILLER_PATTERNS:
            if re.search(pattern, text, re.IGNORECASE):
                self.result.add_warning(f"Filler phrase detected: pattern '{pattern}'")

    def _check_basic_structure(self, text: str):
        """Check basic structural requirements."""
        lines = text.strip().split('\n')

        if len(lines) < 3:
            self.result.add_warning("Very few lines in output")

        # Check for some paragraph structure
        paragraphs = [p for p in text.split('\n\n') if p.strip()]
        if len(paragraphs) < 2 and len(text) > 500:
            self.result.add_suggestion("Consider adding paragraph breaks for readability")


# =============================================================================
# Discovery Validator
# =============================================================================

class DiscoveryValidator(BaseValidator):
    """Validator for discovery summary outputs."""

    # Required sections for discovery summaries
    REQUIRED_SECTIONS = [
        (r'##?\s*(?:general\s*)?information', 'General Information section'),
        (r'##?\s*(?:incident|accident)', 'Incident/Accident section'),
    ]

    # Expected patterns
    EXPECTED_PATTERNS = [
        (r'\d{1,2}/\d{1,2}/\d{2,4}', 'at least one date'),
        (r'[Pp]laintiff|[Dd]efendant', 'party references'),
    ]

    def validate(self, text: str) -> ValidationResult:
        """Validate discovery summary output."""
        # Run base validations
        super().validate(text)

        # Check required sections
        for pattern, name in self.REQUIRED_SECTIONS:
            if not re.search(pattern, text, re.IGNORECASE):
                self.result.add_warning(f"Missing or unclear: {name}")

        # Check expected patterns
        for pattern, description in self.EXPECTED_PATTERNS:
            if not re.search(pattern, text):
                self.result.add_suggestion(f"Consider including {description}")

        # Check for response numbers
        if not re.search(r'(?:FROG|SROG|RFA|RFP|RPD)\s*(?:No\.?\s*)?\d+', text, re.IGNORECASE):
            self.result.add_warning("No discovery response numbers found")

        return self.result


# =============================================================================
# Deposition Validator
# =============================================================================

class DepositionValidator(BaseValidator):
    """Validator for deposition summary outputs."""

    REQUIRED_ELEMENTS = [
        (r'deposition\s+of\s+\w+', 'Deponent identification'),
        (r'##?\s*\w+', 'Topic headings'),
    ]

    def validate(self, text: str) -> ValidationResult:
        """Validate deposition summary output."""
        super().validate(text)

        # Check for deponent info
        if not re.search(r'deposition\s+of', text, re.IGNORECASE):
            self.result.add_error("Missing deponent identification")

        # Check for topic structure
        headings = re.findall(r'^##?\s+(.+)$', text, re.MULTILINE)
        if len(headings) < 3:
            self.result.add_warning(f"Only {len(headings)} topic headings found (expected 5+)")

        # Check for bullet points under topics
        bullets = re.findall(r'^[\-\*]\s+', text, re.MULTILINE)
        if len(bullets) < 10:
            self.result.add_warning("Few bullet points found - may lack detail")

        # Check for testimony quotes or paraphrasing
        if not re.search(r'(?:testified|stated|said|indicated)', text, re.IGNORECASE):
            self.result.add_suggestion("Consider including testimony attribution")

        return self.result


# =============================================================================
# Summary Validator
# =============================================================================

class SummaryValidator(BaseValidator):
    """Validator for general document summaries."""

    def validate(self, text: str) -> ValidationResult:
        """Validate general summary output."""
        super().validate(text)

        # Minimum length for summaries
        self._check_minimum_length(text, min_chars=200)

        # Check for structure
        if not re.search(r'^##?\s+', text, re.MULTILINE):
            self.result.add_suggestion("Consider adding section headings")

        # Check for factual content indicators
        if not re.search(r'\d', text):
            self.result.add_warning("No numbers found - may be missing key facts")

        return self.result


# =============================================================================
# Timeline Validator
# =============================================================================

class TimelineValidator(BaseValidator):
    """Validator for timeline extraction outputs."""

    def validate(self, text: str) -> ValidationResult:
        """Validate timeline output."""
        super().validate(text)

        # Check for date format
        dates = re.findall(r'\d{4}-\d{2}-\d{2}', text)
        if len(dates) < 3:
            self.result.add_warning(f"Only {len(dates)} dates found - may be incomplete")

        # Check for event descriptions
        if '{"events"' not in text and '"events"' not in text:
            # Not JSON format - check for prose timeline
            if not re.search(r'\d{4}', text):
                self.result.add_error("No year references found in timeline")

        return self.result


# =============================================================================
# Contradiction Validator
# =============================================================================

class ContradictionValidator(BaseValidator):
    """Validator for contradiction report outputs."""

    def validate(self, text: str) -> ValidationResult:
        """Validate contradiction report output."""
        super().validate(text)

        # Check for contradiction structure
        if '"contradictions"' in text:
            # JSON format
            if '"claim_1"' not in text or '"claim_2"' not in text:
                self.result.add_warning("Contradiction structure may be incomplete")
        else:
            # Prose format
            if not re.search(r'(?:conflict|contradict|inconsisten)', text, re.IGNORECASE):
                self.result.add_warning("May not contain contradiction analysis")

        return self.result


# =============================================================================
# Factory Function
# =============================================================================

def get_validator(output_type: str) -> BaseValidator:
    """
    Get the appropriate validator for an output type.

    Args:
        output_type: Type of output ("discovery", "deposition", "summary", etc.)

    Returns:
        Appropriate validator instance.
    """
    validators = {
        'discovery': DiscoveryValidator,
        'deposition': DepositionValidator,
        'summary': SummaryValidator,
        'timeline': TimelineValidator,
        'contradiction': ContradictionValidator,
    }

    validator_class = validators.get(output_type.lower(), BaseValidator)
    return validator_class()


def validate_output(text: str, output_type: str) -> ValidationResult:
    """
    Convenience function to validate output.

    Args:
        text: Output text to validate.
        output_type: Type of output.

    Returns:
        ValidationResult.
    """
    validator = get_validator(output_type)
    return validator.validate(text)


# =============================================================================
# Batch Validation
# =============================================================================

def validate_batch(outputs: Dict[str, str], output_type: str) -> Dict[str, ValidationResult]:
    """
    Validate multiple outputs.

    Args:
        outputs: Dictionary of name -> text.
        output_type: Type of outputs.

    Returns:
        Dictionary of name -> ValidationResult.
    """
    results = {}
    validator = get_validator(output_type)

    for name, text in outputs.items():
        results[name] = validator.validate(text)

    return results


def get_batch_summary(results: Dict[str, ValidationResult]) -> Dict[str, Any]:
    """
    Get summary statistics for batch validation.

    Args:
        results: Dictionary of validation results.

    Returns:
        Summary statistics.
    """
    total = len(results)
    valid_count = sum(1 for r in results.values() if r.valid)
    avg_score = sum(r.score for r in results.values()) / total if total > 0 else 0

    all_errors = []
    all_warnings = []

    for name, result in results.items():
        for error in result.errors:
            all_errors.append(f"{name}: {error}")
        for warning in result.warnings:
            all_warnings.append(f"{name}: {warning}")

    return {
        'total': total,
        'valid': valid_count,
        'invalid': total - valid_count,
        'average_score': avg_score,
        'error_count': len(all_errors),
        'warning_count': len(all_warnings),
        'errors': all_errors[:10],  # First 10 errors
        'warnings': all_warnings[:10],  # First 10 warnings
    }
