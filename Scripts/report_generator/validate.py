"""
Report Validator - Check generated reports against gold standard formatting rules.

Thin CLI wrapper around icharlotte_core.word_validator report checks.
All validation logic lives in the shared module.

Usage:
    # Standalone
    python -m Scripts.report_generator.validate path/to/report.docx
    python -m Scripts.report_generator.validate report.docx --verbose
    python -m Scripts.report_generator.validate report.docx --profile custom.json

    # In pipeline
    from Scripts.report_generator.validate import validate_report
    result = validate_report("path/to/report.docx")
    result.print_summary()
"""

import os
import sys
import json
import logging
import argparse
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from icharlotte_core.word_validator import (
    validate_report as _validate_report,
    Finding as _Finding,
    ValidationResult as _ValidationResult,
)

logger = logging.getLogger(__name__)

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DEFAULT_PROFILE_PATH = os.path.join(PROJECT_ROOT, "config", "report_reference_profile.json")


# ---------------------------------------------------------------------------
# Backward-compatible data classes
# These preserve the original API used by pipeline.py and CLI callers.
# Internally they wrap findings from the shared word_validator module.
# ---------------------------------------------------------------------------

@dataclass
class Finding:
    """A single validation finding (backward-compatible wrapper)."""
    level: str          # "PASS", "FAIL", "WARN"
    category: str       # Rule category
    message: str        # Human-readable description
    paragraph_index: Optional[int] = None
    expected: Any = None
    actual: Any = None

    def __str__(self):
        loc = f" (para {self.paragraph_index})" if self.paragraph_index is not None else ""
        return f"[{self.level}] {self.category}{loc}: {self.message}"


@dataclass
class ValidationResult:
    """Aggregated validation results (backward-compatible wrapper)."""
    doc_path: str
    findings: List[Finding] = field(default_factory=list)

    @property
    def pass_count(self):
        return sum(1 for f in self.findings if f.level == "PASS")

    @property
    def fail_count(self):
        return sum(1 for f in self.findings if f.level in ("FAIL", "ERROR"))

    @property
    def warn_count(self):
        return sum(1 for f in self.findings if f.level == "WARN")

    def print_summary(self, verbose=False):
        """Print results to console."""
        name = os.path.basename(self.doc_path)
        print(f"\n=== Report Validation: {name} ===")
        for f in self.findings:
            if verbose or f.level != "PASS":
                print(f"  {f}")
        print(f"\nResults: {self.pass_count} PASS, {self.fail_count} FAIL, {self.warn_count} WARN\n")


def _convert_finding(f: _Finding) -> Finding:
    """Convert a word_validator Finding to the backward-compatible Finding."""
    # Map severity: ERROR -> FAIL for backward compatibility
    level = "FAIL" if f.severity == "ERROR" else f.severity

    # Extract paragraph index from location string if present
    para_idx = None
    if f.location and f.location.startswith("para "):
        try:
            para_idx = int(f.location.split("para ")[1])
        except (ValueError, IndexError):
            pass

    return Finding(
        level=level,
        category=f.rule,
        message=f.message,
        paragraph_index=para_idx,
        expected=f.expected,
        actual=f.actual,
    )


# ---------------------------------------------------------------------------
# Profile loading
# ---------------------------------------------------------------------------

def load_profile(profile_path: str = None) -> Dict:
    """Load the reference profile JSON."""
    if profile_path is None:
        profile_path = DEFAULT_PROFILE_PATH
    if not os.path.exists(profile_path):
        logger.warning(f"Reference profile not found: {profile_path}")
        return {}
    with open(profile_path, 'r', encoding='utf-8') as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def validate_report(doc_path: str, profile_path: str = None) -> ValidationResult:
    """
    Validate a generated report against the reference profile.

    Delegates to icharlotte_core.word_validator.validate_report, then wraps
    the result in the backward-compatible ValidationResult format.

    Args:
        doc_path: Path to the .docx report to validate
        profile_path: Path to reference profile JSON (uses default if None)

    Returns:
        ValidationResult with all findings
    """
    # Load profile
    profile = load_profile(profile_path)

    # Delegate to shared module
    shared_result = _validate_report(doc_path, profile=profile)

    # Wrap in backward-compatible result
    result = ValidationResult(doc_path)
    for f in shared_result.findings:
        result.findings.append(_convert_finding(f))

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Validate report formatting")
    parser.add_argument("doc_path", help="Path to the .docx report to validate")
    parser.add_argument("--profile", default=None, help="Path to reference profile JSON")
    parser.add_argument("--verbose", "-v", action="store_true", help="Show PASS results too")
    args = parser.parse_args()

    result = validate_report(args.doc_path, args.profile)
    result.print_summary(verbose=args.verbose)

    # Exit with non-zero if any failures
    sys.exit(1 if result.fail_count > 0 else 0)


if __name__ == "__main__":
    main()
