"""Extract case numbers (####.###) from various text sources."""

import re

# Matches case numbers like 3850.084, 3850-084, 3850_084, 3850\u2013084, 3850\u2014084
CASE_NUMBER_REGEX = r'(\d{4})[._\-\u2013\u2014](\d{3})'


def extract_case_number(text: str) -> str | None:
    """Extract first case number from text, normalized to ####.### format."""
    if not text:
        return None
    match = re.search(CASE_NUMBER_REGEX, text)
    if match:
        return f"{match.group(1)}.{match.group(2)}"
    return None


def extract_from_multiple(*sources: str) -> str | None:
    """Try extraction from multiple sources, return first hit."""
    for source in sources:
        result = extract_case_number(source)
        if result:
            return result
    return None
