"""Discovery type detection helpers for the Respond to Discovery wizard."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Iterable, Optional


VALID_DISCOVERY_TYPES = {"FI", "SI", "RFA", "RPD"}


@dataclass(frozen=True)
class DiscoveryTypeGuess:
    discovery_type: Optional[str]
    source: str
    matched_text: str = ""


@dataclass(frozen=True)
class DiscoveryTypeResolution:
    discovery_type: Optional[str]
    needs_user_choice: bool
    reason: str = ""


_FILENAME_PATTERNS = (
    ("FI", re.compile(r"\b(fro+g+s?|form\s+interrog(?:atory|atories)?)\b", re.I)),
    ("SI", re.compile(r"\b(sro+g+s?|special\s+interrog(?:atory|atories)?)\b", re.I)),
    ("RFA", re.compile(r"\b(rfa|requests?\s+for\s+admission)\b", re.I)),
    ("RPD", re.compile(r"\b(rfp|rpd|requests?\s+for\s+production)\b", re.I)),
)

_TEXT_PATTERNS = (
    ("FI", re.compile(r"\bform\s+interrogator(?:y|ies)\b", re.I)),
    ("SI", re.compile(r"\bspecial\s+interrogator(?:y|ies)\b", re.I)),
    ("RFA", re.compile(r"\brequests?\s+for\s+admission\b", re.I)),
    ("RPD", re.compile(r"\brequests?\s+for\s+production(?:\s+of\s+documents)?\b", re.I)),
)

_NORMALIZE_PATTERNS = (
    ("FI", re.compile(r"\b(fi|fro+g+s?|form\s+interrogator(?:y|ies))\b", re.I)),
    ("SI", re.compile(r"\b(si|sro+g+s?|special\s+interrogator(?:y|ies))\b", re.I)),
    ("RFA", re.compile(r"\b(rfa|requests?\s+for\s+admission)\b", re.I)),
    ("RPD", re.compile(r"\b(rfp|rpd|requests?\s+for\s+production)\b", re.I)),
)


def _detect(
    value: str,
    source: str,
    patterns: Iterable[tuple[str, re.Pattern[str]]],
) -> DiscoveryTypeGuess:
    text = value or ""
    for discovery_type, pattern in patterns:
        match = pattern.search(text)
        if match:
            return DiscoveryTypeGuess(discovery_type, source, match.group(0))
    return DiscoveryTypeGuess(None, source, "")


def detect_type_from_filename(path_or_name: str) -> DiscoveryTypeGuess:
    """Detect discovery type from common filename abbreviations."""
    name = re.sub(r"[_\-]+", " ", os.path.basename(path_or_name or ""))
    return _detect(name, "filename", _FILENAME_PATTERNS)


def detect_type_from_text(text: str) -> DiscoveryTypeGuess:
    """Detect discovery type from document text, usually first-page text."""
    return _detect(text or "", "text", _TEXT_PATTERNS)


def normalize_discovery_type(value: str | None) -> str:
    """Normalize common discovery type labels to FI/SI/RFA/RPD."""
    raw = (value or "").strip()
    upper = raw.upper()
    if upper in VALID_DISCOVERY_TYPES:
        return upper
    for discovery_type, pattern in _NORMALIZE_PATTERNS:
        if pattern.search(raw):
            return discovery_type
    return ""


def resolve_detected_type(
    filename_guess: DiscoveryTypeGuess,
    text_guess: DiscoveryTypeGuess,
) -> DiscoveryTypeResolution:
    """Resolve filename/text guesses into an automatic type or user choice."""
    filename_type = filename_guess.discovery_type
    text_type = text_guess.discovery_type

    if filename_type and text_type and filename_type != text_type:
        return DiscoveryTypeResolution(
            None,
            True,
            f"Detection conflict: filename={filename_type}, text={text_type}",
        )
    if filename_type:
        return DiscoveryTypeResolution(filename_type, False, "Detected from filename")
    if text_type:
        return DiscoveryTypeResolution(text_type, False, "Detected from text")
    return DiscoveryTypeResolution(None, True, "Discovery type could not be detected")
