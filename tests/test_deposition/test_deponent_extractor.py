"""Tests for DeponentExtractor.extract_deponent_name.

Regression coverage for the Forney depo bug: caption uses U+00B7 middle-dot
leaders between OF and the name, and Zoom boilerplate "the witness is being
spotlighted or locked on..." appears earlier in the document. Previously the
extractor matched the boilerplate (because re.IGNORECASE turned the
[A-Z][A-Za-z]+ anchors into any-letter classes) and returned
"is being spotlighted or locked on".
"""

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "Scripts"))

from Scripts.summarize_deposition import DeponentExtractor


class TestExtractDeponentName:
    def test_simple_caption(self):
        text = "DEPOSITION OF JOHN SMITH\nTaken on January 15, 2024"
        assert DeponentExtractor.extract_deponent_name(text) == "JOHN SMITH"

    def test_caption_with_credential_suffix(self):
        text = "DEPOSITION OF JOHN SMITH, M.D.\nTaken on January 15, 2024"
        # Existing behaviour: name + suffix were captured together.
        name = DeponentExtractor.extract_deponent_name(text)
        assert name.startswith("JOHN SMITH")

    def test_caption_with_middle_dot_leaders(self):
        """Real Forney transcript: middle-dot leaders between OF and name."""
        text = (
            "DEPOSITION OF\n"
            "· · · · · · · · · · PAMELA FORNEY\n"
        )
        assert DeponentExtractor.extract_deponent_name(text) == "PAMELA FORNEY"

    def test_zoom_boilerplate_does_not_match_witness_pattern(self):
        """The 'witness is being spotlighted' Zoom phrase must not be captured."""
        text = (
            "For the purpose of creating a witness-only video recording, "
            "the witness is being spotlighted or locked on all video screens "
            "while in speaker view.\n"
        )
        # No caption-style match; should fall through to empty.
        assert DeponentExtractor.extract_deponent_name(text) == ""

    def test_forney_real_world(self):
        """Realistic mix: Zoom boilerplate first, then caption with middle dots."""
        text = (
            "the witness is being spotlighted or locked on all video screens\n"
            "while in speaker view.  We ask that the witness not remove the\n"
            "spotlight setting until the deposition is concluded.\n"
            "\n"
            "DEPOSITION OF\n"
            "· · · · · · · · PAMELA FORNEY\n"
            "Taken on September 24, 2025\n"
        )
        assert DeponentExtractor.extract_deponent_name(text) == "PAMELA FORNEY"

    def test_deponent_label(self):
        text = "Deponent: Jane Doe\nDate: 1/15/2024"
        assert DeponentExtractor.extract_deponent_name(text) == "Jane Doe"

    def test_witness_caption_all_caps_still_works(self):
        text = "WITNESS: MARY JOHNSON\nTaken on 1/15/2024"
        assert DeponentExtractor.extract_deponent_name(text) == "MARY JOHNSON"

    def test_witness_lowercase_in_prose_is_rejected(self):
        """Lowercase 'witness' in narrative must NOT match the WITNESS caption pattern."""
        text = "The witness then stated that she was not feeling well that day.\n"
        assert DeponentExtractor.extract_deponent_name(text) == ""

    def test_videotaped_caption(self):
        text = "THE VIDEOTAPED DEPOSITION OF ROBERT BROWN\nTaken on 1/15/2024"
        assert DeponentExtractor.extract_deponent_name(text) == "ROBERT BROWN"
