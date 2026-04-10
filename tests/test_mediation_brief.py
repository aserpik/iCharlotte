import json
import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock


class TestStyleExtraction(unittest.TestCase):

    def test_extract_sections_from_text(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        sample_text = (
            "I.     INTRODUCTION\n"
            "This is the introduction paragraph.\n\n"
            "II.     STATEMENT OF FACTS\n"
            "These are the facts of the case.\n"
            "More facts here.\n\n"
            "III.     LIABILITY\n"
            "Liability arguments here.\n"
        )
        gen = MediationBriefGenerator.__new__(MediationBriefGenerator)
        sections = gen._extract_sections_from_text(sample_text)
        self.assertIn("introduction", sections)
        self.assertIn("statement_of_facts", sections)
        self.assertIn("liability", sections)
        self.assertIn("introduction paragraph", sections["introduction"])
        self.assertIn("facts of the case", sections["statement_of_facts"])

    def test_extract_sections_handles_missing_sections(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        sample_text = (
            "I.     INTRODUCTION\n"
            "Intro text.\n\n"
            "VII.     CONCLUSION\n"
            "Conclusion text.\n"
        )
        gen = MediationBriefGenerator.__new__(MediationBriefGenerator)
        sections = gen._extract_sections_from_text(sample_text)
        self.assertIn("introduction", sections)
        self.assertIn("conclusion", sections)
        self.assertNotIn("liability", sections)

    def test_cache_structure(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator.__new__(MediationBriefGenerator)
        gen._sample_dir = "C:\\AI\\Mediation Briefs"
        with patch.object(gen, '_read_sample_pdfs') as mock_read:
            mock_read.return_value = {
                "hashes": {"sample1.pdf": "abc123"},
                "sections": {
                    "introduction": ["Intro text from sample 1"],
                    "liability": ["Liability text from sample 1"],
                }
            }
            with tempfile.TemporaryDirectory() as tmpdir:
                cache_path = os.path.join(tmpdir, "style_cache.json")
                gen._cache_path = cache_path
                gen._save_style_cache(mock_read.return_value)
                with open(cache_path, 'r') as f:
                    cached = json.load(f)
                self.assertIn("source_hashes", cached)
                self.assertIn("sections", cached)
                self.assertEqual(cached["source_hashes"]["sample1.pdf"], "abc123")


if __name__ == '__main__':
    unittest.main()
