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


class TestCaptionHandling(unittest.TestCase):

    def _make_caption_doc(self, tmpdir, include_sig_block=False):
        from docx import Document
        doc = Document()
        doc.add_paragraph("LAW FIRM NAME")
        doc.add_paragraph("CAPTION PAGE")
        doc.add_paragraph("Some caption content")
        if include_sig_block:
            doc.add_paragraph("")
            doc.add_paragraph("DATED: April 10, 2026")
            doc.add_paragraph("")
            doc.add_paragraph("By: ____________________")
            doc.add_paragraph("John Smith, Esq.")
            doc.add_paragraph("State Bar No. 123456")
        path = os.path.join(tmpdir, "case_caption.docx")
        doc.save(path)
        return path

    def test_find_caption_in_folder(self):
        from docx import Document
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            self._make_caption_doc(tmpdir)
            doc2 = Document()
            doc2.add_paragraph("Not a caption")
            doc2.save(os.path.join(tmpdir, "other_doc.docx"))
            result = gen.find_caption_template(tmpdir)
            self.assertIsNotNone(result)
            self.assertIn("caption", os.path.basename(result).lower())

    def test_find_caption_returns_none_when_missing(self):
        from docx import Document
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            doc = Document()
            doc.add_paragraph("Not a caption")
            doc.save(os.path.join(tmpdir, "other.docx"))
            result = gen.find_caption_template(tmpdir)
            self.assertIsNone(result)

    def test_replace_caption_page_text(self):
        from docx import Document
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            caption_path = self._make_caption_doc(tmpdir)
            output_path = os.path.join(tmpdir, "output.docx")
            gen.prepare_caption_template(caption_path, output_path)
            doc = Document(output_path)
            all_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertNotIn("CAPTION PAGE", all_text)
            self.assertIn("DEFENDANT'S CONFIDENTIAL MEDIATION BRIEF", all_text)

    def test_signature_block_detection(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            caption_path = self._make_caption_doc(tmpdir, include_sig_block=True)
            output_path = os.path.join(tmpdir, "output.docx")
            sig_paras = gen.prepare_caption_template(caption_path, output_path)
            self.assertTrue(len(sig_paras) > 0)
            sig_text = " ".join(p.text for p in sig_paras)
            self.assertIn("DATED", sig_text)

    def test_no_signature_block(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            caption_path = self._make_caption_doc(tmpdir, include_sig_block=False)
            output_path = os.path.join(tmpdir, "output.docx")
            sig_paras = gen.prepare_caption_template(caption_path, output_path)
            self.assertEqual(len(sig_paras), 0)


if __name__ == '__main__':
    unittest.main()
