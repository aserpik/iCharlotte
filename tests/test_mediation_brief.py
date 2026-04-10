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


class TestSectionGeneration(unittest.TestCase):

    def test_build_section_prompt_includes_planning_output(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.planning_output = "KEY FACTS:\n- Accident on Jan 1, 2025"
        gen.sections = {}
        gen.document_content = "Document text here"
        gen._style_cache = {"sections": {}}
        prompt = gen._build_section_prompt("statement_of_facts")
        self.assertIn("Accident on Jan 1, 2025", prompt)

    def test_build_section_prompt_includes_prior_sections(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.planning_output = "Planning data"
        gen.sections = {
            "statement_of_facts": "The plaintiff was injured on Main St.",
            "procedural_status": "Trial is set for June 2027.",
        }
        gen.document_content = "Doc text"
        gen._style_cache = {"sections": {}}
        prompt = gen._build_section_prompt("liability")
        self.assertIn("injured on Main St", prompt)
        self.assertIn("Trial is set for June 2027", prompt)

    def test_build_introduction_prompt_includes_all_sections(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.planning_output = "Planning data"
        gen.sections = {
            "statement_of_facts": "Facts text",
            "procedural_status": "Status text",
            "liability": "Liability text",
            "damages": "Damages text",
            "settlement_position": "Settlement text",
            "conclusion": "Conclusion text",
        }
        gen.document_content = "Doc text"
        gen._style_cache = {"sections": {}}
        prompt = gen._build_section_prompt("introduction")
        self.assertIn("Liability text", prompt)
        self.assertIn("Damages text", prompt)
        self.assertIn("Conclusion text", prompt)

    def test_generation_order(self):
        from icharlotte_core.mediation_brief import GENERATION_ORDER
        self.assertEqual(GENERATION_ORDER[-1], "introduction")
        self.assertEqual(len(GENERATION_ORDER), 7)


class TestTextParsing(unittest.TestCase):

    def test_parse_body_paragraphs(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        text = "First paragraph of the section.\n\nSecond paragraph here."
        elements = gen._parse_section_text(text, "statement_of_facts")
        body_elements = [e for e in elements if e["type"] == "body"]
        self.assertEqual(len(body_elements), 2)
        self.assertIn("First paragraph", body_elements[0]["text"])

    def test_parse_subsection_headings(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        text = (
            "Introduction paragraph.\n\n"
            "SUBSECTION: Defendant Owed No Duty Of Care\n"
            "Content under first subsection.\n\n"
            "SUBSECTION: Plaintiff Was Comparatively At Fault\n"
            "Content under second subsection.\n"
        )
        elements = gen._parse_section_text(text, "liability")
        headings = [e for e in elements if e["type"] == "l2_heading"]
        self.assertEqual(len(headings), 2)
        self.assertEqual(headings[0]["text"], "Defendant Owed No Duty Of Care")
        self.assertEqual(headings[1]["text"], "Plaintiff Was Comparatively At Fault")

    def test_parse_deposition_quotes(self):
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        text = (
            "The plaintiff testified as follows:\n\n"
            "I was not paying attention to the road at the time of the accident. "
            "(Smith Depo Trns., at p. 45:12.)\n\n"
            "This admission is significant."
        )
        elements = gen._parse_section_text(text, "liability")
        quotes = [e for e in elements if e["type"] == "depo_quote"]
        self.assertEqual(len(quotes), 1)
        self.assertIn("not paying attention", quotes[0]["text"])


class TestWordAssembly(unittest.TestCase):

    def test_assemble_creates_docx(self):
        from docx import Document
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        gen.sections = {
            "introduction": "This is the introduction.",
            "statement_of_facts": "These are the facts.",
            "procedural_status": "Trial is set for June 2027.",
            "liability": "SUBSECTION: No Duty\nDefendant owed no duty.",
            "damages": "SUBSECTION: No Causation\nNo evidence of causation.",
            "settlement_position": "Policy limits are $1M.",
            "conclusion": "We will prevail at trial.",
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            caption = Document()
            caption.add_paragraph("LAW FIRM")
            caption.add_paragraph("CAPTION PAGE")
            caption_path = os.path.join(tmpdir, "caption.docx")
            caption.save(caption_path)
            output_path = os.path.join(tmpdir, "brief.docx")
            gen.assemble_document(caption_path, output_path)
            self.assertTrue(os.path.exists(output_path))
            doc = Document(output_path)
            all_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("INTRODUCTION", all_text)
            self.assertIn("STATEMENT OF FACTS", all_text)
            self.assertIn("introduction", all_text.lower())


if __name__ == '__main__':
    unittest.main()
