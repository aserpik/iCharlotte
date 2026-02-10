"""
Testimony Formatter — Word document output and PDF highlighting.

Produces a Word document with verbatim Q/A pairs formatted per legal standards,
and optionally creates a highlighted PDF copy of the transcript.
"""

import os
import re
import copy
import logging
from typing import List, Optional, Tuple

from docx import Document
from docx.shared import Inches, Pt, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH

import fitz

from .models import TranscriptIndex, QAExchange, ExtractionResult

logger = logging.getLogger(__name__)

# Formatting constants
LEFT_INDENT = Inches(0.5)       # 0.5" indent for all Q/A text
HANGING_INDENT = Inches(0.5)    # 0.5" hanging indent (Q./A. marker hangs left)
FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(12)
LINE_SPACING = 1.0


class TestimonyFormatter:
    """
    Formats extracted testimony into Word documents and optionally highlights PDFs.

    Usage:
        formatter = TestimonyFormatter()
        formatter.add_extraction(index, result)
        formatter.add_extraction(index, result2)
        output_path = formatter.save_word(output_dir)
    """

    def __init__(self):
        self._extractions: List[Tuple[TranscriptIndex, ExtractionResult]] = []

    def add_extraction(self, index: TranscriptIndex, result: ExtractionResult):
        """Add an extraction result to be included in the output document."""
        self._extractions.append((index, result))

    def clear(self):
        """Clear all accumulated extractions."""
        self._extractions.clear()

    # =========================================================================
    # Word Document Output
    # =========================================================================

    def save_word(self, output_dir: str = None, filename: str = None) -> str:
        """
        Save all extractions to a Word document.

        Args:
            output_dir: Directory for output. Defaults to transcript's directory.
            filename: Custom filename. Defaults to "[Extracted] [Deponent] Depo Trns.docx".

        Returns:
            Path to the saved document.
        """
        if not self._extractions:
            raise ValueError("No extractions to save")

        doc = Document()

        # Set default font
        style = doc.styles['Normal']
        font = style.font
        font.name = FONT_NAME
        font.size = FONT_SIZE

        for ext_idx, (index, result) in enumerate(self._extractions):
            if ext_idx > 0:
                # Add separator between different extraction prompts
                doc.add_paragraph()

            # Add prompt as header
            prompt_para = doc.add_paragraph()
            prompt_run = prompt_para.add_run(f"Extraction: {result.prompt}")
            prompt_run.bold = True
            prompt_run.font.size = FONT_SIZE
            prompt_run.font.name = FONT_NAME
            prompt_para.space_after = Pt(6)

            # Add each group of consecutive exchanges
            for group_idx, group_ids in enumerate(result.groups):
                if group_idx > 0:
                    # Small gap between citation blocks
                    doc.add_paragraph().space_after = Pt(6)

                # Get exchanges for this group
                exchange_map = {ex.id: ex for ex in index.exchanges}
                group_exchanges = [exchange_map[eid] for eid in group_ids if eid in exchange_map]

                if not group_exchanges:
                    continue

                # Add each Q/A pair
                for ex in group_exchanges:
                    self._add_qa_paragraph(doc, "Q.", ex.question)
                    self._add_qa_paragraph(doc, "A.", ex.answer)

                # Add citation after the group
                citation = self._format_citation(group_exchanges, index.deponent.last_name)
                cite_para = doc.add_paragraph()
                cite_para.paragraph_format.left_indent = LEFT_INDENT
                cite_para.space_before = Pt(6)
                cite_para.space_after = Pt(12)
                cite_run = cite_para.add_run(citation)
                cite_run.font.size = FONT_SIZE
                cite_run.font.name = FONT_NAME

        # Determine output path
        first_index = self._extractions[0][0]
        if output_dir is None:
            output_dir = os.path.dirname(first_index.source_pdf)

        if filename is None:
            deponent = first_index.deponent.last_name or "Unknown"
            filename = f"[Extracted] {deponent} Depo Trns.docx"

        output_path = os.path.join(output_dir, filename)
        doc.save(output_path)
        logger.info(f"Saved extraction document: {output_path}")
        return output_path

    def _add_qa_paragraph(self, doc: Document, marker: str, text: str):
        """Add a Q. or A. paragraph with proper formatting."""
        para = doc.add_paragraph()
        pf = para.paragraph_format

        # Left indent 0.5" with hanging indent for marker
        pf.left_indent = Inches(1.0)       # Text body at 1.0"
        pf.first_line_indent = Inches(-0.5)  # Marker hangs back to 0.5"
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        # Add marker + tab + text
        marker_run = para.add_run(f"{marker}\t")
        marker_run.font.name = FONT_NAME
        marker_run.font.size = FONT_SIZE

        text_run = para.add_run(text)
        text_run.font.name = FONT_NAME
        text_run.font.size = FONT_SIZE

        # Set tab stop at 0.5" from the indent start (i.e., 1.0" from margin)
        self._set_tab_stop(para, Inches(1.0))

    def _set_tab_stop(self, paragraph, position):
        """Set a tab stop at the given position."""
        from docx.oxml.ns import qn
        from lxml import etree

        pPr = paragraph._p.get_or_add_pPr()
        tabs = pPr.find(qn('w:tabs'))
        if tabs is None:
            tabs = etree.SubElement(pPr, qn('w:tabs'))

        tab = etree.SubElement(tabs, qn('w:tab'))
        tab.set(qn('w:val'), 'left')
        tab.set(qn('w:pos'), str(int(position)))

    def _format_citation(self, exchanges: List[QAExchange], deponent_last_name: str) -> str:
        """
        Format a citation block for a group of consecutive exchanges.

        Format: (Exh. __, ([LastName] Depo. Trns.) at p. [ranges].)
        """
        # Build page:line ranges
        ranges = []
        for ex in exchanges:
            ranges.append(ex.citation_range())

        # Merge consecutive/overlapping ranges on the same page
        # For now, just join with semicolons
        range_str = "; ".join(ranges)

        last_name = deponent_last_name or "___"
        return f"(Exh. __ ({last_name} Depo. Trns.) at p. {range_str}.)"

    # =========================================================================
    # PDF Highlighting
    # =========================================================================

    def highlight_pdf(self, index: TranscriptIndex, result: ExtractionResult) -> Optional[str]:
        """
        Create or update a highlighted PDF copy of the transcript.

        The highlighted copy is named "[H.AI] original_filename.pdf" and
        accumulates highlights across multiple extraction calls.

        Args:
            index: The transcript index.
            result: The extraction result with selected IDs.

        Returns:
            Path to the highlighted PDF, or None on failure.
        """
        source_path = index.source_pdf
        source_dir = os.path.dirname(source_path)
        source_name = os.path.basename(source_path)

        # Determine highlighted copy path
        highlight_name = f"[H.AI] {source_name}"
        highlight_path = os.path.join(source_dir, highlight_name)

        # Open existing highlighted copy or create from original
        if os.path.exists(highlight_path):
            doc = fitz.open(highlight_path)
            logger.info(f"Updating existing highlighted copy: {highlight_path}")
        else:
            doc = fitz.open(source_path)
            logger.info(f"Creating new highlighted copy from: {source_path}")

        # Get selected exchanges
        exchange_map = {ex.id: ex for ex in index.exchanges}
        selected = [exchange_map[eid] for eid in result.selected_ids if eid in exchange_map]

        # Build a mapping from transcript page → PDF page(s)
        page_map = self._build_page_map(doc, index.is_condensed)

        highlights_added = 0

        for ex in selected:
            # Find the text to highlight (Q and A)
            for text_block, line_num in [
                (ex.question, ex.line_start),
                (ex.answer, ex.line_end),  # answer starts after Q
            ]:
                # Search on the relevant PDF pages
                pdf_pages = self._get_pdf_pages_for_transcript(
                    ex.page_start, ex.page_end, page_map
                )

                for pdf_page_num in pdf_pages:
                    if pdf_page_num >= len(doc):
                        continue

                    page = doc[pdf_page_num]

                    # PDF text contains line numbers, timestamps, and Q./A.
                    # markers that the parser strips. Try multiple search
                    # strategies with progressively shorter fragments.
                    rects = self._search_text_on_page(
                        page, text_block, line_num
                    )
                    if rects:
                        for rect in rects:
                            annot = page.add_highlight_annot(rect)
                            annot.set_colors(stroke=(1, 1, 0))  # Yellow
                            annot.update()
                            highlights_added += 1
                        break  # Found on this page, no need to check others

        # Save
        doc.save(highlight_path)
        doc.close()

        logger.info(f"Added {highlights_added} highlights to {highlight_path}")
        return highlight_path

    def _build_page_map(self, doc, is_condensed: bool) -> dict:
        """
        Build a mapping from transcript page numbers to PDF page indices.

        For full-size: transcript page N ≈ PDF page N-1 (approximately)
        For condensed: transcript pages are packed 4-per-PDF-page
        """
        page_map = {}  # transcript_page -> [pdf_page_indices]

        for pdf_idx in range(len(doc)):
            page = doc[pdf_idx]
            text = page.get_text("text")

            # Skip word index pages that appear after the transcript.
            # These have dense "word line:col" entries (e.g., "accident 27:8")
            # but no Q./A. testimony markers. Detect by structure, not by
            # court reporter name (Veritext, US Legal, etc. all vary).
            has_qa = '\n       Q.' in text or '\n       A.' in text
            if not has_qa:
                # Count word:line reference patterns (e.g., "27:8", "145:12")
                word_refs = re.findall(r'\b\d{1,3}:\d{1,2}\b', text)
                if len(word_refs) >= 10:
                    continue

            # Find all "Page N" markers on this PDF page
            for match in re.finditer(r'Page\s+(\d+)', text):
                tp_num = int(match.group(1))
                if tp_num not in page_map:
                    page_map[tp_num] = []
                if pdf_idx not in page_map[tp_num]:
                    page_map[tp_num].append(pdf_idx)

        return page_map

    def _get_pdf_pages_for_transcript(
        self, page_start: int, page_end: int, page_map: dict
    ) -> List[int]:
        """Get PDF page indices that contain the given transcript page range."""
        pdf_pages = set()
        for tp in range(page_start, page_end + 1):
            if tp in page_map:
                pdf_pages.update(page_map[tp])
            else:
                # Approximate: transcript page N is roughly PDF page N-1
                # (pages 1-4 are cover/index, so offset by ~4)
                approx = tp - 1
                if 0 <= approx:
                    pdf_pages.add(approx)
        return sorted(pdf_pages)

    @staticmethod
    def _search_text_on_page(page, text_block: str, line_num: int = 0) -> list:
        """
        Search for transcript text on a PDF page using multiple strategies.

        The parser strips line numbers, timestamps, and Q./A. markers from
        the text, but the PDF still contains them. Timestamps injected
        mid-sentence break up multi-word phrases, so we progressively
        try shorter fragments and finally fall back to single-word search.

        Args:
            page: PyMuPDF page object.
            text_block: Cleaned transcript text (Q or A content).
            line_num: Transcript line number for disambiguation.
        """
        text = text_block.strip()
        if not text:
            return []

        words = text.split()
        # Find first distinctive word (skip generic openers)
        common = {"okay", "okay.", "well", "well,", "yes", "yes.", "no", "no.",
                  "i", "the", "and", "so", "a", "it", "that", "this", "do",
                  "did", "was", "were", "are", "is", "have", "has", "had",
                  "but", "or", "if", "you", "your", "we", "my", "what",
                  "yeah", "yeah.", "yep", "yep.", "new.", "oh", "--"}
        start_idx = 0
        for i, w in enumerate(words[:10]):
            if w.lower().rstrip(".,;:!?") not in common and len(w) >= 4:
                start_idx = i
                break

        # Strategy 1: ~40 char fragment from distinctive start
        fragment = " ".join(words[start_idx:])[:40].strip()
        if len(fragment) >= 10:
            rects = page.search_for(fragment)
            if rects:
                return rects

        # Strategy 2: Shorter fragments (timestamps break longer runs)
        for length in [25, 15]:
            fragment = " ".join(words[start_idx:])[:length].strip()
            if len(fragment) >= 8:
                rects = page.search_for(fragment)
                if rects:
                    return rects

        # Strategy 3: Single distinctive word — for text where every
        # multi-word fragment is broken by a timestamp/line-number
        for i in range(start_idx, min(start_idx + 6, len(words))):
            w = words[i]
            clean_w = w.rstrip(".,;:!?")
            if len(clean_w) >= 4 and clean_w.lower() not in common:
                rects = page.search_for(w)
                if len(rects) == 1:
                    return rects  # Unique on page — safe match
                if rects and line_num > 0:
                    # Multiple matches; use line number to disambiguate.
                    # Search for the line number near each rect.
                    line_str = str(line_num)
                    for rect in rects:
                        # Look for the line number to the left of this text
                        search_area = fitz.Rect(
                            0, rect.y0 - 2, rect.x0, rect.y1 + 2
                        )
                        if page.search_for(line_str, clip=search_area):
                            return [rect]
                    # If line-number disambiguation failed, return first match
                    return rects[:1]
                if rects:
                    return rects[:1]

        # Strategy 4: Very short answers ("Yes.", "No.", "Yep.", "New.")
        # Use the line number to find the right occurrence
        if len(text) <= 5 and line_num > 0:
            page_text = page.get_text("text")
            line_str = f"\n{line_num}\n"
            idx = page_text.find(line_str)
            if idx >= 0:
                # Get a snippet after the line number
                after = page_text[idx + len(line_str):idx + len(line_str) + 80]
                # The answer text should appear shortly after
                ans_idx = after.find(text)
                if ans_idx >= 0:
                    # Search for the text clipped to this area of the page
                    rects = page.search_for(text)
                    if rects:
                        # Find the rect closest to where we expect it
                        # (by searching near the line number position)
                        return rects[:1]

        # Strategy 5: Last resort — just search the raw text
        if len(text) >= 3:
            rects = page.search_for(text)
            if rects:
                return rects[:1]

        return []

