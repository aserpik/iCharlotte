"""
Deposition Transcript Parser

Reads a PDF deposition transcript and produces a structured Q/A index.
Handles both full-size (1 page per PDF page) and condensed (4 pages per PDF page) formats.
All processing is deterministic — no AI involved.
"""

import os
import re
import logging
from datetime import datetime
from typing import List, Optional, Tuple

from .models import QAExchange, TranscriptIndex, DeponentInfo

logger = logging.getLogger(__name__)


# =============================================================================
# Regex Patterns
# =============================================================================

# Line number at start of line: "  1  " or " 25 " (1-2 digits, surrounded by whitespace)
LINE_NUM_RE = re.compile(r'^\s{0,4}(\d{1,2})\s{1,6}')

# Timestamp at end of line: "10:17:10AM" or "2:05:32PM"
TIMESTAMP_RE = re.compile(r'\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?\s*$', re.IGNORECASE)

# Page header/footer: "Page 7" or "Page  7" at end of extracted text block
PAGE_NUM_RE = re.compile(r'(?:^|\n)\s*Page\s+(\d+)\s*$', re.MULTILINE)

# Condensed page header within text: " Page 7 " or "Page 7"
CONDENSED_PAGE_RE = re.compile(r'\bPage\s+(\d+)\b')

# Q/A markers: "Q." or "A." possibly with leading whitespace (from indentation)
QA_MARKER_RE = re.compile(r'^\s*(Q|A)\.\s+(.*)$')

# Colloquy/non-testimony markers
COLLOQUY_RE = re.compile(
    r'^\s*(?:'
    r'(?:MR|MS|MRS|DR|THE)\.\s+[A-Z]'                # MR. SWAIN:, THE REPORTER:
    r'|THE\s+(?:VIDEOGRAPHER|REPORTER|COURT|WITNESS)'  # THE VIDEOGRAPHER:
    r'|EXAMINATION'                                     # EXAMINATION header
    r'|\(.*\)\s*$'                                      # (Recess taken.)
    r'|---'                                             # --- dividers
    r')',
    re.IGNORECASE
)

# Deponent info patterns (reused from summarize_deposition.py)
DEPONENT_NAME_PATTERNS = [
    re.compile(r'DEPOSITION\s+(?:OF|of)\s+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+)', re.IGNORECASE),
    re.compile(r'[Dd]eponent[:\s]+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+)'),
    re.compile(r'WITNESS[:\s]+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+)'),
]

DATE_PATTERNS = [
    re.compile(r'(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+([A-Z][a-z]+\s+\d{1,2},?\s+\d{4})'),
    re.compile(r'(?:taken|held|conducted)\s+(?:on\s+)?([A-Z][a-z]+\s+\d{1,2},?\s+\d{4})', re.IGNORECASE),
    re.compile(r'([A-Z][a-z]+\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})'),
    re.compile(r'(\d{1,2}/\d{1,2}/\d{4})'),
]

CASE_NUMBER_RE = re.compile(r'(?:No\.|Case\s+No\.?|Case\s+#)\s*([A-Z0-9]+\d+)', re.IGNORECASE)

# Veritext / court reporter footer
FOOTER_RE = re.compile(
    r'(?:Veritext|Calendar-CA|www\.|Job No\.|CSR No\.)',
    re.IGNORECASE
)


class TranscriptParser:
    """
    Parses a PDF deposition transcript into a structured Q/A index.

    Usage:
        parser = TranscriptParser()
        index = parser.parse("path/to/transcript.pdf")
        index.save()  # caches as JSON alongside the PDF
    """

    def parse(self, pdf_path: str, force_reparse: bool = False) -> TranscriptIndex:
        """
        Parse a deposition transcript PDF into a structured index.

        Args:
            pdf_path: Path to the PDF transcript.
            force_reparse: If True, ignore cached index.

        Returns:
            TranscriptIndex with all Q/A exchanges.
        """
        pdf_path = os.path.abspath(pdf_path)

        # Check cache
        if not force_reparse and TranscriptIndex.is_cache_valid(pdf_path):
            cache_path = TranscriptIndex.cache_path_for(pdf_path)
            logger.info(f"Loading cached index: {cache_path}")
            index = TranscriptIndex.load(cache_path)
            # Regenerate raw_text for viewer (not stored in cache)
            index.raw_text = self._rebuild_raw_text(index)
            return index

        logger.info(f"Parsing transcript: {pdf_path}")

        # Extract text from PDF using PyMuPDF (better layout preservation)
        pages_text = self._extract_pdf_text(pdf_path)

        # Detect format: full-size vs condensed
        is_condensed = self._detect_condensed(pages_text)
        logger.info(f"Format detected: {'condensed' if is_condensed else 'full-size'}")

        # Split into transcript pages
        if is_condensed:
            transcript_pages = self._split_condensed_pages(pages_text)
        else:
            transcript_pages = self._process_full_size_pages(pages_text)

        # Extract deponent info from first few pages
        header_text = "\n".join(pages_text[:5]) if pages_text else ""
        deponent = self._extract_deponent_info(header_text)

        # Parse Q/A exchanges from transcript pages
        exchanges = self._parse_qa_exchanges(transcript_pages)

        # Build the raw text for viewer display
        raw_lines = []
        for page_num, lines in transcript_pages:
            raw_lines.append(f"\n--- Page {page_num} ---\n")
            for line_num, text in lines:
                raw_lines.append(f"{line_num:>3}  {text}")
        raw_text = "\n".join(raw_lines)

        # Build index
        max_page = max((p for p, _ in transcript_pages), default=0)
        index = TranscriptIndex(
            source_pdf=pdf_path,
            deponent=deponent,
            exchanges=exchanges,
            total_transcript_pages=max_page,
            is_condensed=is_condensed,
            parse_timestamp=datetime.now().isoformat(),
            raw_text=raw_text,
        )

        # Cache the index
        cache_path = index.save()
        logger.info(f"Parsed {len(exchanges)} Q/A exchanges, cached to {cache_path}")

        return index

    # =========================================================================
    # PDF Text Extraction
    # =========================================================================

    def _extract_pdf_text(self, pdf_path: str) -> List[str]:
        """Extract text from each PDF page using PyMuPDF for better layout."""
        import fitz
        pages = []
        doc = fitz.open(pdf_path)
        for page in doc:
            text = page.get_text("text")
            pages.append(text)
        doc.close()
        return pages

    # =========================================================================
    # Format Detection
    # =========================================================================

    def _detect_condensed(self, pages_text: List[str]) -> bool:
        """
        Detect if transcript is condensed (4 transcript pages per PDF page).

        Condensed transcripts typically have multiple "Page N" markers per PDF page
        and a footer like "2 (Pages 2 - 5)".
        """
        if len(pages_text) < 2:
            return False

        # Check a middle page for multiple Page markers
        check_pages = pages_text[1:min(4, len(pages_text))]
        for page_text in check_pages:
            page_markers = PAGE_NUM_RE.findall(page_text)
            if len(page_markers) >= 3:
                return True

            # Also check for condensed footer pattern: "2 (Pages 6 - 9)"
            if re.search(r'\d+\s+\(Pages?\s+\d+\s*-\s*\d+\)', page_text):
                return True

        return False

    # =========================================================================
    # Full-Size Processing
    # =========================================================================

    def _process_full_size_pages(self, pages_text: List[str]) -> List[Tuple[int, List[Tuple[int, str]]]]:
        """
        Process full-size transcript: 1 transcript page per PDF page.

        PyMuPDF extracts line numbers on their own line, followed by
        the content on the next line. E.g.:
            "1\\n       Q.   Good morning...    10:17:10AM\\n2\\n  begin..."

        We pair each standalone number with its content line.

        Returns: List of (page_number, [(line_number, cleaned_text), ...])
        """
        result = []

        for pdf_page_idx, page_text in enumerate(pages_text):
            raw_lines = page_text.split('\n')

            # Find transcript page number
            page_num = self._find_page_number(raw_lines, page_text)
            if page_num is None:
                continue

            # Pair line numbers with content lines
            # Pattern: a line that is just a number (1-25) followed by a content line
            numbered_lines = []
            i = 0
            while i < len(raw_lines):
                stripped = raw_lines[i].strip()

                # Check if this line is a standalone line number (1-25)
                if stripped.isdigit() and 1 <= int(stripped) <= 25:
                    line_num = int(stripped)
                    # Next line is the content
                    if i + 1 < len(raw_lines):
                        content = raw_lines[i + 1]
                        # Remove timestamp at end
                        content = TIMESTAMP_RE.sub('', content)
                        content = content.strip()

                        # Skip footer/page lines
                        if content and not FOOTER_RE.search(content):
                            if not re.match(r'^Page\s+\d+$', content, re.IGNORECASE):
                                numbered_lines.append((line_num, content))
                        i += 2
                        continue

                i += 1

            if numbered_lines:
                result.append((page_num, numbered_lines))

        return result

    def _find_page_number(self, lines: List[str], full_text: str) -> Optional[int]:
        """Extract the transcript page number from page text."""
        # Look for "Page N" pattern (usually at bottom)
        match = PAGE_NUM_RE.search(full_text)
        if match:
            return int(match.group(1))

        # Look in last few lines for just a number
        for line in reversed(lines[-5:]):
            stripped = line.strip()
            if stripped.isdigit() and 1 <= int(stripped) <= 999:
                return int(stripped)

        return None

    def _parse_line(self, line: str) -> Optional[Tuple[int, str]]:
        """
        Parse a single transcript line, extracting line number and clean content.

        Returns (line_number, cleaned_text) or None if not a numbered content line.
        """
        # Skip empty lines
        if not line.strip():
            return None

        # Skip footer lines
        if FOOTER_RE.search(line):
            return None

        # Match line number at start
        match = LINE_NUM_RE.match(line)
        if not match:
            return None

        line_num = int(match.group(1))
        if line_num < 1 or line_num > 25:
            return None

        # Extract content after line number
        content = line[match.end():]

        # Remove timestamp at end
        content = TIMESTAMP_RE.sub('', content)

        # Clean trailing whitespace
        content = content.rstrip()

        # Skip if content is just "Page N" (page number at bottom)
        if re.match(r'^\s*Page\s+\d+\s*$', content, re.IGNORECASE):
            return None

        return (line_num, content)

    # =========================================================================
    # Condensed Processing
    # =========================================================================

    def _split_condensed_pages(self, pages_text: List[str]) -> List[Tuple[int, List[Tuple[int, str]]]]:
        """
        Process condensed transcript: 4 transcript pages per PDF page.

        Strategy: Split the extracted text on "Page N" boundaries to recover
        individual transcript pages, then parse each one.
        """
        result = []

        for pdf_page_text in pages_text:
            # Find all "Page N" markers and their positions
            page_markers = list(CONDENSED_PAGE_RE.finditer(pdf_page_text))
            if not page_markers:
                continue

            # Split text into segments between Page markers
            # Each segment is one transcript page's content
            for i, marker in enumerate(page_markers):
                page_num = int(marker.group(1))

                # Get text from after this page marker to the next one (or end)
                start = marker.end()
                end = page_markers[i + 1].start() if i + 1 < len(page_markers) else len(pdf_page_text)
                segment = pdf_page_text[start:end]

                # Also include text before the first page marker if this is the first
                if i == 0:
                    pre_text = pdf_page_text[:marker.start()]
                    # Check if pre_text contains numbered lines that belong to a previous page
                    # (this handles the case where PyMuPDF interleaves columns)

                # Parse numbered lines from segment
                numbered_lines = []
                for line in segment.split('\n'):
                    parsed = self._parse_line(line)
                    if parsed:
                        numbered_lines.append(parsed)

                if numbered_lines:
                    result.append((page_num, numbered_lines))

        # Sort by page number (condensed extraction may not be in order)
        result.sort(key=lambda x: x[0])

        # Deduplicate pages (same page might appear in multiple PDF pages)
        seen = set()
        deduped = []
        for page_num, lines in result:
            if page_num not in seen:
                seen.add(page_num)
                deduped.append((page_num, lines))

        return deduped

    # =========================================================================
    # Deponent Info Extraction
    # =========================================================================

    def _extract_deponent_info(self, header_text: str) -> DeponentInfo:
        """Extract deponent information from the first few pages of the transcript."""
        info = DeponentInfo()

        # Extract name
        for pattern in DEPONENT_NAME_PATTERNS:
            match = pattern.search(header_text)
            if match:
                name = match.group(1).strip()
                name = re.sub(r'\s+', ' ', name).rstrip('.,;:')
                info.full_name = name
                # Extract last name (last word, ignoring suffixes like M.D.)
                name_parts = name.split()
                if name_parts:
                    info.last_name = name_parts[-1]
                break

        # Extract date
        for pattern in DATE_PATTERNS:
            match = pattern.search(header_text)
            if match:
                info.deposition_date = match.group(1).strip()
                break

        # Extract case number
        match = CASE_NUMBER_RE.search(header_text)
        if match:
            info.case_number = match.group(1)

        # Extract case name from caption (e.g., "SALTARELLI vs. CITY OF HESPERIA")
        vs_match = re.search(
            r'([A-Z][A-Z\s,;]+?)\s+vs?\.?\s+([A-Z][A-Z\s,;]+?)(?:\s*\)|\s*$)',
            header_text, re.MULTILINE
        )
        if vs_match:
            plaintiff = vs_match.group(1).strip().rstrip(',;')
            defendant = vs_match.group(2).strip().rstrip(',;')
            info.case_name = f"{plaintiff} v. {defendant}"

        # Extract volume
        vol_match = re.search(r'Volume\s+(\d+)', header_text, re.IGNORECASE)
        if vol_match:
            info.volume = vol_match.group(1)

        logger.info(f"Deponent: {info.full_name} ({info.last_name}), Date: {info.deposition_date}")
        return info

    # =========================================================================
    # Q/A Exchange Parsing
    # =========================================================================

    def _parse_qa_exchanges(
        self, transcript_pages: List[Tuple[int, List[Tuple[int, str]]]]
    ) -> List[QAExchange]:
        """
        Parse all Q/A exchanges from the transcript pages.

        Walks through every line, tracking Q and A markers. When a complete
        Q+A pair is found, it's recorded as an exchange.
        """
        exchanges = []
        exchange_id = 0

        # State machine
        current_q_text = ""
        current_q_page = 0
        current_q_line = 0
        current_a_text = ""
        current_a_page = 0
        current_a_line = 0
        state = "idle"  # idle | in_question | in_answer

        for page_num, lines in transcript_pages:
            for line_num, text in lines:
                # Skip empty content
                if not text.strip():
                    continue

                # Skip colloquy lines (MR. SWAIN:, THE VIDEOGRAPHER:, etc.)
                if COLLOQUY_RE.match(text):
                    # If we're in the middle of collecting, colloquy interrupts
                    # but doesn't end the Q/A pair — the next Q or A continues it
                    continue

                # Check for Q/A markers
                qa_match = QA_MARKER_RE.match(text)

                if qa_match:
                    marker = qa_match.group(1)
                    content = qa_match.group(2).strip()

                    if marker == 'Q':
                        # If we were in an answer, save the complete exchange
                        if state == "in_answer" and current_q_text and current_a_text:
                            exchange_id += 1
                            exchanges.append(QAExchange(
                                id=exchange_id,
                                question=current_q_text.strip(),
                                answer=current_a_text.strip(),
                                page_start=current_q_page,
                                line_start=current_q_line,
                                page_end=current_a_page,
                                line_end=current_a_line,
                            ))

                        # Start new question
                        current_q_text = content
                        current_q_page = page_num
                        current_q_line = line_num
                        state = "in_question"

                    elif marker == 'A':
                        if state == "in_question":
                            # Start answer
                            current_a_text = content
                            current_a_page = page_num
                            current_a_line = line_num
                            state = "in_answer"
                        elif state == "in_answer":
                            # Another A without a Q — rare but append to current answer
                            current_a_text += " " + content
                            current_a_page = page_num
                            current_a_line = line_num

                else:
                    # Continuation line (no Q/A marker)
                    if state == "in_question":
                        current_q_text += " " + text.strip()
                    elif state == "in_answer":
                        current_a_text += " " + text.strip()
                        current_a_page = page_num
                        current_a_line = line_num

        # Don't forget the last exchange
        if state == "in_answer" and current_q_text and current_a_text:
            exchange_id += 1
            exchanges.append(QAExchange(
                id=exchange_id,
                question=current_q_text.strip(),
                answer=current_a_text.strip(),
                page_start=current_q_page,
                line_start=current_q_line,
                page_end=current_a_page,
                line_end=current_a_line,
            ))

        logger.info(f"Parsed {len(exchanges)} Q/A exchanges from {len(transcript_pages)} pages")
        return exchanges

    # =========================================================================
    # Helpers
    # =========================================================================

    def _rebuild_raw_text(self, index: TranscriptIndex) -> str:
        """Rebuild display text from exchanges (when loading from cache)."""
        lines = []
        for ex in index.exchanges:
            lines.append(f"--- p.{ex.page_start}:{ex.line_start} ---")
            lines.append(f"Q. {ex.question}")
            lines.append(f"A. {ex.answer}")
            lines.append("")
        return "\n".join(lines)
