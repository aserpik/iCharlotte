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
# Use [^\d\n] to handle any separator (middle dot \xb7, non-breaking space \xa0, etc.)
PAGE_NUM_RE = re.compile(r'(?:^|\n)\s*Page[^\d\n]{0,5}(\d+)\s*$', re.MULTILINE)

# Condensed page header within text: " Page 7 " or "Page 7"
CONDENSED_PAGE_RE = re.compile(r'\bPage[^\d\n]{0,5}(\d+)\b')

# Q/A markers: "Q." or "A." (or "Q"/"A" followed by spaces, no period) with leading whitespace
QA_MARKER_RE = re.compile(r'^\s*(Q|A)\.?\s{2,}(.*)$')

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

        # Validate: verify every exchange's text exists in the PDF
        exchanges = self._validate_exchanges(pdf_path, exchanges, pages_text)

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
        self._pdf_path = pdf_path  # Stash for y-coordinate parsing
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

        Uses y-coordinate matching to pair line numbers with their content,
        since PyMuPDF may extract line-number and content spans in different
        text blocks, causing sequential pairing to mismatch lines.

        Returns: List of (page_number, [(line_number, cleaned_text), ...])
        """
        import fitz

        result = []
        doc = fitz.open(self._pdf_path)

        for pdf_page_idx, page_text in enumerate(pages_text):
            raw_lines = page_text.split('\n')

            # Find transcript page number from the raw text
            page_num = self._find_page_number(raw_lines, page_text)
            if page_num is None:
                continue

            # Use y-coordinate matching for accurate line pairing
            page = doc[pdf_page_idx]
            numbered_lines = self._extract_lines_by_position(page)

            if numbered_lines:
                result.append((page_num, numbered_lines))

        doc.close()
        return result

    def _extract_lines_by_position(self, page) -> List[Tuple[int, str]]:
        """
        Extract numbered transcript lines by matching y-coordinates.

        First pass: find line number spans (digits 1-25 at x < 100).
        Second pass: collect all text spans at the same y-coordinate,
        excluding the line number itself, to build the content.
        """
        text_dict = page.get_text("dict")

        # First pass: find line numbers and their y-centers
        line_nums = {}  # line_num -> {"y_center": float, "x_right": float}
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span["text"].strip()
                    bbox = span["bbox"]
                    if bbox[0] < 100 and text.isdigit():
                        num = int(text)
                        if 1 <= num <= 25:
                            y_center = (bbox[1] + bbox[3]) / 2
                            line_nums[num] = {
                                "y_center": y_center,
                                "x_right": bbox[2],
                            }

        if not line_nums:
            return []

        # Second pass: collect content spans for each line number by y-proximity
        line_content = {num: [] for num in line_nums}

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    bbox = span["bbox"]
                    span_y = (bbox[1] + bbox[3]) / 2
                    text = span["text"]

                    # Skip line number spans themselves
                    if bbox[0] < 100 and text.strip().isdigit():
                        continue

                    # Match to closest line number by y-coordinate
                    best_num = None
                    best_dist = 999
                    for num, info in line_nums.items():
                        dist = abs(span_y - info["y_center"])
                        if dist < best_dist:
                            best_dist = dist
                            best_num = num

                    if best_num is not None and best_dist < 8:
                        line_content[best_num].append((bbox[0], text))

        # Assemble lines: sort spans by x-coordinate, join, clean
        numbered_lines = []
        for num in sorted(line_nums.keys()):
            spans = line_content[num]
            if not spans:
                continue

            # Sort by x position and join
            spans.sort(key=lambda s: s[0])
            content = "".join(s[1] for s in spans).strip()

            # Remove timestamp at end
            content = TIMESTAMP_RE.sub('', content)
            content = content.strip()

            # Skip footer/page lines
            if not content or FOOTER_RE.search(content):
                continue
            if re.match(r'^Page[^\d\n]{0,5}\d+$', content, re.IGNORECASE):
                continue

            numbered_lines.append((num, content))

        return numbered_lines

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
        if re.match(r'^\s*Page[^\d\n]{0,5}\d+\s*$', content, re.IGNORECASE):
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
                info.full_name = name.title() if name.isupper() else name
                # Extract last name (last word, ignoring suffixes like M.D.)
                name_parts = name.split()
                if name_parts:
                    raw_last = name_parts[-1]
                    info.last_name = raw_last.title() if raw_last.isupper() else raw_last
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
                # Normalize: replace middle dots (·, \xb7) and non-breaking
                # spaces (\xa0) with regular spaces — common in PDF extraction
                text = text.replace('\xb7', ' ').replace('\xa0', ' ')

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
    # Validation
    # =========================================================================

    def _validate_exchanges(
        self, pdf_path: str, exchanges: List[QAExchange],
        pages_text: List[str]
    ) -> List[QAExchange]:
        """
        Validate that every exchange's text actually exists in the PDF.

        For each exchange, takes a distinctive fragment (8+ words) from the
        question and answer, then checks that it appears in the raw text of
        the PDF pages the exchange references. Exchanges that fail validation
        are dropped with a warning.

        Returns the list of validated exchanges (with original IDs preserved).
        """
        if not exchanges:
            return exchanges

        # Build a map: transcript page number -> raw PDF text
        # We need to figure out which PDF page corresponds to which transcript page.
        # Use the same Page N detection from the page text.
        page_text_map = {}  # transcript_page -> combined raw text
        for pdf_idx, text in enumerate(pages_text):
            for match in PAGE_NUM_RE.finditer(text):
                tp = int(match.group(1))
                if tp not in page_text_map:
                    page_text_map[tp] = ""
                page_text_map[tp] += " " + text

        validated = []
        failed = 0

        for ex in exchanges:
            # Get the raw PDF text for this exchange's page range
            raw_text = ""
            for tp in range(ex.page_start, ex.page_end + 1):
                raw_text += " " + page_text_map.get(tp, "")

            if not raw_text.strip():
                # Can't validate — no PDF text found for this page range.
                # Keep the exchange but warn.
                logger.warning(
                    f"Exchange {ex.id}: no PDF text for pages "
                    f"{ex.page_start}-{ex.page_end}, cannot validate"
                )
                validated.append(ex)
                continue

            # Normalize for comparison: collapse whitespace, lowercase
            raw_norm = re.sub(r'\s+', ' ', raw_text).lower()

            # Check question and answer fragments
            q_ok = self._verify_fragment(ex.question, raw_norm)
            a_ok = self._verify_fragment(ex.answer, raw_norm)

            if q_ok and a_ok:
                validated.append(ex)
            else:
                failed += 1
                which = []
                if not q_ok:
                    which.append("Q")
                if not a_ok:
                    which.append("A")
                logger.warning(
                    f"Exchange {ex.id} FAILED validation ({'+'.join(which)} "
                    f"not found in PDF) at p{ex.page_start}:{ex.line_start}: "
                    f"Q={ex.question[:60]}..."
                )

        if failed:
            logger.error(
                f"VALIDATION: {failed}/{len(exchanges)} exchanges failed "
                f"text verification — these were DROPPED from the index"
            )
        else:
            logger.info(
                f"VALIDATION: all {len(exchanges)} exchanges verified "
                f"against PDF text"
            )

        return validated

    @staticmethod
    def _verify_fragment(text: str, raw_norm: str) -> bool:
        """
        Check that a distinctive fragment of the extracted text exists in
        the raw PDF text.

        Uses progressively shorter fragments: tries 8 words, then 5, then 3.
        Very short texts (1-2 words) are auto-accepted since they're common
        answers like "Yes", "No", "Correct" that appear everywhere in a transcript.
        All comparisons are whitespace-normalized and case-insensitive.
        """
        words = text.split()
        if len(words) <= 2:
            return True  # Short answers like "Yes", "No", "Correct" — skip

        # Try progressively shorter fragments
        for length in [8, 5, 3]:
            if len(words) >= length:
                # Take fragment from the middle (less likely to be truncated)
                mid = max(0, len(words) // 2 - length // 2)
                fragment = ' '.join(words[mid:mid + length]).lower()
                fragment = re.sub(r'\s+', ' ', fragment)
                if fragment in raw_norm:
                    return True

        # Try 2 consecutive words from start and middle
        if len(words) >= 2:
            for start in [0, len(words) // 2]:
                frag = ' '.join(words[start:start + 2]).lower()
                if frag in raw_norm:
                    return True

        # Single distinctive word (4+ chars)
        for w in words:
            if len(w) >= 4 and w.lower() in raw_norm:
                return True

        return False

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
