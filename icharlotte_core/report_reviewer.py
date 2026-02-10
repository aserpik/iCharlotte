"""
Report Review Engine — reviews junior associate draft reports against
gold standard voice, style, formatting, and optionally case data for completeness.

Two-pass approach:
  Pass 1: Structural scan (rule-based, fast) — checks formatting, completeness, headings
  Pass 2: Voice/content review (LLM, section-by-section) — applies Track Changes

Usage from the AI Assistant (word_hotkey.py):
    reviewer = ReportReviewer(doc_com, case_path="/path/to/case")
    structural, total = reviewer.run(
        include_case_data=True,
        extra_doc_paths=["/path/to/complaint.pdf"],
        progress_callback=lambda msg: status_label.setText(msg),
    )
"""

import logging
import os
import re
import time
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────────────────────────────
# Data classes
# ──────────────────────────────────────────────────────────────────────

@dataclass
class SectionInfo:
    """A detected section in the draft report."""
    name: str               # Canonical section name (e.g., "SETTLEMENT STATUS")
    raw_name: str           # As found in document (e.g., "SETTLEMENT")
    para_index: int         # 1-based COM paragraph index of the heading
    range_start: int        # Character position of the section content start
    range_end: int          # Character position of the section content end
    text: str = ""          # Extracted section content text
    heading_range_start: int = 0  # Character position of the heading itself
    heading_range_end: int = 0


# ──────────────────────────────────────────────────────────────────────
# Expected section order and typical lengths
# ──────────────────────────────────────────────────────────────────────

EXPECTED_SECTIONS = [
    "FACTUAL BACKGROUND",
    "PROCEDURAL STATUS",
    "INVESTIGATION",
    "DISCOVERY",
    "EXPERTS",
    "MEDICAL RECORD REVIEW",
    "EVALUATION OF LIABILITY",
    "EVALUATION OF EXPOSURE",
    "SETTLEMENT STATUS",
    "FURTHER CASE HANDLING",
]

# Minimum chars before a section is flagged as "thin"
MIN_SECTION_LENGTH = 200


# ──────────────────────────────────────────────────────────────────────
# ReportReviewer class
# ──────────────────────────────────────────────────────────────────────

class ReportReviewer:
    """Orchestrates the two-pass review of a junior associate's draft report."""

    def __init__(self, doc_com, case_path: Optional[str] = None):
        """
        Args:
            doc_com: win32com Word Document object (the open draft report)
            case_path: Optional case folder path for loading AI OUTPUT data
        """
        self.doc = doc_com
        self.case_path = case_path
        self._style_guide = None
        self._section_examples = {}
        self._llm_caller = None

    # ────── Public API ──────

    def run(self, selection_range: Optional[Tuple[int, int]] = None,
            include_case_data: bool = False,
            extra_doc_paths: Optional[List[str]] = None,
            progress_callback: Optional[Callable[[str], None]] = None,
            ) -> Tuple[Optional['ValidationResult'], int]:
        """Main entry point for the report review.

        Args:
            selection_range: (start, end) character positions for selection-only review.
                             If None, reviews the full document.
            include_case_data: Whether to load case data from AI OUTPUT folder
            extra_doc_paths: Additional document paths to extract text from
            progress_callback: Called with progress messages for UI display

        Returns:
            (structural_findings, total_changes_applied)
            structural_findings is None for selection-only reviews
        """
        def progress(msg):
            logger.info(msg)
            if progress_callback:
                progress_callback(msg)

        # Load style guide and examples
        progress("Loading style guide and examples...")
        self._load_style_resources()

        # Load optional context
        case_data = None
        if include_case_data and self.case_path:
            progress("Loading case data from AI OUTPUT...")
            case_data = self._load_case_data()

        extra_texts = []
        if extra_doc_paths:
            progress(f"Extracting text from {len(extra_doc_paths)} additional documents...")
            extra_texts = self._extract_extra_docs(extra_doc_paths)

        # Get full document text for context
        full_doc_text = self._get_full_doc_text()

        total_changes = 0

        if selection_range:
            # Selection-only mode: Skip Pass 1, go straight to LLM review
            progress("Reviewing selected text...")
            sel_text = self._get_range_text(selection_range[0], selection_range[1])
            if sel_text.strip():
                revised = self._review_text(
                    sel_text, "SELECTED TEXT", full_doc_text,
                    case_data, extra_texts
                )
                if revised and revised.strip() != sel_text.strip():
                    progress("Applying redlines to selection...")
                    changes = self._apply_redline(
                        selection_range[0], selection_range[1],
                        sel_text, revised
                    )
                    total_changes += changes
            progress(f"Review complete — {total_changes} tracked changes applied.")
            return None, total_changes

        # Full document mode
        # Pass 1: Structural scan
        progress("Pass 1: Scanning document structure...")
        sections = self.detect_sections()
        progress(f"Pass 1: Found {len(sections)} sections, checking formatting...")
        structural = self.run_structural_scan(sections)
        structural.print_summary()

        error_count = structural.error_count
        warn_count = structural.warn_count
        if error_count or warn_count:
            progress(f"STRUCTURAL SCAN — {error_count} errors, {warn_count} warnings")
        else:
            progress("STRUCTURAL SCAN — no issues found")

        # Pass 2: Section-by-section LLM review with redlines
        reviewable = [s for s in sections if s.text.strip()]
        for i, section in enumerate(reviewable, 1):
            progress(f"Pass 2: Reviewing {section.name} ({i}/{len(reviewable)})...")

            revised = self._review_text(
                section.text, section.name, full_doc_text,
                case_data, extra_texts
            )

            if not revised or revised.strip() == section.text.strip():
                progress(f"  {section.name}: no changes needed")
                continue

            progress(f"Applying redlines to {section.name}...")
            changes = self._apply_redline(
                section.range_start, section.range_end,
                section.text, revised
            )
            total_changes += changes

            # Fix heading name if it doesn't match canonical
            if section.raw_name != section.name:
                self._fix_heading_name(section)

        progress(f"Review complete — {len(reviewable)} sections reviewed, "
                 f"{total_changes} tracked changes applied.")
        return structural, total_changes

    # ────── Pass 1: Structural Scan ──────

    def detect_sections(self) -> List[SectionInfo]:
        """Scan the document for section headings via COM.

        Looks for ALL CAPS paragraphs that are bold, matching known headings.
        Returns SectionInfo list with content ranges.
        """
        from icharlotte_core.word_validator import fuzzy_match_section_heading

        sections = []
        paras = self.doc.Paragraphs
        para_count = paras.Count
        logger.info(f"Scanning {para_count} paragraphs for section headings...")

        heading_indices = []  # (1-based index, canonical_name, raw_name, para_obj)

        for i in range(1, para_count + 1):
            try:
                para = paras(i)
                text = para.Range.Text.strip().rstrip('\r')
                if not text:
                    continue

                # Quick checks: must be short-ish and ALL CAPS
                if len(text) > 80 or not text.isupper():
                    continue

                # Check if bold
                try:
                    is_bold = para.Range.Bold
                    if is_bold != -1 and is_bold != True:  # noqa: E712
                        continue
                except:
                    continue

                # Fuzzy match to canonical section name
                canonical = fuzzy_match_section_heading(text)
                if canonical:
                    heading_indices.append((i, canonical, text, para))
                    logger.info(f"  Found: '{text}' -> '{canonical}' (para {i})")

            except Exception as e:
                logger.debug(f"Error scanning para {i}: {e}")
                continue

        # Build SectionInfo with content ranges
        for idx, (para_idx, canonical, raw, para) in enumerate(heading_indices):
            heading_start = para.Range.Start
            heading_end = para.Range.End

            # Content starts after this heading paragraph
            content_start = heading_end

            # Content ends at the next heading's start, or doc end
            if idx + 1 < len(heading_indices):
                next_para = heading_indices[idx + 1][3]
                content_end = next_para.Range.Start
            else:
                # Last section — content goes to end of document body
                # (before closing/signature block, approximated as doc end)
                content_end = self.doc.Content.End

            # Extract text
            try:
                content_range = self.doc.Range(content_start, content_end)
                section_text = content_range.Text or ""
            except:
                section_text = ""

            sections.append(SectionInfo(
                name=canonical,
                raw_name=raw,
                para_index=para_idx,
                range_start=content_start,
                range_end=content_end,
                text=section_text,
                heading_range_start=heading_start,
                heading_range_end=heading_end,
            ))

        logger.info(f"Detected {len(sections)} sections")
        return sections

    def run_structural_scan(self, sections: List[SectionInfo]) -> 'ValidationResult':
        """Pass 1: Rule-based formatting and completeness checks.

        Returns a ValidationResult with findings.
        """
        from icharlotte_core.word_validator import Finding, ValidationResult

        findings = []

        # 1. Check which expected sections are present
        found_names = {s.name for s in sections}
        for expected in EXPECTED_SECTIONS:
            if expected in found_names:
                findings.append(Finding(
                    severity="PASS", rule="section_present",
                    message=f"Section found: {expected}",
                ))
            else:
                findings.append(Finding(
                    severity="WARN", rule="section_missing",
                    message=f"Missing section: {expected}",
                    expected=expected, actual="not found",
                ))

        # 2. Check section order
        found_ordered = [s.name for s in sections]
        expected_filtered = [s for s in EXPECTED_SECTIONS if s in found_names]
        if found_ordered != expected_filtered:
            findings.append(Finding(
                severity="WARN", rule="section_order",
                message="Sections are not in the expected order",
                expected=", ".join(expected_filtered),
                actual=", ".join(found_ordered),
            ))
        else:
            findings.append(Finding(
                severity="PASS", rule="section_order",
                message="Sections are in the expected order",
            ))

        # 3. Check heading name corrections needed
        for section in sections:
            if section.raw_name != section.name:
                findings.append(Finding(
                    severity="WARN", rule="heading_name",
                    message=f'Heading "{section.raw_name}" should be "{section.name}"',
                    expected=section.name, actual=section.raw_name,
                    location=f"para {section.para_index}",
                ))

        # 4. Check for thin sections
        if self._style_guide:
            section_stats = self._style_guide.get("sections", {})
        else:
            section_stats = {}

        for section in sections:
            content_len = len(section.text.strip())
            typical = section_stats.get(section.name, {}).get(
                "typical_length_chars", 3000
            )

            if content_len < MIN_SECTION_LENGTH and section.name not in (
                "SETTLEMENT STATUS", "FURTHER CASE HANDLING"
            ):
                findings.append(Finding(
                    severity="WARN", rule="thin_section",
                    message=f"{section.name} is very thin ({content_len} chars, "
                            f"typical: ~{typical})",
                    expected=f"~{typical} chars", actual=f"{content_len} chars",
                ))

        # 5. Check heading formatting (bold + underline + ALL CAPS)
        for section in sections:
            try:
                heading_range = self.doc.Range(
                    section.heading_range_start, section.heading_range_end
                )
                is_bold = heading_range.Bold
                is_underline = heading_range.Underline
                text = heading_range.Text.strip().rstrip('\r')

                issues = []
                if is_bold != -1 and is_bold != True:  # noqa: E712
                    issues.append("not bold")
                if not is_underline:
                    issues.append("not underlined")
                if not text.isupper():
                    issues.append("not ALL CAPS")

                if issues:
                    findings.append(Finding(
                        severity="WARN", rule="heading_formatting",
                        message=f"{section.name} heading: {', '.join(issues)}",
                        location=f"para {section.para_index}",
                    ))
            except Exception as e:
                logger.debug(f"Could not check heading formatting for {section.name}: {e}")

        return ValidationResult(
            context="Report structural scan",
            findings=findings,
        )

    # ────── Pass 2: LLM Section Review ──────

    def _review_text(self, text: str, section_name: str, full_doc_text: str,
                     case_data: Optional[Dict], extra_texts: List[str]) -> Optional[str]:
        """Send a section to the LLM for voice/style/content review.

        Returns the revised text, or None on failure.
        """
        if not self._llm_caller:
            logger.error("No LLM caller available — skipping review")
            return None

        prompt = self._build_review_prompt(
            text, section_name, full_doc_text, case_data, extra_texts
        )

        try:
            result = self._llm_caller.call(
                prompt=prompt,
                text=text,
                task_type="summary",
                agent_id="agent_report_review",
            )
            return result.strip() if result else None
        except Exception as e:
            logger.error(f"LLM review failed for {section_name}: {e}")
            return None

    def _build_review_prompt(self, section_text: str, section_name: str,
                             full_doc_text: str, case_data: Optional[Dict],
                             extra_texts: List[str]) -> str:
        """Build the LLM prompt for reviewing a section."""

        general_style = self._style_guide.get("general", {}) if self._style_guide else {}
        hedging = general_style.get("hedging_phrases", [
            "We believe", "It is anticipated that", "It appears that",
            "Based on our review", "In our assessment",
        ])

        # Examples
        example_block = ""
        examples = self._section_examples.get(section_name, [])
        if examples:
            for i, ex in enumerate(examples[:2], 1):
                truncated = ex[:10000] if len(ex) > 10000 else ex
                example_block += (
                    f"\n=== EXAMPLE {i} — THIS IS THE TARGET VOICE AND STYLE ===\n"
                    f"{truncated}\n"
                    f"=== END EXAMPLE {i} ===\n"
                )

        # Full document context (truncated)
        doc_context = ""
        if full_doc_text:
            truncated_doc = full_doc_text[:15000] if len(full_doc_text) > 15000 else full_doc_text
            doc_context = (
                "\n=== FULL DOCUMENT (for cross-reference context only) ===\n"
                f"{truncated_doc}\n"
                "=== END FULL DOCUMENT ===\n"
            )

        # Case data context
        case_context = ""
        if case_data:
            sections_data = case_data.get("sections", {})
            relevant = sections_data.get(section_name, "")
            if relevant:
                truncated_case = relevant[:8000] if len(relevant) > 8000 else relevant
                case_context = (
                    "\n=== CASE FILE DATA (facts that may be missing from the draft) ===\n"
                    f"{truncated_case}\n"
                    "=== END CASE FILE DATA ===\n"
                )

        # Extra documents context
        extra_context = ""
        if extra_texts:
            combined = "\n---\n".join(t[:5000] for t in extra_texts)
            if len(combined) > 12000:
                combined = combined[:12000] + "\n[truncated]"
            extra_context = (
                "\n=== ADDITIONAL REFERENCE DOCUMENTS ===\n"
                f"{combined}\n"
                "=== END ADDITIONAL DOCUMENTS ===\n"
            )

        prompt = f"""You are reviewing the "{section_name}" section of a litigation report written by a junior associate. Your task is to revise it to match the senior attorney's voice, tone, style, and thoroughness.

IMPORTANT — REDLINE MODE INSTRUCTIONS:
Your output will be compared word-by-word against the original text to generate Track Changes in Word. Follow these rules:
1. Preserve any text that does not need to change EXACTLY as-is — same wording, same sentence structure.
2. PRESERVE THE EXACT PARAGRAPH STRUCTURE — keep the same number of paragraphs and blank lines. Do NOT merge or split paragraphs.
3. Only modify the specific words, sentences, or passages that need improvement.
4. If a sentence is fine as-is, copy it verbatim.
5. Do NOT add commentary, explanations, or markers — output ONLY the revised section text.

=== VOICE AND STYLE (match the examples precisely) ===
1. Write from defense counsel's perspective, addressing the insurance carrier.
2. Use hedging phrases naturally: {', '.join(hedging)}.
3. Formal, professional tone. Refer to parties as "Plaintiff" and "Defendant" (not first names).
4. Complete, polished paragraphs. No bullet points in body text.
5. Sub-headings on their own line, separated by blank lines.
6. No markdown (no **, ##, *). Plain text only.
{example_block}
{doc_context}
{case_context}
{extra_context}
=== THE ASSOCIATE'S DRAFT TO REVIEW AND REVISE ===
{section_text}

=== INSTRUCTIONS ===
Revise the above section text to match the senior attorney's voice and style shown in the examples. Fix any formatting issues. {"Add any significant facts or details from the case file data that the associate missed." if case_data else ""}
Preserve the associate's correct analysis — only change what needs improvement.
Output ONLY the revised section text, nothing else."""

        return prompt.strip()

    # ────── Redline Application ──────

    def _apply_redline(self, range_start: int, range_end: int,
                       original_text: str, revised_text: str) -> int:
        """Apply Track Changes for a section. Returns number of changes applied."""
        try:
            from icharlotte_core.word_hotkey import apply_flat_diff_redline
            from icharlotte_core.word_validator import validate_after_edit

            r = apply_flat_diff_redline(
                doc_com=self.doc,
                range_start=range_start,
                range_end=range_end,
                original_text=original_text,
                revised_text=revised_text,
                auto_enable_track_changes=True,
            )

            if r['success']:
                # Lightweight post-edit validation
                time.sleep(0.3)
                try:
                    val = validate_after_edit(self.doc, range_start, range_end)
                    if val.has_errors:
                        val.print_summary()
                        logger.warning("Post-redline validation found errors")
                except Exception as ve:
                    logger.debug(f"Validation error: {ve}")

                return r['total_changes']

            logger.warning("Redline application returned failure")
            return 0

        except Exception as e:
            logger.error(f"Failed to apply redline: {e}")
            return 0

    def _fix_heading_name(self, section: SectionInfo):
        """Apply a tracked change to correct a heading name."""
        try:
            heading_range = self.doc.Range(
                section.heading_range_start, section.heading_range_end
            )
            text = heading_range.Text.strip().rstrip('\r')

            # Enable track changes
            if not self.doc.TrackRevisions:
                self.doc.TrackRevisions = True

            # Replace heading text (preserving paragraph mark)
            content_range = self.doc.Range(
                section.heading_range_start,
                section.heading_range_end - 1  # exclude \r
            )
            content_range.Text = section.name
            logger.info(f"Fixed heading: '{text}' -> '{section.name}'")

        except Exception as e:
            logger.warning(f"Could not fix heading name for {section.name}: {e}")

    # ────── Resource Loading ──────

    def _load_style_resources(self):
        """Load style guide, section examples, and LLM caller."""
        try:
            from Scripts.report_generator.style_library import (
                get_style_guide, get_section_examples, REPORT_SECTIONS
            )
            self._style_guide = get_style_guide()
            for section in REPORT_SECTIONS:
                exs = get_section_examples(section, max_examples=2)
                if exs:
                    self._section_examples[section] = exs
            logger.info(f"Loaded style guide + {len(self._section_examples)} section examples")
        except Exception as e:
            logger.warning(f"Could not load style resources: {e}")

        try:
            from icharlotte_core.llm_config import LLMCaller
            self._llm_caller = LLMCaller()
        except Exception as e:
            logger.error(f"Could not create LLM caller: {e}")

    def _load_case_data(self) -> Optional[Dict]:
        """Load case data using gather.py."""
        if not self.case_path:
            return None

        try:
            # Extract file number from case path
            match = re.search(r'(\d{4}\.\d{3})', self.case_path)
            if not match:
                logger.warning(f"Could not extract file number from: {self.case_path}")
                return None

            file_number = match.group(1)
            from Scripts.report_generator.gather import gather_case_data
            data = gather_case_data(file_number, "FSR")
            logger.info(f"Loaded case data for {file_number}: "
                        f"{len(data.get('sections', {}))} sections")
            return data
        except Exception as e:
            logger.error(f"Failed to load case data: {e}")
            return None

    def _extract_extra_docs(self, paths: List[str]) -> List[str]:
        """Extract text from additional document files."""
        texts = []
        for path in paths:
            try:
                ext = os.path.splitext(path)[1].lower()
                text = ""

                if ext in ('.docx',):
                    try:
                        from docx import Document
                        doc = Document(path)
                        text = "\n".join(p.text for p in doc.paragraphs)
                    except Exception:
                        pass

                elif ext in ('.pdf',):
                    try:
                        from icharlotte_core.document_processor import DocumentProcessor
                        processor = DocumentProcessor()
                        text = processor.extract_text(path)
                    except Exception:
                        pass

                elif ext in ('.doc',):
                    try:
                        import win32com.client
                        import pythoncom
                        pythoncom.CoInitialize()
                        try:
                            word = win32com.client.Dispatch("Word.Application")
                            doc = word.Documents.Open(path, ReadOnly=True, Visible=False)
                            text = doc.Content.Text or ""
                            doc.Close(False)
                        finally:
                            pythoncom.CoUninitialize()
                    except Exception:
                        pass

                if text.strip():
                    texts.append(text.strip())
                    logger.info(f"Extracted {len(text)} chars from {os.path.basename(path)}")
                else:
                    logger.warning(f"No text extracted from {os.path.basename(path)}")

            except Exception as e:
                logger.error(f"Error extracting {path}: {e}")

        return texts

    # ────── Helpers ──────

    def _get_full_doc_text(self) -> str:
        """Get the full document text via COM."""
        try:
            return self.doc.Content.Text or ""
        except Exception as e:
            logger.error(f"Could not get document text: {e}")
            return ""

    def _get_range_text(self, start: int, end: int) -> str:
        """Get text from a specific range."""
        try:
            return self.doc.Range(start, end).Text or ""
        except Exception as e:
            logger.error(f"Could not get range text: {e}")
            return ""
