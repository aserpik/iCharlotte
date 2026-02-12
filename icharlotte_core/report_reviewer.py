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
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


def _com_retry(func, retries=3, delay=2.0):
    """Retry a callable on COM RPC_E_CALL_REJECTED (-2147418111)."""
    import pywintypes
    for attempt in range(retries):
        try:
            return func()
        except pywintypes.com_error as e:
            if e.hresult == -2147418111 and attempt < retries - 1:
                logger.warning(f"COM busy (attempt {attempt + 1}), "
                               f"retrying in {delay}s...")
                time.sleep(delay)
            else:
                raise

# COM retry settings — Word sometimes rejects calls when busy processing


# ──────────────────────────────────────────────────────────────────────
# Data classes
# ──────────────────────────────────────────────────────────────────────

@dataclass
class SubheadingInfo:
    """A detected L1 or L2 subheading within a section."""
    level: str              # "L1" or "L2"
    text: str               # Full text of the subheading line
    range_start: int        # Character position start
    range_end: int          # Character position end


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
    subheadings: List[SubheadingInfo] = field(default_factory=list)


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

# Sections containing tabular data — skip LLM review, tables get mangled
TABULAR_SECTIONS = {"LITIGATION BUDGET"}

# Minimum preservation ratio — if the LLM rewrites more than this, reject it
MIN_PRESERVATION_RATIO = 0.25

# Minimum chars for a subsection to be worth reviewing individually
MIN_SUBSECTION_LENGTH = 100

# Max parallel LLM calls for subsection review
MAX_PARALLEL_LLM_CALLS = 6


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
            only_sections: Optional[List[str]] = None,
            skip_structural: bool = False,
            skip_final_validation: bool = False,
            ) -> Tuple[Optional['ValidationResult'], int]:
        """Main entry point for the report review.

        Args:
            selection_range: (start, end) character positions for selection-only review.
                             If None, reviews the full document.
            include_case_data: Whether to load case data from AI OUTPUT folder
            extra_doc_paths: Additional document paths to extract text from
            progress_callback: Called with progress messages for UI display
            only_sections: If set, only review sections whose names contain
                           one of these strings (case-insensitive partial match)
            skip_structural: Skip Pass 1 structural scan

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
                        sel_text, revised, "SELECTED TEXT"
                    )
                    total_changes += changes
            progress(f"Review complete — {total_changes} tracked changes applied.")
            return None, total_changes

        # Full document mode
        # Pass 1: Structural scan
        structural = None
        sections = self.detect_sections()
        if skip_structural:
            progress(f"Pass 1: Skipped (found {len(sections)} sections)")
        else:
            progress("Pass 1: Scanning document structure...")
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
        # CRITICAL: After each redline application, character positions shift.
        # We re-detect sections after each edit to get fresh ranges.
        reviewable_names = [s.name for s in sections
                            if s.text.strip()
                            and s.raw_name not in TABULAR_SECTIONS
                            and s.name not in TABULAR_SECTIONS]

        # Filter to requested sections if specified
        if only_sections:
            only_upper = [s.upper() for s in only_sections]
            reviewable_names = [
                name for name in reviewable_names
                if any(filt in name for filt in only_upper)
            ]
            progress(f"Filtered to {len(reviewable_names)} sections: "
                     f"{', '.join(reviewable_names)}")

        for i, section_name in enumerate(reviewable_names, 1):
            section_label = f"{section_name} ({i}/{len(reviewable_names)})"
            try:
                # Re-detect sections to get fresh character positions
                current_sections = self.detect_sections()
                section = None
                for s in current_sections:
                    if s.name == section_name:
                        section = s
                        break
                if not section:
                    progress(f"  {section_name}: could not re-locate, skipping")
                    continue

                fresh_text = self._get_range_text(section.range_start, section.range_end)
                if not fresh_text.strip():
                    progress(f"  {section_name}: empty, skipping")
                    continue

                progress(f"Pass 2: Reviewing {section_label}...")

                # Choose review strategy based on subheadings
                if section.subheadings:
                    # Subsection-level parallel review
                    changes = self._review_subsections_parallel(
                        section, full_doc_text, case_data, extra_texts, progress
                    )
                    total_changes += changes
                else:
                    # Single-block review (no subheadings)
                    revised = self._review_text(
                        fresh_text, section_name, full_doc_text,
                        case_data, extra_texts
                    )

                    if not revised or revised.strip() == fresh_text.strip():
                        progress(f"  {section_name}: no changes needed")
                        continue

                    preservation = self._calc_preservation(fresh_text, revised)
                    if preservation < MIN_PRESERVATION_RATIO:
                        progress(f"  {section_name}: LLM rewrote too aggressively "
                                 f"({preservation:.0%} preserved, min {MIN_PRESERVATION_RATIO:.0%}), skipping")
                        logger.warning(f"Rejected LLM output for {section_name}: "
                                       f"only {preservation:.0%} preserved")
                        continue

                    revised = self._restore_subheadings_in_revised(fresh_text, revised)

                    dropped = self._check_subheadings_preserved(fresh_text, revised)
                    if dropped:
                        progress(f"  {section_name}: LLM dropped subheadings "
                                 f"({', '.join(dropped)}), skipping")
                        logger.warning(f"Rejected LLM output for {section_name}: "
                                       f"dropped subheadings: {dropped}")
                        continue

                    progress(f"  Applying redlines to {section_name} "
                             f"({preservation:.0%} preserved)...")
                    changes = self._apply_redline(
                        section.range_start, section.range_end,
                        fresh_text, revised, section_name
                    )
                    total_changes += changes

                # Fix subheading formatting (bold, underline, indent)
                if total_changes > 0 and section.subheadings:
                    time.sleep(1.0)  # let Word settle before formatting pass
                    self._fix_subheading_formatting(section)

            except Exception as e:
                logger.error(f"Error reviewing {section_name}: {e}")
                progress(f"  {section_name}: error ({e}), continuing...")

        # Final full-document validation
        if not skip_final_validation:
            progress("Running final validation...")
            final_sections = self.detect_sections()
            self._run_final_validation(final_sections, total_changes)
        else:
            progress("Skipping final validation.")

        progress(f"Review complete — {len(reviewable_names)} sections reviewed, "
                 f"{total_changes} tracked changes applied.")
        return structural, total_changes

    # ────── Pass 1: Structural Scan ──────

    def detect_sections(self) -> List[SectionInfo]:
        """Scan the document for section headings via COM.

        Uses bulk text read (1 COM call) + Python-side filtering to find
        candidates, then confirms with COM only for the ~10-30 matches.
        This is ~100x faster than iterating all paragraphs via COM.
        """
        from icharlotte_core.word_validator import fuzzy_match_section_heading

        t0 = time.time()

        # 1. Bulk read: one COM call for full document text
        full_text = _com_retry(lambda: self.doc.Content.Text) or ""
        paragraphs = full_text.split('\r')
        logger.info(f"Bulk scan: {len(paragraphs)} paragraphs in "
                     f"{time.time() - t0:.2f}s")

        # 2. Python-side: find candidate section headings by text pattern
        candidates = []  # (para_index_0based, canonical_name, raw_text)
        for i, raw_line in enumerate(paragraphs):
            text = raw_line.strip()
            if not text or len(text) > 80:
                continue
            if text.isupper():
                canonical = fuzzy_match_section_heading(text)
                if canonical:
                    candidates.append((i, canonical, text))

        # 3. COM confirmation: only touch the ~10 candidate paragraphs
        #    Use Find to locate each heading by exact text and get positions
        heading_matches = []  # (canonical, raw_text, heading_start, heading_end)

        for _, canonical, raw_text in candidates:
            try:
                # Use Find on doc content to locate the heading text
                search_range = _com_retry(lambda: self.doc.Content)
                search_range.Find.ClearFormatting()
                search_range.Find.Text = raw_text
                search_range.Find.Font.Bold = True
                search_range.Find.Forward = True
                search_range.Find.Wrap = 0  # wdFindStop

                # Find all occurrences (in case of duplicates, take first
                # that hasn't already been claimed)
                claimed_starts = {h[2] for h in heading_matches}
                found = False
                while search_range.Find.Execute():
                    if search_range.Start not in claimed_starts:
                        # Verify it's actually a paragraph on its own
                        # (not embedded in body text)
                        para_text = search_range.Paragraphs(1).Range.Text.strip().rstrip('\r')
                        if para_text == raw_text:
                            heading_matches.append((
                                canonical, raw_text,
                                search_range.Paragraphs(1).Range.Start,
                                search_range.Paragraphs(1).Range.End,
                            ))
                            logger.info(f"  Found: '{raw_text}' -> '{canonical}'")
                            found = True
                            break
                    # Move past this match
                    search_range.Start = search_range.End

                if not found:
                    logger.debug(f"  Could not confirm heading: '{raw_text}'")

            except Exception as e:
                logger.debug(f"Error confirming heading '{raw_text}': {e}")

        # Sort by document position
        heading_matches.sort(key=lambda h: h[2])

        # 4. Build SectionInfo with content ranges
        sections = []
        doc_end = _com_retry(lambda: self.doc.Content.End)

        for idx, (canonical, raw, h_start, h_end) in enumerate(heading_matches):
            content_start = h_end
            if idx + 1 < len(heading_matches):
                content_end = heading_matches[idx + 1][2]
            else:
                content_end = doc_end

            try:
                section_text = _com_retry(
                    lambda s=content_start, e=content_end:
                    self.doc.Range(s, e).Text or ""
                )
            except Exception:
                section_text = ""

            # 5. Detect subheadings within this section's text
            subheadings = self._detect_subheadings_bulk(
                section_text, content_start, content_end)

            sections.append(SectionInfo(
                name=canonical,
                raw_name=raw,
                para_index=0,  # no longer tracking para index
                range_start=content_start,
                range_end=content_end,
                text=section_text,
                heading_range_start=h_start,
                heading_range_end=h_end,
                subheadings=subheadings,
            ))

        total_subs = sum(len(s.subheadings) for s in sections)
        elapsed = time.time() - t0
        logger.info(f"Detected {len(sections)} sections, "
                     f"{total_subs} subheadings in {elapsed:.2f}s")
        return sections

    def _detect_subheadings_bulk(self, section_text: str,
                                  section_start: int,
                                  section_end: int) -> List[SubheadingInfo]:
        """Detect L1/L2 subheadings within a section.

        Two detection paths:
        1. Text-based: lines starting with "A.\t" (L1) or "1.\t" (L2)
        2. COM-based: bold + listType=3 + short + not ALL CAPS (Word applies
           the letter/number prefix via list formatting, not in the text)

        Uses bulk text to find candidates first, then confirms with COM.
        """
        if not section_text:
            return []

        subheadings = []
        claimed_starts = set()

        # Path 1: Text-based detection (explicit "A.\t" / "1.\t" in text)
        text_candidates = []
        for line in section_text.split('\r'):
            text = line.strip()
            if not text or len(text) > 80 or text.isupper():
                continue
            if re.match(r'^[A-Z]\.[\t ]', text):
                text_candidates.append(("L1", text))
            elif re.match(r'^[0-9]+\.[\t ]', text):
                text_candidates.append(("L2", text))

        for level, raw_text in text_candidates:
            try:
                search_range = _com_retry(
                    lambda s=section_start, e=section_end:
                    self.doc.Range(s, e)
                )
                search_range.Find.ClearFormatting()
                search_text = raw_text[:40] if len(raw_text) > 40 else raw_text
                search_range.Find.Text = search_text
                search_range.Find.Forward = True
                search_range.Find.Wrap = 0

                while search_range.Find.Execute():
                    if search_range.Start not in claimed_starts:
                        para = search_range.Paragraphs(1)
                        p_start = para.Range.Start
                        p_end = para.Range.End
                        claimed_starts.add(p_start)
                        subheadings.append(SubheadingInfo(
                            level=level, text=raw_text,
                            range_start=p_start, range_end=p_end,
                        ))
                        break
                    search_range.Start = search_range.End
            except Exception as e:
                logger.debug(f"Error finding subheading '{raw_text[:30]}': {e}")

        # Path 2: Word list-formatted subheadings (bold + listType=3)
        # The "A.", "B.", "1.", "2." prefixes are rendered by Word, not in text.
        # Candidate: short label line (not a sentence), not ALL CAPS, not empty.
        # Filter tightly to avoid checking hundreds of body text lines via COM.
        list_candidates = []
        for line in section_text.split('\r'):
            text = line.strip()
            if not text or text.isupper():
                continue
            # Subheading labels are typically short (< 45 chars),
            # don't end with sentence-ending punctuation,
            # contain at least one alphabetic word, and have 1-6 words.
            words = text.split()
            has_alpha = any(w[0].isalpha() for w in words if w)
            if (len(text) <= 45
                    and 1 <= len(words) <= 6
                    and has_alpha
                    and not text.endswith(('.', ',', ';', ':', '!', '?'))
                    and not text.startswith('$')
                    and text not in [c[1] for c in text_candidates]):
                list_candidates.append(text)

        if list_candidates:
            # Use Find with bold constraint to locate each candidate
            for raw_text in list_candidates:
                try:
                    search_range = _com_retry(
                        lambda s=section_start, e=section_end:
                        self.doc.Range(s, e)
                    )
                    search_range.Find.ClearFormatting()
                    search_range.Find.Text = raw_text
                    search_range.Find.Font.Bold = True
                    search_range.Find.Forward = True
                    search_range.Find.Wrap = 0

                    while search_range.Find.Execute():
                        para = search_range.Paragraphs(1)
                        p_start = para.Range.Start
                        if p_start in claimed_starts:
                            search_range.Start = search_range.End
                            continue

                        # Verify: must be the full paragraph text (not
                        # a bold word inside body text)
                        para_text = para.Range.Text.strip().rstrip('\r')
                        if para_text != raw_text:
                            search_range.Start = search_range.End
                            continue

                        # Must be list-formatted (outline numbered)
                        try:
                            lt = para.Range.ListFormat.ListType
                        except Exception:
                            lt = 0
                        if lt != 3:
                            search_range.Start = search_range.End
                            continue

                        # Classify L1 vs L2 using list NumberStyle
                        level = None
                        try:
                            ns = para.Range.ListFormat.ListTemplate.ListLevels(1).NumberStyle
                            if ns == 3:   # wdListNumberStyleUppercaseLetter
                                level = "L1"
                            elif ns == 0:  # wdListNumberStyleArabic
                                level = "L2"
                        except Exception:
                            pass

                        # Fallback: formatting heuristics
                        if not level:
                            try:
                                is_underlined = (para.Range.Underline == 1)
                                left = para.Format.LeftIndent
                                if is_underlined or left <= 40:
                                    level = "L1"
                                else:
                                    level = "L2"
                            except Exception:
                                level = "L1"

                        p_end = para.Range.End
                        claimed_starts.add(p_start)
                        subheadings.append(SubheadingInfo(
                            level=level, text=raw_text,
                            range_start=p_start, range_end=p_end,
                        ))
                        break

                except Exception as e:
                    logger.debug(f"Error finding list subheading '{raw_text[:30]}': {e}")

        # Sort by position
        subheadings.sort(key=lambda s: s.range_start)
        return subheadings

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

        # 3. Check for thin sections
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

        # 4. Check heading formatting (bold + underline + ALL CAPS)
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

        # 5. Check subheading formatting using pre-detected subheadings
        for section in sections:
            for sub in section.subheadings:
                try:
                    para_range = self.doc.Range(sub.range_start, sub.range_end)
                    para = para_range.Paragraphs(1)
                    label = sub.text[:40]
                    issues = []

                    if sub.level == "L1":
                        if para.Range.Underline != 1:
                            issues.append("not underlined")
                        try:
                            lf = para.Range.ListFormat
                            if lf.ListType == 3:
                                tp = lf.ListTemplate.ListLevels(1).TextPosition
                                if abs(tp - 72) > 2:
                                    issues.append(f"list text position {tp:.0f}pt (expected 72pt)")
                        except Exception:
                            pass

                    elif sub.level == "L2":
                        r = para.Range
                        raw = r.Text.rstrip('\r')
                        text_offset = 0
                        for ch in raw:
                            if ch.isalpha():
                                break
                            text_offset += 1
                        if text_offset < len(raw):
                            text_range = self.doc.Range(
                                r.Start + text_offset, r.End - 1)
                            if text_range.Underline != 1:
                                issues.append("text not underlined")

                    if issues:
                        findings.append(Finding(
                            severity="WARN", rule="subheading_formatting",
                            message=f"{section.name} {sub.level} subheading '{label}': {', '.join(issues)}",
                        ))
                except Exception as e:
                    logger.debug(f"Error checking subheading '{sub.text[:30]}' in {section.name}: {e}")

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

        # Sections that require substantive analytical review
        is_analytical = section_name in (
            "EVALUATION OF LIABILITY",
            "EVALUATION OF EXPOSURE",
            "SETTLEMENT STATUS",
        )

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

        # Full document context
        # Analytical sections get the FULL document (no truncation) since they
        # need to cross-reference facts from all other sections
        doc_context = ""
        if full_doc_text:
            if is_analytical:
                truncated_doc = full_doc_text
                context_label = "for substantive cross-reference — READ CAREFULLY"
            else:
                truncated_doc = full_doc_text[:15000] if len(full_doc_text) > 15000 else full_doc_text
                context_label = "for cross-reference context only"
            doc_context = (
                f"\n=== FULL DOCUMENT ({context_label}) ===\n"
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

        substantive_block = ""
        if is_analytical:
            substantive_block = """
SUBSTANTIVE REVIEW — ANALYTICAL SECTIONS:
This is an analytical section that requires SUBSTANTIVE review, not just style review. Using the FULL DOCUMENT as context (especially the INVESTIGATION, DISCOVERY, and FACTUAL BACKGROUND sections), evaluate the associate's legal analysis and make corrections where warranted:

- Does the analysis accurately reflect the facts described elsewhere in the report? Flag and correct any misstatements or inconsistencies with the factual record.
- Is the legal reasoning sound and complete? Are all relevant elements addressed (e.g., for liability: duty, breach, causation, damages)?
- Are the conclusions supported by the evidence discussed in other sections? Strengthen weak arguments by citing specific facts from the report.
- Are there obvious gaps — important facts from the INVESTIGATION or DISCOVERY sections that should be discussed here but are missing? Add brief references where appropriate.
- Is the analysis balanced from a defense perspective? Ensure favorable facts are highlighted and unfavorable facts are properly hedged.
- Are damage figures, settlement values, or exposure ranges reasonable given the facts? Flag any that seem unsupported.

When making substantive changes, integrate them naturally into the existing text. Do NOT add new sub-sections or dramatically restructure the analysis. Insert missing analysis as additional sentences within existing paragraphs where they fit logically.
"""

        prompt = f"""You are reviewing the "{section_name}" section of a litigation report written by a junior associate. Your task is to make TARGETED, SURGICAL improvements to match the senior attorney's voice, tone, and style{", AND to substantively review the legal analysis" if is_analytical else ""}.

CRITICAL — REDLINE MODE RULES:
Your output will be compared character-by-character against the original to generate Track Changes in Microsoft Word. You MUST follow these rules precisely:

1. DO NOT include the section heading (e.g., "{section_name}") in your output. You are revising ONLY the body content BELOW the heading.
2. DO NOT include any sub-headings from the NEXT section. Only output content for THIS section.
3. PRESERVE the original text as much as possible. Copy sentences that are acceptable VERBATIM — do not rephrase text that is already adequate.
4. PRESERVE THE EXACT PARAGRAPH STRUCTURE — same number of paragraphs. Do NOT merge or split paragraphs.
5. PRESERVE all tabular data, lists of medical providers, billing summaries, dollar amounts, dates, and factual details EXACTLY as written — copy them character-for-character. Do NOT insert spaces into proper nouns, provider names, or dollar amounts (e.g., "American" must stay "American", not "A merican").
6. PRESERVE ALL SUB-HEADINGS EXACTLY AS WRITTEN. Any line that is a lettered sub-heading (e.g., "A.\tSomething", "B.\tSomething") or numbered sub-heading (e.g., "1.\tSomething", "2.\tSomething") MUST appear in your output VERBATIM on its own line, character-for-character identical including tabs. Do NOT modify, rephrase, reorder, merge, or remove sub-headings.{self._format_subheading_list(section_text)}
7. Make only targeted word/phrase/sentence changes where the voice, tone, or style clearly needs improvement.
8. If a section is already well-written, output it unchanged. It is perfectly acceptable to make zero changes.
9. Do NOT add commentary, explanations, or markers — output ONLY the revised section body text.
{substantive_block}
WHAT TO IMPROVE (only where clearly needed):
- Shift passive or casual phrasing to defense counsel's formal voice
- Add hedging language where appropriate: {', '.join(hedging)}
- Refer to parties as "Plaintiff" and "Defendant" (not first names)
- Fix grammatical errors, awkward phrasing, or unclear sentences
- Ensure content flows logically
{f"- Strengthen or correct the legal analysis as described in SUBSTANTIVE REVIEW above" if is_analytical else ""}

WHAT TO LEAVE ALONE:
- Factual content, dates, dollar amounts, case numbers
- Tables, lists, and structured data
- Sub-headings within the section
- Text that is already professional and well-written
- Paragraph structure and line breaks
{example_block}
{doc_context}
{case_context}
{extra_context}
=== THE ASSOCIATE'S DRAFT SECTION BODY (below the "{section_name}" heading) ===
{section_text}

=== INSTRUCTIONS ===
Revise the above section body text with MINIMAL, TARGETED changes. {"Add any significant facts from the case file data that the associate missed, but integrate them naturally — do not reorganize existing text." if case_data else ""}
{"IMPORTANT: This is an analytical section. In addition to style improvements, substantively review the analysis against the facts in the full document. Correct errors, strengthen weak reasoning, and add references to specific facts where the analysis is unsupported." if is_analytical else "Preserve the associate's analysis and structure — only change what genuinely needs improvement."}
Output ONLY the revised section body text, nothing else. Do NOT include the section heading."""

        return prompt.strip()

    # ────── Subsection-Level Review ──────

    @staticmethod
    def _split_into_subsections(section: SectionInfo) -> List[Tuple[str, int, int, str]]:
        """Split a section into subsection chunks using detected subheadings.

        Returns list of (name, range_start, range_end, text) tuples.
        Includes an intro chunk (text before the first subheading) if non-trivial.
        """
        chunks = []

        if not section.subheadings:
            # No subheadings — return the whole section as one chunk
            return [("FULL", section.range_start, section.range_end, section.text)]

        # Intro: section content start → first subheading start
        first_sub = section.subheadings[0]
        if first_sub.range_start > section.range_start:
            intro_len = first_sub.range_start - section.range_start
            if intro_len > 0:
                # Text from section start to first subheading
                intro_text = section.text[:intro_len] if len(section.text) >= intro_len else section.text
                chunks.append((
                    "INTRO",
                    section.range_start,
                    first_sub.range_start,
                    intro_text,
                ))

        # Each subheading → next subheading (or section end)
        for i, sub in enumerate(section.subheadings):
            # Content starts at the subheading's paragraph end
            content_start = sub.range_end
            # Content ends at next subheading start, or section end
            if i + 1 < len(section.subheadings):
                content_end = section.subheadings[i + 1].range_start
            else:
                content_end = section.range_end

            # Calculate text offset within section.text
            text_offset_start = sub.range_start - section.range_start
            text_offset_end = content_end - section.range_start

            sub_text = section.text[text_offset_start:text_offset_end] if text_offset_end <= len(section.text) else ""

            chunks.append((
                sub.text,  # subsection name (e.g., "Causation")
                sub.range_start,  # include the subheading itself in the range
                content_end,
                sub_text,
            ))

        return chunks

    def _review_subsections_parallel(
        self, section: SectionInfo, full_doc_text: str,
        case_data: Optional[Dict], extra_texts: List[str],
        progress: Callable[[str], None],
    ) -> int:
        """Review a section's subsections in parallel, then apply redlines sequentially.

        Returns total number of tracked changes applied.
        """
        chunks = self._split_into_subsections(section)

        # Filter out tiny chunks
        reviewable = [
            (name, rs, re_, text)
            for name, rs, re_, text in chunks
            if len(text.strip()) >= MIN_SUBSECTION_LENGTH
        ]

        if not reviewable:
            progress(f"  {section.name}: no reviewable subsections")
            return 0

        progress(f"  {section.name}: reviewing {len(reviewable)} subsections in parallel...")

        # Fire LLM calls in parallel
        results = {}  # name → (original, revised, range_start, range_end)

        def _do_review(name, text):
            """Thread target: call LLM for one subsection."""
            # For subsection review, pass the full section as context
            prompt = self._build_subsection_review_prompt(
                subsection_text=text,
                subsection_name=name,
                section_name=section.name,
                full_section_text=section.text,
                full_doc_text=full_doc_text,
                case_data=case_data,
                extra_texts=extra_texts,
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
                logger.error(f"LLM review failed for {section.name}/{name}: {e}")
                return None

        with ThreadPoolExecutor(max_workers=MAX_PARALLEL_LLM_CALLS) as pool:
            futures = {}
            for name, rs, re_, text in reviewable:
                future = pool.submit(_do_review, name, text)
                futures[future] = (name, rs, re_, text)

            for future in as_completed(futures):
                name, rs, re_, original = futures[future]
                try:
                    revised = future.result()
                    if revised and revised.strip() != original.strip():
                        # Preservation check
                        preservation = self._calc_preservation(original, revised)
                        if preservation < MIN_PRESERVATION_RATIO:
                            progress(f"    {name}: too aggressive "
                                     f"({preservation:.0%} preserved), skipping")
                            continue
                        results[name] = (original, revised, rs, re_, preservation)
                        progress(f"    {name}: ready ({preservation:.0%} preserved)")
                    else:
                        progress(f"    {name}: no changes needed")
                except Exception as e:
                    logger.error(f"Error in parallel review for {name}: {e}")

        if not results:
            progress(f"  {section.name}: no subsection changes to apply")
            return 0

        # Apply redlines in REVERSE document order to avoid position shifts
        ordered = sorted(results.values(), key=lambda r: r[2], reverse=True)
        total_changes = 0

        for original, revised, rs, re_, preservation in ordered:
            # Re-read the current text at this range (positions may have shifted
            # from later-in-document edits, but since we go in reverse, they shouldn't)
            current_text = self._get_range_text(rs, re_)
            if not current_text.strip():
                continue

            # Use current_text if it still matches original, otherwise skip
            # (safety check in case something shifted unexpectedly)
            current_preservation = self._calc_preservation(original, current_text)
            if current_preservation < 0.9:
                logger.warning(f"Subsection text shifted unexpectedly, skipping")
                continue

            changes = self._apply_redline(rs, re_, current_text, revised,
                                          f"{section.name}/subsection")
            total_changes += changes

        return total_changes

    def _build_subsection_review_prompt(
        self, subsection_text: str, subsection_name: str,
        section_name: str, full_section_text: str,
        full_doc_text: str, case_data: Optional[Dict],
        extra_texts: List[str],
    ) -> str:
        """Build LLM prompt for reviewing a single subsection."""

        is_analytical = section_name in (
            "EVALUATION OF LIABILITY",
            "EVALUATION OF EXPOSURE",
            "SETTLEMENT STATUS",
        )

        general_style = self._style_guide.get("general", {}) if self._style_guide else {}
        hedging = general_style.get("hedging_phrases", [
            "We believe", "It is anticipated that", "It appears that",
            "Based on our review", "In our assessment",
        ])

        # Section context (read-only)
        section_context = (
            f"\n=== FULL SECTION CONTEXT (read-only — do NOT include other subsections in output) ===\n"
            f"{full_section_text[:15000]}\n"
            f"=== END FULL SECTION CONTEXT ===\n"
        )

        # Full document context for analytical sections
        doc_context = ""
        if full_doc_text and is_analytical:
            doc_context = (
                f"\n=== FULL DOCUMENT (for substantive cross-reference) ===\n"
                f"{full_doc_text}\n"
                f"=== END FULL DOCUMENT ===\n"
            )
        elif full_doc_text:
            truncated = full_doc_text[:10000]
            doc_context = (
                f"\n=== DOCUMENT CONTEXT (for reference only) ===\n"
                f"{truncated}\n"
                f"=== END DOCUMENT CONTEXT ===\n"
            )

        substantive_block = ""
        if is_analytical:
            substantive_block = """
SUBSTANTIVE REVIEW:
This is an analytical subsection. Using the full document and section context, evaluate the legal analysis:
- Does it accurately reflect the facts? Flag and correct misstatements.
- Is the legal reasoning sound and complete?
- Are conclusions supported by evidence? Strengthen weak arguments with specific facts.
- Is the analysis balanced from a defense perspective?
"""

        prompt = f"""You are reviewing the "{subsection_name}" subsection within the "{section_name}" section of a litigation report.

CRITICAL RULES:
1. You are revising ONLY this subsection. Do NOT include text from other subsections.
2. If the subsection starts with a sub-heading line (e.g., "Causation", "Damages"), include it VERBATIM as the first line.
3. PRESERVE the original text as much as possible. Copy sentences that are acceptable VERBATIM.
4. PRESERVE THE EXACT PARAGRAPH STRUCTURE — same number of paragraphs.
5. PRESERVE all factual data, dates, dollar amounts, provider names, billing details EXACTLY as written.
6. Make only targeted word/phrase/sentence changes where voice, tone, or style clearly needs improvement.
7. If the text is already well-written, output it unchanged.
8. Do NOT add commentary or markers — output ONLY the revised subsection text.
{substantive_block}
WHAT TO IMPROVE (only where clearly needed):
- Shift passive or casual phrasing to defense counsel's formal voice
- Add hedging language where appropriate: {', '.join(hedging)}
- Refer to parties as "Plaintiff" and "Defendant" (not first names)
- Fix grammatical errors, awkward phrasing, or unclear sentences
{f"- Strengthen or correct the legal analysis as described above" if is_analytical else ""}

WHAT TO LEAVE ALONE:
- Factual content, dates, dollar amounts, case numbers
- Tables, lists, and structured data
- Text that is already professional and well-written
- Paragraph structure and line breaks
{section_context}
{doc_context}
=== THE SUBSECTION TO REVISE ===
{subsection_text}

=== INSTRUCTIONS ===
Revise the above subsection with MINIMAL, TARGETED changes. Output ONLY the revised subsection text."""

        return prompt.strip()

    # ────── Redline Application ──────

    def _format_subheading_list(self, section_text: str) -> str:
        """Format a list of subheadings found in section text for the prompt."""
        subs = self._extract_subheadings(section_text)
        if not subs:
            return ""
        lines = "\n".join(f"   - {s}" for s in subs)
        return f"\n   The following sub-headings appear in this section and MUST be preserved exactly:\n{lines}"

    # Regex for subheading lines: "A.\tTitle" or "1.\tTitle" (with optional
    # leading whitespace and tab/spaces after the marker)
    _SUBHEADING_RE = re.compile(
        r'^\s*([A-Z]\.[\t ]+\S.*|[0-9]+\.[\t ]+\S.*)', re.MULTILINE
    )

    @staticmethod
    def _extract_subheadings(text: str) -> List[str]:
        """Extract subheading lines from section text.

        Returns full line text for each L1 (A., B.) or L2 (1., 2.) subheading.
        """
        results = []
        for line in text.split('\r'):
            line_s = line.strip()
            if re.match(r'^[A-Z]\.[\t ]', line_s) or re.match(r'^[0-9]+\.[\t ]', line_s):
                results.append(line_s)
        # Also try \n split (LLM output uses \n)
        if not results:
            for line in text.split('\n'):
                line_s = line.strip()
                if re.match(r'^[A-Z]\.[\t ]', line_s) or re.match(r'^[0-9]+\.[\t ]', line_s):
                    results.append(line_s)
        return results

    @staticmethod
    def _check_subheadings_preserved(original: str, revised: str) -> List[str]:
        """Return list of subheading labels found in original but missing from revised."""
        orig_subs = ReportReviewer._extract_subheadings(original)
        if not orig_subs:
            return []

        dropped = []
        for label in orig_subs:
            # Normalize: collapse whitespace, strip tab differences
            label_norm = re.sub(r'[\t ]+', ' ', label).strip()
            found = False
            for rev_line in revised.replace('\r', '\n').split('\n'):
                rev_norm = re.sub(r'[\t ]+', ' ', rev_line).strip()
                if rev_norm == label_norm:
                    found = True
                    break
            if not found:
                dropped.append(label[:60])
        return dropped

    def _restore_subheadings_in_revised(self, original: str, revised: str) -> str:
        """If the LLM mangled subheading lines, restore them from the original.

        Matches by the letter/number prefix (A., B., 1., 2.) and replaces
        the full line in revised with the original version.
        """
        orig_subs = self._extract_subheadings(original)
        if not orig_subs:
            return revised

        # Build map: prefix -> original line
        prefix_map = {}
        for sub in orig_subs:
            m = re.match(r'^([A-Z]\.|[0-9]+\.)', sub)
            if m:
                prefix_map[m.group(1)] = sub

        # Fix revised lines that start with a known prefix but differ
        lines = revised.replace('\r\n', '\n').split('\n')
        fixed = []
        for line in lines:
            line_s = line.strip()
            m = re.match(r'^([A-Z]\.|[0-9]+\.)', line_s)
            if m and m.group(1) in prefix_map:
                orig_line = prefix_map[m.group(1)]
                orig_norm = re.sub(r'[\t ]+', ' ', orig_line).strip()
                rev_norm = re.sub(r'[\t ]+', ' ', line_s).strip()
                if rev_norm != orig_norm:
                    # LLM changed the subheading — restore original
                    logger.info(f"  Restoring subheading: '{line_s[:50]}' -> "
                                f"'{orig_line[:50]}'")
                    fixed.append(orig_line)
                    continue
            fixed.append(line)
        return '\n'.join(fixed)

    @staticmethod
    def _classify_subheading(para) -> Optional[str]:
        """Classify a paragraph as L1, L2, or None.

        Uses Word list formatting when available (most reliable):
        - listType=3 (outline numbered) with uppercase-letter NumberStyle → L1
        - listType=3 with arabic NumberStyle → L2

        Falls back to formatting heuristics for non-list subheadings:
        - Must be bold, short (< 80 chars), not ALL CAPS
        - L1: bold + underlined + indented (left > 20pt)
        - L2: bold + more deeply indented (left > 50pt)
        """
        try:
            text = para.Range.Text.strip().rstrip('\r')
            if not text or len(text) > 80:
                return None
            if text.isupper():
                return None
            if para.Range.Bold != -1 and para.Range.Bold != True:  # noqa: E712
                return None

            # Check Word list formatting first
            try:
                lf = para.Range.ListFormat
                if lf.ListType == 3:  # outline numbered
                    ns = lf.ListTemplate.ListLevels(1).NumberStyle
                    # 3 = wdListNumberStyleUppercaseLetter (A., B.)
                    # 0 = wdListNumberStyleArabic (1., 2., 3.)
                    if ns == 3:
                        return "L1"
                    elif ns == 0:
                        return "L2"
            except Exception:
                pass

            # Fallback: formatting-based detection
            pf = para.Format
            is_underlined = (para.Range.Underline == 1)
            left = pf.LeftIndent

            if re.match(r'^[A-Z]\.[\t ]', text):
                return "L1"
            if re.match(r'^[0-9]+\.[\t ]', text):
                return "L2"
            if is_underlined and left > 20:
                return "L1"
            if not is_underlined and left > 50:
                return "L2"

            return None
        except Exception:
            return None

    def _fix_subheading_formatting(self, section: 'SectionInfo'):
        """After redline, fix subheading formatting via COM.

        Uses pre-detected subheadings from detect_sections() (bulk method).
        Re-detects to get fresh positions after redline shifts.

        L1 (A., B.): bold, paragraph-level underline, list TextPosition → 72pt (1.0")
        L2 (1., 2.): bold, character-level underline on TEXT ONLY (not number/tab gap)
        """
        try:
            # Re-detect sections to get fresh positions after redline
            current_sections = self.detect_sections()
            fresh = None
            for s in current_sections:
                if s.name == section.name:
                    fresh = s
                    break
            if not fresh:
                logger.warning(f"Could not re-locate {section.name} for formatting fix")
                return

            fixed_count = 0
            fixed_templates = set()

            for sub in fresh.subheadings:
                try:
                    para = self.doc.Range(sub.range_start, sub.range_end).Paragraphs(1)
                    text = para.Range.Text.strip().rstrip('\r')

                    # Ensure bold
                    if para.Range.Bold != -1:
                        logger.info(f"  Fixing bold on {sub.level}: '{text[:40]}'")
                        para.Range.Bold = -1

                    # Strip leading whitespace/tabs from list-formatted paragraphs
                    if para.Range.ListFormat.ListType == 3:
                        raw = para.Range.Text.rstrip('\r')
                        leading = 0
                        for ch in raw:
                            if ch in ' \t':
                                leading += 1
                            else:
                                break
                        if leading > 0:
                            strip_range = self.doc.Range(
                                para.Range.Start, para.Range.Start + leading)
                            logger.info(f"  Stripping {leading} leading chars from "
                                        f"{sub.level}: '{text[:40]}'")
                            strip_range.Delete()

                    if sub.level == "L1":
                        if para.Range.Underline != 1:
                            logger.info(f"  Fixing underline on L1: '{text[:40]}'")
                            para.Range.Underline = 1

                        try:
                            lf = para.Range.ListFormat
                            if lf.ListType == 3:
                                lt_id = id(lf.ListTemplate)
                                if lt_id not in fixed_templates:
                                    ll = lf.ListTemplate.ListLevels(1)
                                    if abs(ll.TextPosition - 72) > 2:
                                        logger.info(f"  Fixing L1 list TextPosition: "
                                                    f"{ll.TextPosition:.0f} → 72pt")
                                        ll.TextPosition = 72
                                    fixed_templates.add(lt_id)
                        except Exception as e:
                            logger.debug(f"  Could not fix L1 list template: {e}")

                    elif sub.level == "L2":
                        try:
                            lf = para.Range.ListFormat
                            if lf.ListType == 3:
                                lt_id = id(lf.ListTemplate)
                                if lt_id not in fixed_templates:
                                    ll = lf.ListTemplate.ListLevels(1)
                                    if abs(ll.TextPosition - 108) > 2:
                                        logger.info(f"  Fixing L2 list TextPosition: "
                                                    f"{ll.TextPosition:.0f} → 108pt")
                                        ll.TextPosition = 108
                                    fixed_templates.add(lt_id)
                        except Exception as e:
                            logger.debug(f"  Could not fix L2 list template: {e}")

                        # Character-level underline on TEXT ONLY
                        r = para.Range
                        raw = r.Text.rstrip('\r')

                        if r.Underline == 1:
                            r.Underline = 0

                        text_offset = 0
                        for ch in raw:
                            if ch.isalpha():
                                break
                            text_offset += 1

                        if text_offset < len(raw):
                            text_range = self.doc.Range(
                                r.Start + text_offset, r.End - 1)
                            logger.info(f"  Fixing underline on L2 (text-only): "
                                        f"'{text_range.Text.strip()[:40]}'")
                            text_range.Underline = 1

                    fixed_count += 1

                except Exception as e:
                    logger.debug(f"  Error fixing subheading '{sub.text[:30]}': {e}")

            logger.info(f"  Fixed formatting on {fixed_count} subheadings "
                        f"in {section.name}")

        except Exception as e:
            logger.warning(f"Could not fix subheading formatting: {e}")
            import traceback
            traceback.print_exc()

    def _apply_redline(self, range_start: int, range_end: int,
                       original_text: str, revised_text: str,
                       section_name: str = "") -> int:
        """Apply Track Changes for a section. Returns number of changes applied."""
        try:
            from icharlotte_core.word_hotkey import apply_flat_diff_redline

            def _do_redline():
                return apply_flat_diff_redline(
                    doc_com=self.doc,
                    range_start=range_start,
                    range_end=range_end,
                    original_text=original_text,
                    revised_text=revised_text,
                    auto_enable_track_changes=True,
                )

            r = _do_redline()

            if r['success']:
                time.sleep(0.3)
                logger.info(f"Applied {r['total_changes']} changes to {section_name} "
                            f"({r.get('deleted_chars', 0)} del, "
                            f"{r.get('inserted_chars', 0)} ins)")
                return r['total_changes']

            logger.warning(f"Redline application returned failure for {section_name}")
            return 0

        except Exception as e:
            logger.error(f"Failed to apply redline for {section_name}: {e}")
            return 0

    def _run_final_validation(self, sections: List[SectionInfo],
                              total_changes: int = 0):
        """Run final validation after all sections have been redlined.

        Uses the shared word_validator.py checks:
        - check_paragraph_marks_preserved (across full range)
        - Section heading formatting check (only actual section headings, not all bold text)
        """
        from icharlotte_core.word_validator import (
            Finding, ValidationResult,
            check_paragraph_marks_preserved,
        )

        if not sections:
            logger.warning("No sections to validate")
            return

        # Full range covering all reviewed sections
        full_start = min(s.heading_range_start for s in sections)
        full_end = max(s.range_end for s in sections)

        result = ValidationResult(
            context=f"Final report review validation "
                    f"({len(sections)} sections, {total_changes} changes)"
        )

        # 1. Paragraph marks preserved (no formatting destruction)
        result.findings.extend(
            check_paragraph_marks_preserved(self.doc, full_start, full_end))

        # 2. Check actual section headings are still bold/underline/ALL CAPS
        for section in sections:
            try:
                hr = self.doc.Range(section.heading_range_start,
                                    section.heading_range_end)
                text = hr.Text.strip().rstrip('\r')
                issues = []
                if hr.Bold != -1 and hr.Bold != True:  # noqa: E712
                    issues.append("lost bold")
                if not hr.Underline:
                    issues.append("lost underline")
                if not text.isupper():
                    issues.append("lost ALL CAPS")
                if issues:
                    result.findings.append(Finding(
                        severity="ERROR",
                        rule="section_heading_formatting",
                        message=f"{section.name}: {', '.join(issues)}",
                        location=f"chars {section.heading_range_start}-{section.heading_range_end}",
                    ))
                else:
                    result.findings.append(Finding(
                        severity="PASS",
                        rule="section_heading_formatting",
                        message=f"{section.name} heading intact",
                    ))
            except Exception as e:
                logger.debug(f"Could not check heading {section.name}: {e}")

        result.print_summary()

        if result.has_errors:
            logger.error("VALIDATION ERRORS found — review tracked changes carefully")
        else:
            logger.info(f"Final validation passed: {len(sections)} headings intact, "
                        f"no errors in {total_changes} changes")

    def _calc_preservation(self, original: str, revised: str) -> float:
        """Calculate what fraction of the original text is preserved in the revision.

        Uses diff_match_patch for accuracy. Returns 0.0-1.0.
        """
        try:
            from diff_match_patch import diff_match_patch
            dmp = diff_match_patch()
            diffs = dmp.diff_main(original, revised)
            dmp.diff_cleanupSemantic(diffs)

            equal_chars = sum(len(text) for op, text in diffs if op == 0)
            total_original = len(original)
            if total_original == 0:
                return 1.0
            return equal_chars / total_original
        except Exception:
            # Fallback: simple ratio
            shorter = min(len(original), len(revised))
            longer = max(len(original), len(revised))
            if longer == 0:
                return 1.0
            return shorter / longer

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
            return _com_retry(lambda: self.doc.Range(start, end).Text or "")
        except Exception as e:
            logger.error(f"Could not get range text: {e}")
            return ""
