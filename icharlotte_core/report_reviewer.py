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

# Sections containing tabular data — skip LLM review, tables get mangled
TABULAR_SECTIONS = {"LITIGATION BUDGET"}

# Minimum preservation ratio — if the LLM rewrites more than this, reject it
MIN_PRESERVATION_RATIO = 0.25


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

                revised = self._review_text(
                    fresh_text, section_name, full_doc_text,
                    case_data, extra_texts
                )

                if not revised or revised.strip() == fresh_text.strip():
                    progress(f"  {section_name}: no changes needed")
                    continue

                # Check preservation ratio before applying
                preservation = self._calc_preservation(fresh_text, revised)
                if preservation < MIN_PRESERVATION_RATIO:
                    progress(f"  {section_name}: LLM rewrote too aggressively "
                             f"({preservation:.0%} preserved, min {MIN_PRESERVATION_RATIO:.0%}), skipping")
                    logger.warning(f"Rejected LLM output for {section_name}: "
                                   f"only {preservation:.0%} preserved")
                    continue

                # Restore any mangled subheading lines from the original
                revised = self._restore_subheadings_in_revised(fresh_text, revised)

                # Check subheadings are preserved (after restoration attempt)
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
                if changes > 0:
                    time.sleep(1.0)  # let Word settle before formatting pass
                    self._fix_subheading_formatting(section)

            except Exception as e:
                logger.error(f"Error reviewing {section_name}: {e}")
                progress(f"  {section_name}: error ({e}), continuing...")

        # Final full-document validation
        progress("Running final validation...")
        final_sections = self.detect_sections()
        self._run_final_validation(final_sections, total_changes)

        progress(f"Review complete — {len(reviewable_names)} sections reviewed, "
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
        paras = _com_retry(lambda: self.doc.Paragraphs)
        para_count = _com_retry(lambda: paras.Count)
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

        prompt = f"""You are reviewing the "{section_name}" section of a litigation report written by a junior associate. Your task is to make TARGETED, SURGICAL improvements to match the senior attorney's voice, tone, and style.

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

WHAT TO IMPROVE (only where clearly needed):
- Shift passive or casual phrasing to defense counsel's formal voice
- Add hedging language where appropriate: {', '.join(hedging)}
- Refer to parties as "Plaintiff" and "Defendant" (not first names)
- Fix grammatical errors, awkward phrasing, or unclear sentences
- Ensure content flows logically

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
Preserve the associate's analysis and structure — only change what genuinely needs improvement.
Output ONLY the revised section body text, nothing else. Do NOT include the section heading."""

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

    def _fix_subheading_formatting(self, section: 'SectionInfo'):
        """After redline, ensure subheading paragraphs have correct bold,
        underline, and indentation via COM.

        L1 (A., B.): bold, underline, left=1440 twips, hanging=720
        L2 (1., 2.): bold, underline, left=2160 twips, hanging=720
        """
        try:
            range_obj = self.doc.Range(section.range_start, section.range_end)
            paras = range_obj.Paragraphs
            for i in range(1, paras.Count + 1):
                para = paras(i)
                text = para.Range.Text.strip().rstrip('\r')
                if not text:
                    continue

                is_l1 = bool(re.match(r'^[A-Z]\.[\t ]', text))
                is_l2 = bool(re.match(r'^[0-9]+\.[\t ]', text))

                if not is_l1 and not is_l2:
                    continue

                # Fix bold
                if para.Range.Bold != -1:  # -1 means all bold
                    logger.info(f"  Fixing bold on subheading: '{text[:40]}'")
                    para.Range.Bold = -1

                # Fix underline (wdUnderlineSingle = 1)
                if para.Range.Underline != 1:
                    logger.info(f"  Fixing underline on subheading: '{text[:40]}'")
                    para.Range.Underline = 1

                # Fix indentation
                pf = para.Format
                if is_l1:
                    # L1: left 1.0" (72pt), hanging 0.5" (36pt)
                    if abs(pf.LeftIndent - 72) > 2:
                        logger.info(f"  Fixing L1 indent on: '{text[:40]}'")
                        pf.LeftIndent = 72      # 1.0" in points
                        pf.FirstLineIndent = -36  # hanging 0.5"
                elif is_l2:
                    # L2: left 1.5" (108pt), hanging 0.5" (36pt)
                    if abs(pf.LeftIndent - 108) > 2:
                        logger.info(f"  Fixing L2 indent on: '{text[:40]}'")
                        pf.LeftIndent = 108     # 1.5" in points
                        pf.FirstLineIndent = -36  # hanging 0.5"

        except Exception as e:
            logger.warning(f"Could not fix subheading formatting: {e}")

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
