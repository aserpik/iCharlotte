"""
MediationBriefGenerator — style extraction from sample briefs and caching.

Reads sample mediation brief PDFs to extract per-section style excerpts.
These excerpts are used as few-shot examples when generating new briefs.
Cache is hash-validated so it auto-refreshes when samples change.
"""

import hashlib
import json
import logging
import os
import re
import shutil
from typing import Dict, List, Optional

from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT

import fitz  # PyMuPDF
from docx import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from icharlotte_core.llm_config import LLMCaller
from icharlotte_core.prompt_manager import get_prompt_manager

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Module-level constants
# ---------------------------------------------------------------------------

SAMPLE_BRIEFS_DIR = r"C:\AI\Mediation Briefs"

# Canonical display order (matches roman numeral order in the document)
SECTION_ORDER = [
    "introduction",
    "statement_of_facts",
    "procedural_status",
    "liability",
    "damages",
    "settlement_position",
    "conclusion",
]

# Generation order: introduction last (written after the other sections are done)
GENERATION_ORDER = [
    "statement_of_facts",
    "procedural_status",
    "liability",
    "damages",
    "settlement_position",
    "conclusion",
    "introduction",
]

# Maps canonical section name → (roman_numeral_str, HEADING_TITLE)
SECTION_HEADINGS: Dict[str, tuple] = {
    "introduction":       ("I",   "INTRODUCTION"),
    "statement_of_facts": ("II",  "STATEMENT OF FACTS"),
    "procedural_status":  ("III", "PROCEDURAL STATUS"),
    "liability":          ("IV",  "LIABILITY"),
    "damages":            ("V",   "DAMAGES"),
    "settlement_position":("VI",  "SETTLEMENT POSITION"),
    "conclusion":         ("VII", "CONCLUSION"),
}

# Regex matching roman-numeral section headings as they appear in sample briefs.
# Examples:
#   "I.     INTRODUCTION"
#   "II.  STATEMENT OF FACTS"
#   "IV.   FACTUAL BACKGROUND"
_HEADING_PATTERN = re.compile(
    r"^\s*(I{1,3}|IV|V?I{0,3}|VI{1,2}|VII)\s*\.\s+"
    r"([A-Z][A-Z \t]+[A-Z])\s*$",
    re.MULTILINE,
)

# Maps variant heading titles found in real briefs → canonical section names
_HEADING_TO_SECTION: Dict[str, str] = {
    # introduction variants
    "INTRODUCTION":              "introduction",
    "OVERVIEW":                  "introduction",
    # statement of facts variants
    "STATEMENT OF FACTS":        "statement_of_facts",
    "FACTUAL BACKGROUND":        "statement_of_facts",
    "FACTS":                     "statement_of_facts",
    "BACKGROUND":                "statement_of_facts",
    # procedural status variants
    "PROCEDURAL STATUS":         "procedural_status",
    "PROCEDURAL BACKGROUND":     "procedural_status",
    "PROCEDURAL HISTORY":        "procedural_status",
    "STATUS OF LITIGATION":      "procedural_status",
    # liability variants
    "LIABILITY":                 "liability",
    "LIABILITY ANALYSIS":        "liability",
    "ANALYSIS OF LIABILITY":     "liability",
    # damages variants
    "DAMAGES":                   "damages",
    "DAMAGES ANALYSIS":          "damages",
    "ANALYSIS OF DAMAGES":       "damages",
    "INJURIES AND DAMAGES":      "damages",
    "SPECIAL DAMAGES":           "damages",
    # settlement position variants
    "SETTLEMENT POSITION":       "settlement_position",
    "SETTLEMENT VALUE":          "settlement_position",
    "VALUATION":                 "settlement_position",
    "DEMAND":                    "settlement_position",
    # conclusion variants
    "CONCLUSION":                "conclusion",
    "SUMMARY":                   "conclusion",
}

# Paths for prompt files and style cache
PROMPTS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "Scripts", "prompts", "mediation_brief",
)
CACHE_PATH = os.path.join(PROMPTS_DIR, "style_cache.json")


# ---------------------------------------------------------------------------
# MediationBriefGenerator
# ---------------------------------------------------------------------------

class MediationBriefGenerator:
    """Generates mediation briefs by extracting style from samples and calling LLM."""

    # ------------------------------------------------------------------
    # Style and formatting constants
    # ------------------------------------------------------------------

    # Sections whose content affects the Introduction — regenerate it when any of these change.
    _INTRO_TRIGGERS = {"liability", "damages", "statement_of_facts", "conclusion"}

    STYLE_GUIDE = """
STYLE AND TONE GUIDE:
- Write as a senior defense litigation attorney addressing a mediator
- Tone: professional, authoritative, persuasive, and firm
- Use active voice and strong declarative statements
- Avoid hedging or weak qualifiers ("perhaps", "might", "it could be argued")
- Present defense arguments as the clear and logical reading of the evidence
- When discussing plaintiff's position, highlight inconsistencies and weaknesses
- Use specific facts, dates, and evidence — avoid vague generalities
- When including deposition quotes, clean them of transcript artifacts (line numbers, dashes, extra characters) but keep them verbatim
- Citation format for deposition quotes: (LastName Depo Trns., at p. PageNum:LineNum.)
- Do not use placeholder text like [TBD] or [INSERT] — write around missing information naturally
- Be thorough and detailed — length is not a concern
"""

    FORMATTING_RULES = """
FORMATTING RULES:
- Do NOT include level-one section headings (roman numerals) — they are added by the system
- For the LIABILITY and DAMAGES sections: mark each subsection with "SUBSECTION: Title Text" on its own line, followed by the content paragraphs. The system will convert these to properly formatted level-two headings.
- For deposition quotes: put the quote on its own paragraph, preceded by a blank line. Follow with the citation on the same paragraph: (LastName Depo Trns., at p. PageNum:LineNum.)
- Write in plain text. Do not use markdown formatting (no **, ##, etc.)
"""

    # Compile-time regexes for section text parsing
    _SUBSECTION_RE = re.compile(r'^SUBSECTION:\s*(.+)$', re.MULTILINE)
    _DEPO_CITE_RE = re.compile(r'\([A-Z][a-z]+ Depo Trns\., at p\. \d+:\d+\.\)')

    def __init__(self):
        self._sample_dir: str = SAMPLE_BRIEFS_DIR
        self._cache_path: str = CACHE_PATH
        self._style_cache: Optional[Dict] = None

        # Per-section generated content (populated during generation)
        self.sections: Dict[str, str] = {}
        self.planning_output: str = ""
        self.document_content: str = ""
        self.caption_template_path: Optional[str] = None
        self.is_active: bool = False

    # ------------------------------------------------------------------
    # Section parsing
    # ------------------------------------------------------------------

    def _extract_sections_from_text(self, text: str) -> Dict[str, str]:
        """Parse a sample brief's text into a dict of section_name → body_text.

        Identifies section boundaries by roman-numeral heading lines, then maps
        each heading to a canonical section name via ``_HEADING_TO_SECTION``.
        Sections whose headings are unrecognised are silently skipped.
        """
        matches = list(_HEADING_PATTERN.finditer(text))
        if not matches:
            return {}

        sections: Dict[str, str] = {}

        for i, match in enumerate(matches):
            heading_title = match.group(2).strip()
            canonical = _HEADING_TO_SECTION.get(heading_title)
            if canonical is None:
                # Try partial match (heading may have extra words)
                for variant, name in _HEADING_TO_SECTION.items():
                    if heading_title.startswith(variant) or variant in heading_title:
                        canonical = name
                        break

            if canonical is None:
                logger.debug("Unrecognised heading: %r — skipping", heading_title)
                continue

            # Body text runs from end of this heading to start of next heading
            body_start = match.end()
            body_end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            body = text[body_start:body_end].strip()

            if body:
                sections[canonical] = body

        return sections

    # ------------------------------------------------------------------
    # Hashing / caching
    # ------------------------------------------------------------------

    def _hash_file(self, path: str) -> str:
        """Return the MD5 hex digest of the file at *path*."""
        h = hashlib.md5()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()

    def _read_sample_pdfs(self) -> Dict:
        """Read all PDFs in the sample directory, extract sections from each.

        Returns a dict with keys:
          - ``"hashes"``:   {filename: md5_hex}
          - ``"sections"``: {section_name: [excerpt1, excerpt2]}  (best 2 per section)
        """
        if not os.path.isdir(self._sample_dir):
            logger.warning("Sample briefs directory not found: %s", self._sample_dir)
            return {"hashes": {}, "sections": {}}

        pdf_files = sorted(
            f for f in os.listdir(self._sample_dir)
            if f.lower().endswith(".pdf")
        )

        if not pdf_files:
            logger.warning("No PDF files found in %s", self._sample_dir)
            return {"hashes": {}, "sections": {}}

        hashes: Dict[str, str] = {}
        # section_name → list of (text, source_file) sorted by length descending
        raw: Dict[str, List[tuple]] = {s: [] for s in SECTION_ORDER}

        for filename in pdf_files:
            full_path = os.path.join(self._sample_dir, filename)
            try:
                hashes[filename] = self._hash_file(full_path)
                doc = fitz.open(full_path)
                try:
                    text = "\n".join(page.get_text() for page in doc)
                finally:
                    doc.close()

                found = self._extract_sections_from_text(text)
                for section, body in found.items():
                    raw[section].append((body, filename))

            except Exception as exc:
                logger.error("Failed to process sample PDF %s: %s", filename, exc)

        # Keep the 2 longest excerpts per section
        sections: Dict[str, List[str]] = {}
        for section, entries in raw.items():
            if not entries:
                continue
            entries.sort(key=lambda x: len(x[0]), reverse=True)
            sections[section] = [text for text, _ in entries[:2]]

        return {"hashes": hashes, "sections": sections}

    def _save_style_cache(self, data: Dict) -> None:
        """Persist style data to the JSON cache file.

        The on-disk format uses ``source_hashes`` (not ``hashes``) so callers
        can distinguish the raw extraction payload from the persisted format.
        """
        cache_dir = os.path.dirname(self._cache_path)
        os.makedirs(cache_dir, exist_ok=True)

        payload = {
            "source_hashes": data.get("hashes", {}),
            "sections": data.get("sections", {}),
        }
        with open(self._cache_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, ensure_ascii=False)
        logger.debug("Style cache saved to %s", self._cache_path)

    def _load_style_cache(self) -> Optional[Dict]:
        """Load and validate the style cache from disk.

        Validates that all PDF files recorded in the cache still exist and their
        MD5 hashes match.  Returns ``None`` if the cache is absent, corrupt, or
        stale.
        """
        if not os.path.isfile(self._cache_path):
            return None

        try:
            with open(self._cache_path, "r", encoding="utf-8") as fh:
                cached = json.load(fh)
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning("Could not read style cache: %s", exc)
            return None

        source_hashes = cached.get("source_hashes", {})
        if not source_hashes:
            # Cache with no file references is considered stale
            return None

        # Validate each hash
        for filename, expected_hash in source_hashes.items():
            full_path = os.path.join(self._sample_dir, filename)
            if not os.path.isfile(full_path):
                logger.info("Cache stale: %s no longer exists", filename)
                return None
            try:
                actual_hash = self._hash_file(full_path)
            except OSError as exc:
                logger.warning("Cannot hash %s: %s", filename, exc)
                return None
            if actual_hash != expected_hash:
                logger.info("Cache stale: %s has changed", filename)
                return None

        return cached

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_style_excerpts(self) -> Dict[str, List[str]]:
        """Return style excerpts keyed by section name.

        On first call the sample PDFs are parsed and the result is cached.
        Subsequent calls return the in-memory cache.  The on-disk JSON cache is
        used across sessions; it is invalidated whenever any source PDF changes.
        """
        if self._style_cache is not None:
            return self._style_cache.get("sections", {})

        # Try loading from disk first
        cached = self._load_style_cache()
        if cached is not None:
            logger.debug("Using style cache from disk")
            self._style_cache = cached
            return cached.get("sections", {})

        # Extract fresh from sample PDFs
        logger.info("Extracting style from sample briefs in %s", self._sample_dir)
        data = self._read_sample_pdfs()
        self._save_style_cache(data)

        # Normalise to the same format as the on-disk cache
        self._style_cache = {
            "source_hashes": data.get("hashes", {}),
            "sections": data.get("sections", {}),
        }
        return self._style_cache.get("sections", {})

    # ------------------------------------------------------------------
    # LLM prompt construction and generation
    # ------------------------------------------------------------------

    def _get_section_prompt(self, section_name: str) -> str:
        """Load the main prompt for *section_name* from the Workbench (PromptManager).

        Falls back to reading the file directly from PROMPTS_DIR if the
        PromptManager returns nothing.  Returns an empty string if neither
        source has a prompt for this section.
        """
        pm = get_prompt_manager()
        prompt = pm.get_prompt("mediation_brief", section_name)
        if prompt:
            return prompt

        # Direct file fallback
        fallback_path = os.path.join(PROMPTS_DIR, f"{section_name}_current.txt")
        if os.path.isfile(fallback_path):
            with open(fallback_path, "r", encoding="utf-8") as fh:
                return fh.read()

        logger.warning("No prompt found for section %r", section_name)
        return ""

    def _build_system_prompt(self, section_name: str) -> str:
        """Return the hard-coded system prompt for the given section.

        Combines the style guide, formatting rules, and a defense attorney
        persona statement.
        """
        persona = (
            "You are a senior defense litigation attorney with decades of trial experience. "
            "You are writing a confidential mediation brief on behalf of the defendant. "
            "Your goal is to present the strongest possible defense position to the mediator, "
            "supported by the evidence and the facts of the case."
        )
        return f"{persona}\n{self.STYLE_GUIDE}\n{self.FORMATTING_RULES}"

    def _build_section_prompt(self, section_name: str, refinement_instruction: str = "") -> str:
        """Build the full prompt for *section_name*.

        Assembles:
        1. Main prompt loaded from the Workbench.
        2. Optional refinement instruction.
        3. Style excerpt from cached samples (first excerpt, truncated at 10 000 chars).
        4. Planning pass output.
        5. Previously generated sections — ALL sections when writing the
           introduction; only sections that precede *section_name* in
           GENERATION_ORDER otherwise.
        """
        parts: List[str] = []

        # 1. Main section prompt
        main_prompt = self._get_section_prompt(section_name)
        if main_prompt:
            parts.append(main_prompt)

        # 2. Refinement instruction
        if refinement_instruction:
            parts.append(f"\nREFINEMENT INSTRUCTION:\n{refinement_instruction}")

        # 3. Style excerpt
        excerpts = self.get_style_excerpts()
        section_excerpts = excerpts.get(section_name, [])
        if section_excerpts:
            excerpt = section_excerpts[0][:10000]
            parts.append(
                f"\nSTYLE REFERENCE — example from a prior mediation brief "
                f"(use only as a stylistic guide):\n{excerpt}"
            )

        # 4. Planning pass output
        if self.planning_output:
            parts.append(f"\nPLANNING ANALYSIS:\n{self.planning_output}")

        # 5. Previously generated sections
        if section_name == "introduction":
            # Introduction is written last — include every other section
            prior_sections = {
                s: self.sections[s]
                for s in SECTION_ORDER
                if s != "introduction" and s in self.sections
            }
        else:
            # Include all sections that were generated before this one
            current_index = GENERATION_ORDER.index(section_name) if section_name in GENERATION_ORDER else -1
            prior_names = GENERATION_ORDER[:current_index] if current_index > 0 else []
            prior_sections = {s: self.sections[s] for s in prior_names if s in self.sections}

        if prior_sections:
            prior_text = "\n\n".join(
                f"[{s.upper().replace('_', ' ')}]\n{body}"
                for s, body in prior_sections.items()
            )
            parts.append(f"\nPREVIOUSLY DRAFTED SECTIONS (for context and consistency):\n{prior_text}")

        return "\n".join(parts)

    def generate_section(self, section_name: str, refinement_instruction: str = "") -> Optional[str]:
        """Generate content for *section_name* using the LLM.

        Calls LLMCaller with the section-specific system prompt and the
        assembled section prompt, using the document content as the text
        body.  Stores the result in ``self.sections[section_name]`` and
        returns it.  Returns None if the LLM call fails.
        """
        system_prompt = self._build_system_prompt(section_name)
        section_prompt = self._build_section_prompt(section_name, refinement_instruction)

        caller = LLMCaller()
        # Pass system prompt prepended to the section prompt so the caller
        # receives both in the user-visible prompt parameter.
        full_prompt = f"{system_prompt}\n\n{section_prompt}"

        result = caller.call(
            prompt=full_prompt,
            text=self.document_content,
            agent_id="agent_mediation_brief",
        )

        if result:
            self.sections[section_name] = result
            logger.info("Generated section %r (%d chars)", section_name, len(result))
        else:
            logger.error("LLM returned no content for section %r", section_name)

        return result

    def run_planning_pass(self) -> Optional[str]:
        """Run the planning pass over the document content.

        Loads the planning prompt, calls the LLM, stores the output in
        ``self.planning_output``, and returns it.  Returns None if the
        LLM call fails.
        """
        planning_prompt = self._get_section_prompt("planning")
        if not planning_prompt:
            logger.warning("Planning prompt not found — skipping planning pass")
            return None

        caller = LLMCaller()
        result = caller.call(
            prompt=planning_prompt,
            text=self.document_content,
            agent_id="agent_mediation_brief",
        )

        if result:
            self.planning_output = result
            logger.info("Planning pass complete (%d chars)", len(result))
        else:
            logger.error("Planning pass: LLM returned no content")

        return result

    # ------------------------------------------------------------------
    # Caption template handling
    # ------------------------------------------------------------------

    # XML namespace constants for Word document XML
    _W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    _W_T = f"{{{_W_NS}}}t"
    _W_P = f"{{{_W_NS}}}p"
    _W_R = f"{{{_W_NS}}}r"

    # Signature block indicator phrases (case-insensitive search)
    _SIGNATURE_INDICATORS = [
        "By:",
        "DATED:",
        "Respectfully submitted",
        "State Bar No.",
        "Attorney for",
        "Counsel for",
    ]

    def find_caption_template(self, folder: str) -> Optional[str]:
        """Search *folder* for a .docx file with "caption" in its name.

        Returns the full path to the first match (case-insensitive), or None
        if no matching file is found.
        """
        if not os.path.isdir(folder):
            return None
        for entry in os.listdir(folder):
            if entry.lower().endswith(".docx") and "caption" in entry.lower():
                return os.path.join(folder, entry)
        return None

    def prepare_caption_template(self, caption_path: str, output_path: str) -> list:
        """Copy *caption_path* to *output_path*, process it, and return extracted
        signature paragraph elements.

        Processing steps:
        1. Copy original file to output_path (never modify the original).
        2. Replace "CAPTION PAGE" text in body and footers with the styled
           "DEFENDANT'S CONFIDENTIAL MEDIATION BRIEF" title.
        3. Extract (and remove) any trailing signature block paragraphs.
        4. Save the modified document.

        Returns a list of extracted signature paragraph elements (may be empty).
        """
        shutil.copy2(caption_path, output_path)
        doc = DocxDocument(output_path)

        self._replace_caption_page_body(doc)
        self._replace_caption_page_footers(doc)
        sig_elements = self._extract_signature_block(doc)

        doc.save(output_path)
        return sig_elements

    def _replace_caption_page_body(self, doc) -> None:
        """Search ALL w:t elements in the document body XML for "CAPTION PAGE"
        (including text inside nested tables) and replace the containing paragraph
        with three styled runs:
          - "DEFENDANT'S " (bold)
          - "CONFIDENTIAL" (bold + underline)
          - " MEDIATION BRIEF" (bold)

        Uses raw XML iteration because caption templates typically use nested
        table layouts where doc.paragraphs won't find the text.
        """
        W_T = self._W_T
        W_P = self._W_P
        W_R = self._W_R

        for t_elem in doc.element.body.iter(W_T):
            if t_elem.text and "CAPTION PAGE" in t_elem.text.upper():
                # Walk up to the enclosing w:p
                para_elem = t_elem.getparent()
                while para_elem is not None and para_elem.tag != W_P:
                    para_elem = para_elem.getparent()
                if para_elem is None:
                    continue

                # Remove all existing runs from this paragraph
                for run_elem in list(para_elem.iter(W_R)):
                    parent = run_elem.getparent()
                    if parent is not None:
                        parent.remove(run_elem)

                # Helper to build a bold run element
                def _bold_run(text: str, underline: bool = False) -> OxmlElement:
                    run = OxmlElement("w:r")
                    rPr = OxmlElement("w:rPr")
                    rPr.append(OxmlElement("w:b"))
                    if underline:
                        u_elem = OxmlElement("w:u")
                        u_elem.set(qn("w:val"), "single")
                        rPr.append(u_elem)
                    run.append(rPr)
                    t = OxmlElement("w:t")
                    t.text = text
                    t.set(qn("xml:space"), "preserve")
                    run.append(t)
                    return run

                para_elem.append(_bold_run("DEFENDANT'S "))
                para_elem.append(_bold_run("CONFIDENTIAL", underline=True))
                para_elem.append(_bold_run(" MEDIATION BRIEF"))
                return  # Only replace the first occurrence

    def _replace_caption_page_footers(self, doc) -> None:
        """Iterate all section footers and replace "CAPTION PAGE" paragraphs
        with the same three-run styled title as in the body replacement.
        """
        for section in doc.sections:
            for footer in (
                section.footer,
                section.even_page_footer,
                section.first_page_footer,
            ):
                if footer is None:
                    continue
                for para in footer.paragraphs:
                    if "CAPTION PAGE" in para.text.upper():
                        # Clear existing runs
                        for run in list(para.runs):
                            run._element.getparent().remove(run._element)
                        # Add styled runs
                        r1 = para.add_run("DEFENDANT'S ")
                        r1.bold = True
                        r2 = para.add_run("CONFIDENTIAL")
                        r2.bold = True
                        r2.underline = True
                        r3 = para.add_run(" MEDIATION BRIEF")
                        r3.bold = True

    def _extract_signature_block(self, doc) -> list:
        """Scan backwards through the last 15 paragraphs looking for signature
        block indicators.  When found, extract all paragraphs from that point
        to the end of the document, remove them from the body, and return the
        extracted paragraph objects.

        Returns a list of Paragraph objects (may be empty if no signature block
        is detected).  The elements are removed from the document body so they
        can be re-inserted elsewhere in the assembled brief.
        """
        indicators = self._SIGNATURE_INDICATORS
        all_paras = doc.paragraphs

        # Search the last 15 paragraphs for a signature indicator
        search_paras = all_paras[-15:] if len(all_paras) > 15 else all_paras
        sig_start_index = None

        for i, para in enumerate(search_paras):
            text = para.text
            for indicator in indicators:
                if indicator.lower() in text.lower():
                    # Map back to index in the full paragraph list
                    offset = len(all_paras) - len(search_paras)
                    sig_start_index = offset + i
                    break
            if sig_start_index is not None:
                break

        if sig_start_index is None:
            return []

        # Collect paragraphs from sig_start_index to end
        sig_paras = list(all_paras[sig_start_index:])

        # Remove them from the document body in reverse order to preserve indices
        for para in reversed(sig_paras):
            para_elem = para._element
            parent = para_elem.getparent()
            if parent is not None:
                parent.remove(para_elem)

        return sig_paras

    # ------------------------------------------------------------------
    # Section text parsing
    # ------------------------------------------------------------------

    def _parse_section_text(self, text: str, section_name: str) -> List[Dict]:
        """Parse LLM output into structured elements.

        Splits *text* on double-newlines into paragraphs, then classifies each
        paragraph as one of:
          - ``"l2_heading"``  — line matching ``^SUBSECTION: <title>``
          - ``"depo_quote"``  — paragraph containing a deposition citation
          - ``"body"``        — everything else

        When a SUBSECTION line also has additional content on the same line
        (after the heading text), a separate body element is emitted for it.

        Returns a list of dicts with keys ``"type"`` and ``"text"``.
        """
        elements: List[Dict] = []
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]

        for para in paragraphs:
            subsection_match = self._SUBSECTION_RE.match(para)
            if subsection_match:
                heading_text = subsection_match.group(1).strip()
                # Check if SUBSECTION: line has inline content after it
                # (i.e., there is more text on the same line beyond the title)
                # The regex captures everything after "SUBSECTION: " on that line.
                # If the paragraph has additional lines after the SUBSECTION line,
                # emit them as a body element.
                elements.append({"type": "l2_heading", "text": heading_text})
                # Remainder: lines in the same paragraph after the SUBSECTION line
                remainder_lines = para[subsection_match.end():].strip()
                if remainder_lines:
                    elements.append({"type": "body", "text": remainder_lines})
            elif self._DEPO_CITE_RE.search(para):
                elements.append({"type": "depo_quote", "text": para})
            else:
                elements.append({"type": "body", "text": para})

        return elements

    # ------------------------------------------------------------------
    # Word document formatting helpers
    # ------------------------------------------------------------------

    def _add_tab_stop(self, para, position_inches: float = 0.5) -> None:
        """Add a left-aligned tab stop at *position_inches* to *para*."""
        pPr = para._element.get_or_add_pPr()
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "left")
        tab.set(qn("w:pos"), str(int(position_inches * 1440)))  # twips
        tabs.append(tab)
        pPr.append(tabs)

    def _add_l1_heading(self, doc, roman: str, title: str):
        """Add a level-1 section heading formatted as ``I.     INTRODUCTION``.

        Uses a hanging indent (left=0.5", hanging=0.5") with a tab stop at 0.5"
        so the title text aligns neatly after the roman numeral.
        """
        para = doc.add_paragraph()
        pf = para.paragraph_format
        pf.left_indent = Inches(0.5)
        pf.first_line_indent = Inches(-0.5)
        pf.space_before = Pt(12)
        pf.space_after = Pt(6)
        self._add_tab_stop(para, position_inches=0.5)

        # Run 1: roman numeral — bold only
        r1 = para.add_run(f"{roman}.")
        r1.bold = True

        # Tab character
        para.add_run("\t")

        # Run 2: title — bold + underline
        r2 = para.add_run(title)
        r2.bold = True
        r2.underline = True

        return para

    def _add_l2_heading(self, doc, letter: str, title: str):
        """Add a level-2 subsection heading formatted as ``A.     Title Text``.

        Same hanging indent pattern as L1 headings.
        """
        para = doc.add_paragraph()
        pf = para.paragraph_format
        pf.left_indent = Inches(0.5)
        pf.first_line_indent = Inches(-0.5)
        pf.space_before = Pt(10)
        pf.space_after = Pt(4)
        self._add_tab_stop(para, position_inches=0.5)

        # Run 1: letter — bold only
        r1 = para.add_run(f"{letter}.")
        r1.bold = True

        # Tab character
        para.add_run("\t")

        # Run 2: title — bold + underline
        r2 = para.add_run(title)
        r2.bold = True
        r2.underline = True

        return para

    def _add_body_paragraph(self, doc, text: str):
        """Add a normal body paragraph with space_after=6pt."""
        para = doc.add_paragraph(text)
        para.paragraph_format.space_after = Pt(6)
        return para

    def _add_depo_quote(self, doc, text: str):
        """Add a deposition quote paragraph with left indent and spacing."""
        para = doc.add_paragraph(text)
        pf = para.paragraph_format
        pf.left_indent = Inches(0.5)
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        return para

    # ------------------------------------------------------------------
    # Document assembly
    # ------------------------------------------------------------------

    def assemble_document(
        self,
        caption_path: str,
        output_path: str,
        signature_paragraphs: Optional[list] = None,
    ) -> None:
        """Assemble the full mediation brief Word document.

        Steps:
        1. Call ``prepare_caption_template()`` to copy and process the caption.
        2. Open the processed document.
        3. Add a page break after the caption content.
        4. For each section in SECTION_ORDER, add L1 heading and body elements.
        5. Append the signature block if present.
        6. Save and validate.

        Args:
            caption_path: Path to the source caption .docx file.
            output_path: Destination path for the assembled brief.
            signature_paragraphs: Pre-extracted signature paragraph elements.
                If None, they are extracted from the caption template automatically.
        """
        # Step 1: Prepare caption (copies to output_path, extracts signature)
        extracted_sig = self.prepare_caption_template(caption_path, output_path)
        if signature_paragraphs is None:
            signature_paragraphs = extracted_sig

        # Step 2: Open the processed document
        doc = DocxDocument(output_path)

        # Step 3: Page break after caption
        page_break_para = doc.add_paragraph()
        run = page_break_para.add_run()
        run.add_break(
            __import__("docx.enum.text", fromlist=["WD_BREAK"]).WD_BREAK.PAGE
        )

        # Step 4: Add each section
        letter_seq = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        for section_name in SECTION_ORDER:
            roman, heading_title = SECTION_HEADINGS[section_name]
            self._add_l1_heading(doc, roman, heading_title)

            section_text = self.sections.get(section_name, "")
            if not section_text:
                continue

            elements = self._parse_section_text(section_text, section_name)
            letter_counter = 0  # resets per section

            for elem in elements:
                etype = elem["type"]
                etext = elem["text"]

                if etype == "l2_heading":
                    letter = letter_seq[letter_counter % len(letter_seq)]
                    self._add_l2_heading(doc, letter, etext)
                    letter_counter += 1
                elif etype == "depo_quote":
                    self._add_depo_quote(doc, etext)
                else:
                    self._add_body_paragraph(doc, etext)

        # Step 5: Append signature block
        if signature_paragraphs:
            for sig_para in signature_paragraphs:
                doc.element.body.append(sig_para._element)

        # Step 6: Save
        doc.save(output_path)
        logger.info("Assembled mediation brief saved to %s", output_path)

        # Validate (best-effort)
        try:
            from icharlotte_core.word_validator import validate_report
            result = validate_report(output_path)
            result.print_summary()
        except Exception as e:
            logger.warning("Validation skipped: %s", e)

    # ------------------------------------------------------------------
    # Pipeline orchestration
    # ------------------------------------------------------------------

    def generate_all_sections(self, progress_callback=None) -> None:
        """Run the full generation pipeline: planning pass then all sections.

        Calls ``run_planning_pass()`` first (step 0), then iterates through
        ``GENERATION_ORDER`` generating each section in sequence.

        Args:
            progress_callback: Optional callable ``(section_name, index, total)``
                called before each step (including the planning pass at index 0).
        """
        total = len(GENERATION_ORDER) + 1  # +1 for planning pass

        # Step 0: planning pass
        if progress_callback is not None:
            progress_callback("planning", 0, total)
        self.run_planning_pass()

        # Steps 1..N: generate each section
        for i, section_name in enumerate(GENERATION_ORDER, start=1):
            if progress_callback is not None:
                progress_callback(section_name, i, total)
            result = self.generate_section(section_name)
            if result:
                self.sections[section_name] = result

        self.is_active = True

    # ------------------------------------------------------------------
    # Conversational refinement
    # ------------------------------------------------------------------

    def _parse_routing_response(self, response: str) -> List[str]:
        """Parse the LLM routing response into a list of valid section names.

        Args:
            response: Raw string returned by the routing LLM call.

        Returns:
            A list of canonical section names (those present in SECTION_ORDER).
            Returns an empty list when the response is "none" or contains no
            recognisable section names.
        """
        cleaned = response.strip().lower()
        if cleaned == "none":
            return []
        parts = [p.strip() for p in cleaned.split(",")]
        return [p for p in parts if p in SECTION_ORDER]

    def route_refinement(self, user_message: str) -> List[str]:
        """Ask the LLM which sections of the brief need to be updated.

        Loads the routing prompt, asks the LLM to identify which sections are
        affected by *user_message*, and returns the parsed list.  An empty list
        means the message is not a brief refinement request.

        Args:
            user_message: The user's chat message.

        Returns:
            List of canonical section names to regenerate (may be empty).
        """
        routing_prompt = self._get_section_prompt("routing")
        full_prompt = f"{routing_prompt}\n\nUser message: {user_message}"

        caller = LLMCaller()
        result = caller.call(
            prompt=full_prompt,
            text="",
            agent_id="agent_mediation_brief",
        )

        if not result:
            logger.warning("route_refinement: LLM returned no response")
            return []

        return self._parse_routing_response(result)

    def refine_sections(
        self,
        section_names: List[str],
        instruction: str,
        progress_callback=None,
    ) -> List[str]:
        """Regenerate the specified sections with a refinement instruction.

        When any of the regenerated sections is an _INTRO_TRIGGERS member,
        the Introduction is also regenerated afterwards (without the user's
        instruction — it simply rebuilds itself based on the updated sections).

        Args:
            section_names: Canonical names of the sections to regenerate.
            instruction:   Refinement instruction applied to the requested sections.
            progress_callback: Optional callable ``(section_name,)`` called after
                each section is regenerated.

        Returns:
            List of all section names that were regenerated (including
            "introduction" if it was triggered automatically).
        """
        # Build the regeneration list, adding introduction at the end if needed
        to_regenerate = list(section_names)
        if any(s in self._INTRO_TRIGGERS for s in section_names):
            if "introduction" not in to_regenerate:
                to_regenerate.append("introduction")
            else:
                # Ensure introduction is always last
                to_regenerate.remove("introduction")
                to_regenerate.append("introduction")

        regenerated: List[str] = []
        for section_name in to_regenerate:
            # Apply instruction only to the sections the user asked about;
            # introduction regenerates purely from context, no user instruction.
            if section_name == "introduction" and section_name not in section_names:
                section_instruction = ""
            else:
                section_instruction = instruction

            result = self.generate_section(section_name, refinement_instruction=section_instruction)
            if result:
                self.sections[section_name] = result
                regenerated.append(section_name)

            if progress_callback is not None:
                progress_callback(section_name)

        return regenerated

    def reset(self) -> None:
        """Clear all generated state, resetting the generator to its initial condition."""
        self.sections = {}
        self.planning_output = ""
        self.document_content = ""
        self.caption_template_path = None
        self.is_active = False


# ---------------------------------------------------------------------------
# Background worker
# ---------------------------------------------------------------------------

from PySide6.QtCore import QThread, Signal  # noqa: E402 (import after class definition)


class MediationBriefWorker(QThread):
    """QThread worker that runs the full mediation brief generation pipeline.

    Emits granular signals so a UI can display live progress.

    Signals:
        section_started(str, int, int): section_name, index, total — emitted
            before each step (including planning pass at index 0).
        section_complete(str, str): section_name, section_text — emitted after
            each section (not emitted for the planning pass).
        all_complete(dict): mapping of section_name → text for all sections,
            emitted when generation finishes successfully.
        error(str): error message emitted if an unhandled exception occurs.
    """

    section_started = Signal(str, int, int)
    section_complete = Signal(str, str)
    all_complete = Signal(dict)
    error = Signal(str)

    def __init__(self, generator: MediationBriefGenerator, parent=None):
        super().__init__(parent)
        self._generator = generator
        self._stop_requested: bool = False

    def request_stop(self) -> None:
        """Request that the worker stop after the current step."""
        self._stop_requested = True

    def run(self) -> None:
        """Execute the generation pipeline, emitting signals at each step."""
        try:
            total = len(GENERATION_ORDER) + 1

            # Step 0: planning pass
            self.section_started.emit("planning", 0, total)
            self._generator.run_planning_pass()

            # Steps 1..N: generate each section
            for i, section_name in enumerate(GENERATION_ORDER, start=1):
                if self._stop_requested:
                    break
                self.section_started.emit(section_name, i, total)
                result = self._generator.generate_section(section_name)
                if result:
                    self._generator.sections[section_name] = result
                    self.section_complete.emit(section_name, result)

            self._generator.is_active = True
            self.all_complete.emit(dict(self._generator.sections))

        except Exception as exc:
            logger.exception("MediationBriefWorker: unhandled error")
            self.error.emit(str(exc))
