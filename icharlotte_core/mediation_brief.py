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

import fitz  # PyMuPDF
from docx import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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
