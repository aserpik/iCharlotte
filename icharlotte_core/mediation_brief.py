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
from typing import Dict, List, Optional

import fitz  # PyMuPDF

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
