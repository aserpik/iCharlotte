"""
Set tracker for the discovery generation pipeline.

Scans a case's propounded folder to determine next set numbers, last request
numbers, and resolve previous set definitions/instructions.
"""
import logging
import os
import re
from typing import Dict, List, Optional

from .models import DiscoveryType, Party, SetTrackerResult
from .templates import TemplateLoader

logger = logging.getLogger(__name__)

# Pattern for filenames like "SI (1) tPLF.pdf", "RPD(2) tCity.docx"
_FILENAME_PATTERN = re.compile(
    r"^(SI|RPD|RFA)\s*\((\d+)\)\s*t(\w+)\.\w+$",
    re.IGNORECASE,
)

# Pattern for "fOUR Client" folder name (case-insensitive, optional space)
_FOUR_CLIENT_PATTERN = re.compile(r"f\s*our\s*client", re.IGNORECASE)


class SetTracker:
    """Scan a case folder to track existing discovery sets.

    Parameters
    ----------
    case_path : str
        Root path of the case folder.
    """

    def __init__(self, case_path: str):
        self.case_path = case_path
        self._propounded_folder: Optional[str] = None
        self._propounded_resolved = False

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    @staticmethod
    def parse_discovery_filename(filename: str) -> Optional[Dict]:
        """Parse a discovery filename into its components.

        Parameters
        ----------
        filename : str
            Filename like ``"SI (1) tPLF.pdf"`` or ``"RPD(2) tCity.pdf"``.

        Returns
        -------
        dict or None
            ``{"type": "SI", "set_number": 1, "party_abbrev": "PLF"}``
            or ``None`` for non-matching filenames.
        """
        m = _FILENAME_PATTERN.match(filename)
        if not m:
            return None
        return {
            "type": m.group(1).upper(),
            "set_number": int(m.group(2)),
            "party_abbrev": m.group(3),
        }

    def get_next_set_number(self, disc_type: DiscoveryType, party: Party) -> int:
        """Return the next set number for the given type and party.

        Scans existing filenames in the propounded folder.  Returns ``max + 1``
        of all found set numbers, or ``1`` if none exist.
        """
        existing = self._scan_existing_sets(disc_type, party)
        if not existing:
            return 1
        return max(existing) + 1

    def get_last_request_number(self, disc_type: DiscoveryType, party: Party) -> int:
        """Find the highest-numbered request in the latest set file.

        Returns 0 if no file can be read or no request headers are found.
        """
        latest = self._find_latest_set_file(disc_type, party)
        if not latest:
            return 0
        try:
            text = self._read_file_text(latest)
        except Exception:
            logger.warning("Could not read file for request count: %s", latest)
            return 0

        numbers = [
            int(m.group(1))
            for m in re.finditer(disc_type.request_header_pattern, text, re.IGNORECASE)
        ]
        return max(numbers) if numbers else 0

    def resolve_previous_set(
        self,
        disc_type: DiscoveryType,
        party: Party,
        discovery_dir: str,
    ) -> SetTrackerResult:
        """Full resolution of previous set data with cascading fallback.

        Attempts to find definitions and instructions from previous discovery
        documents.  Falls back to standard definitions via TemplateLoader if
        no prior documents are found.

        Parameters
        ----------
        disc_type : DiscoveryType
            The type of discovery (SI, RPD, RFA).
        party : Party
            The party the discovery is directed to.
        discovery_dir : str
            Path to the discovery templates directory (for fallback).

        Returns
        -------
        SetTrackerResult
        """
        next_set = self.get_next_set_number(disc_type, party)
        last_req = self.get_last_request_number(disc_type, party)

        # --- Try to find a readable file ---
        source_file = self._find_readable_set_file(disc_type, party)

        if source_file:
            try:
                text = self._read_file_text(source_file)
                definitions = self._extract_definitions(text)
                instructions = self._extract_instructions(text, disc_type)
                ext = os.path.splitext(source_file)[1].lower()
                method = f"docx_extraction" if ext == ".docx" else "pdf_extraction"
                return SetTrackerResult(
                    next_set_number=next_set,
                    last_request_number=last_req,
                    previous_definitions=definitions,
                    previous_instructions=instructions,
                    resolution_method=method,
                    source_file=source_file,
                )
            except Exception:
                logger.warning(
                    "Could not extract from %s, falling back", source_file
                )

        # --- Fallback to standard definitions ---
        definitions = ""
        instructions = ""
        try:
            loader = TemplateLoader(discovery_dir)
            definitions = loader.load_defined_terms()
        except Exception:
            logger.debug("Could not load standard defined terms")
        # Instructions require a template path; skip if unavailable

        return SetTrackerResult(
            next_set_number=next_set,
            last_request_number=last_req,
            previous_definitions=definitions,
            previous_instructions=instructions,
            resolution_method="standard_definitions",
            source_file=None,
        )

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _find_propounded_folder(self) -> Optional[str]:
        """Find the DISCOVERY/PROPOUNDED/fOUR Client folder.

        Uses case-insensitive matching for each path component.  The client
        folder name is matched with ``r"f\\s*our\\s*client"``.

        Returns
        -------
        str or None
            Absolute path to the propounded folder, or None if not found.
        """
        if self._propounded_resolved:
            return self._propounded_folder

        self._propounded_resolved = True
        self._propounded_folder = None

        # Walk: case_path -> DISCOVERY -> PROPOUNDED -> fOUR Client
        # Each component is matched case-insensitively.
        discovery_dir = self._find_child_ci(self.case_path, "discovery")
        if not discovery_dir:
            return None

        propounded_dir = self._find_child_ci(discovery_dir, "propounded")
        if not propounded_dir:
            return None

        # The client folder is matched with a regex pattern
        try:
            for entry in os.listdir(propounded_dir):
                full = os.path.join(propounded_dir, entry)
                if os.path.isdir(full) and _FOUR_CLIENT_PATTERN.match(entry):
                    self._propounded_folder = full
                    return self._propounded_folder
        except OSError:
            pass

        return None

    @staticmethod
    def _find_child_ci(parent: str, target_name: str) -> Optional[str]:
        """Find a child directory by name, case-insensitively."""
        target_lower = target_name.lower()
        try:
            for entry in os.listdir(parent):
                if entry.lower() == target_lower and os.path.isdir(
                    os.path.join(parent, entry)
                ):
                    return os.path.join(parent, entry)
        except OSError:
            pass
        return None

    def _get_search_dirs(self, party: Party) -> List[str]:
        """Get directories to search for discovery files.

        Returns the base propounded folder plus any party-specific subfolders
        (prefixed with 't', e.g. ``tPltf/``).
        """
        propounded = self._find_propounded_folder()
        if not propounded:
            return []

        dirs = [propounded]

        # Look for party subfolders (tPltf, tDef, etc.)
        try:
            for entry in os.listdir(propounded):
                full = os.path.join(propounded, entry)
                if os.path.isdir(full) and entry.lower().startswith("t"):
                    dirs.append(full)
        except OSError:
            pass

        return dirs

    def _scan_existing_sets(self, disc_type: DiscoveryType, party: Party) -> List[int]:
        """Get list of existing set numbers for the given type and party."""
        sets: List[int] = []
        for search_dir in self._get_search_dirs(party):
            try:
                for fname in os.listdir(search_dir):
                    parsed = self.parse_discovery_filename(fname)
                    if parsed is None:
                        continue
                    if parsed["type"] != disc_type.abbreviation:
                        continue
                    if parsed["party_abbrev"].lower() != party.abbreviation.lower():
                        continue
                    sets.append(parsed["set_number"])
            except OSError:
                continue
        return sets

    def _find_latest_set_file(
        self, disc_type: DiscoveryType, party: Party
    ) -> Optional[str]:
        """Find the file for the latest set, preferring .docx over .pdf."""
        existing = self._scan_existing_sets(disc_type, party)
        if not existing:
            return None

        latest_set = max(existing)

        # Search for .docx first, then .pdf
        for ext in (".docx", ".pdf"):
            result = self._find_set_file_by_set_num(disc_type, party, latest_set, ext)
            if result:
                return result

        return None

    def _find_set_file_by_set_num(
        self,
        disc_type: DiscoveryType,
        party: Party,
        set_num: int,
        ext: str,
    ) -> Optional[str]:
        """Find a specific set file across all search directories."""
        for search_dir in self._get_search_dirs(party):
            result = self._find_set_file_in_dir(search_dir, disc_type, party, set_num, ext)
            if result:
                return result
        return None

    @staticmethod
    def _find_set_file_in_dir(
        directory: str,
        disc_type: DiscoveryType,
        party: Party,
        set_num: int,
        ext: str,
    ) -> Optional[str]:
        """Find a specific set file in a directory.

        Parameters
        ----------
        directory : str
            Directory to search in.
        disc_type : DiscoveryType
            Type of discovery.
        party : Party
            Target party.
        set_num : int
            Set number to find.
        ext : str
            File extension (e.g. ``".docx"``, ``".pdf"``).

        Returns
        -------
        str or None
            Full path to the matching file, or None.
        """
        try:
            for fname in os.listdir(directory):
                if not fname.lower().endswith(ext.lower()):
                    continue
                parsed = SetTracker.parse_discovery_filename(fname)
                if parsed is None:
                    continue
                if (
                    parsed["type"] == disc_type.abbreviation
                    and parsed["set_number"] == set_num
                    and parsed["party_abbrev"].lower() == party.abbreviation.lower()
                ):
                    return os.path.join(directory, fname)
        except OSError:
            pass
        return None

    def _find_readable_set_file(
        self, disc_type: DiscoveryType, party: Party
    ) -> Optional[str]:
        """Find a readable file for the latest set.

        Cascading search:
        1. Check NOTES/AI OUTPUT/DISCOVERY REQUESTS/ for .docx
        2. Check propounded folder for .docx or .pdf
        3. Broader search in case folder

        Returns the first file found, preferring .docx.
        """
        # 1. Check NOTES/AI OUTPUT/DISCOVERY REQUESTS/
        notes_dir = self._find_notes_discovery_dir()
        if notes_dir:
            existing = self._scan_existing_sets(disc_type, party)
            if existing:
                latest_set = max(existing)
                for ext in (".docx", ".pdf"):
                    result = self._find_set_file_in_dir(
                        notes_dir, disc_type, party, latest_set, ext
                    )
                    if result:
                        return result

        # 2. Check propounded folder (.docx first, then .pdf)
        latest = self._find_latest_set_file(disc_type, party)
        if latest:
            return latest

        return None

    def _find_notes_discovery_dir(self) -> Optional[str]:
        """Find the NOTES/AI OUTPUT/DISCOVERY REQUESTS/ directory."""
        notes = self._find_child_ci(self.case_path, "notes")
        if not notes:
            return None
        ai_output = self._find_child_ci(notes, "ai output")
        if not ai_output:
            return None
        disc_req = self._find_child_ci(ai_output, "discovery requests")
        if not disc_req:
            return None
        return disc_req

    @staticmethod
    def _read_file_text(filepath: str) -> str:
        """Read text content from a .docx or .pdf file.

        Parameters
        ----------
        filepath : str
            Path to the file.

        Returns
        -------
        str
            Extracted text content.

        Raises
        ------
        ValueError
            If the file extension is not supported.
        FileNotFoundError
            If the file does not exist.
        """
        if not os.path.isfile(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")

        ext = os.path.splitext(filepath)[1].lower()

        if ext == ".docx":
            from docx import Document as DocxDocument
            doc = DocxDocument(filepath)
            return "\n".join(p.text for p in doc.paragraphs)

        if ext == ".pdf":
            try:
                import fitz  # PyMuPDF
                doc = fitz.open(filepath)
                text_parts = []
                for page in doc:
                    text_parts.append(page.get_text())
                doc.close()
                return "\n".join(text_parts)
            except ImportError:
                raise ValueError(
                    "PyMuPDF (fitz) is required to read PDF files but is not installed"
                )

        raise ValueError(f"Unsupported file extension: {ext}")

    @staticmethod
    def _extract_definitions(text: str) -> str:
        """Extract the definitions section from document text.

        Looks for text between a DEFINITIONS heading and the next major section
        (INSTRUCTIONS, request headers, or section heading).

        Returns empty string if no definitions section is found.
        """
        lines = text.splitlines()
        start_re = re.compile(r"^\s*DEFINITIONS?\s*$", re.IGNORECASE)
        stop_res = [
            re.compile(r"INSTRUCTIONS\s+TO\s+ANSWERING\s+PARTY", re.IGNORECASE),
            re.compile(r"SPECIAL\s+INTERROGATOR", re.IGNORECASE),
            re.compile(r"REQUEST\s+FOR\s+PRODUCTION", re.IGNORECASE),
            re.compile(r"REQUEST\s+FOR\s+ADMISSION", re.IGNORECASE),
            re.compile(r"REQUESTS?\s+FOR\s+PRODUCTION", re.IGNORECASE),
            re.compile(r"REQUESTS?\s+FOR\s+ADMISSION", re.IGNORECASE),
        ]

        collecting = False
        section_lines: List[str] = []

        for line in lines:
            stripped = line.strip()

            if not collecting:
                if start_re.match(stripped):
                    collecting = True
                continue

            # Check stop conditions
            if stripped and any(r.search(stripped) for r in stop_res):
                break

            section_lines.append(line)

        return "\n".join(section_lines).strip()

    @staticmethod
    def _extract_instructions(text: str, disc_type: DiscoveryType) -> str:
        """Extract the instructions section from document text.

        Looks for text between "INSTRUCTIONS TO ANSWERING PARTY" and the
        first request header or section heading for the given discovery type.

        Returns empty string if no instructions section is found.
        """
        lines = text.splitlines()
        start_re = re.compile(
            r"INSTRUCTIONS\s+TO\s+ANSWERING\s+PARTY", re.IGNORECASE
        )
        stop_res = [
            re.compile(disc_type.request_header_pattern, re.IGNORECASE),
            re.compile(re.escape(disc_type.section_heading), re.IGNORECASE),
            re.compile(
                r"(?:SPECIAL\s+INTERROGATORIES|REQUESTS?\s+FOR\s+"
                r"(?:PRODUCTION|ADMISSION)).*SET",
                re.IGNORECASE,
            ),
        ]

        collecting = False
        section_lines: List[str] = []

        for line in lines:
            stripped = line.strip()

            if not collecting:
                if start_re.search(stripped):
                    collecting = True
                    section_lines.append(line)
                continue

            # Check stop conditions
            if stripped and any(r.search(stripped) for r in stop_res):
                break

            section_lines.append(line)

        return "\n".join(section_lines).strip()
