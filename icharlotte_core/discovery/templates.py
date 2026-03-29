"""
Template loading and variable substitution for the discovery pipeline.

Provides:
- TemplateLoader: loads standard templates, defined terms, instructions,
  and preamble blocks from .docx files in the discovery directory.
- extract_requests_from_text: parses plain text into DiscoveryRequest objects.
- substitute_variables: replaces {key} placeholders and contextual ____ blanks.
"""
import os
import re
from typing import Dict, List

from docx import Document as DocxDocument

from .models import DiscoveryRequest, DiscoveryType


# ---------------------------------------------------------------------------
# Standard type → folder name mapping
# ---------------------------------------------------------------------------

_STANDARD_TYPE_FOLDERS: Dict[str, str] = {
    "negligence": "Standard Negligence Discovery (5800.070)",
    "wrongful_death": "Standard Wrongful Death Discovery",
}

# Discovery type abbreviation → template filename
_TEMPLATE_FILENAMES: Dict[str, str] = {
    "SI": "SI(1) tPltf.docx",
    "RPD": "RPD(1) tPltf.docx",
    "RFA": "RFA(1) tPltf.docx",
}


# ---------------------------------------------------------------------------
# Module-level functions
# ---------------------------------------------------------------------------

def extract_requests_from_text(
    text: str, disc_type: DiscoveryType
) -> List[DiscoveryRequest]:
    """Parse plain text into a list of :class:`DiscoveryRequest` objects.

    Splits on request headers matching ``disc_type.request_header_pattern``.
    Lines starting with ``(`` after the request body are treated as inline
    definitions.  Empty lines and ``///`` lines are skipped.

    For templates that lack explicit numbered headers (common with RPD
    templates), requests are identified as non-empty, non-definition content
    blocks separated by blank lines and auto-numbered sequentially.
    """
    if not text or not text.strip():
        return []

    pattern = disc_type.request_header_pattern  # e.g. r"SPECIAL\s+INTERROGATORY\s+NO\.\s*(\d+)"

    # Try header-based splitting first
    requests = _extract_with_headers(text, pattern)
    if requests:
        return requests

    # Fallback: auto-number blocks for templates without explicit numbered headers.
    # This is common for RPD templates where requests are just body paragraphs
    # separated by blank lines under a section heading.
    return _extract_without_headers(text, disc_type)


def _extract_with_headers(text: str, pattern: str) -> List[DiscoveryRequest]:
    """Split text on numbered request headers and extract requests."""
    # Find all header matches (with their positions)
    header_re = re.compile(pattern, re.IGNORECASE)
    matches = list(header_re.finditer(text))

    if not matches:
        return []

    requests: List[DiscoveryRequest] = []
    for i, match in enumerate(matches):
        number = int(match.group(1))

        # Body starts after the header (skip optional colon and whitespace)
        body_start = match.end()
        # Skip colon right after header if present
        rest = text[body_start:]
        if rest and rest[0] == ":":
            body_start += 1

        # Body ends at the next header or end of text
        if i + 1 < len(matches):
            body_end = matches[i + 1].start()
        else:
            body_end = len(text)

        body_text = text[body_start:body_end]
        req = _parse_request_body(number, body_text)
        if req.text:  # Only include requests with actual content
            requests.append(req)

    return requests


def _extract_without_headers(
    text: str, disc_type: DiscoveryType
) -> List[DiscoveryRequest]:
    """Extract requests from text without explicit numbered headers.

    Used for templates (like RPD) where requests are body paragraphs
    separated by blank lines under a section heading.
    """
    lines = text.splitlines()

    # Find the section heading to start parsing after it
    section_heading = disc_type.section_heading
    start_idx = 0
    heading_pattern = re.compile(re.escape(section_heading), re.IGNORECASE)
    # Also match partial headings like "REQUEST FOR PRODUCTION OF DOCUMENTS, SET ONE"
    alt_patterns = {
        DiscoveryType.RPD: re.compile(
            r"REQUEST\s+FOR\s+PRODUCTION\s+OF\s+DOCUMENTS,\s+SET", re.IGNORECASE
        ),
        DiscoveryType.RFA: re.compile(
            r"REQUESTS?\s+FOR\s+ADMISSION,?\s+SET", re.IGNORECASE
        ),
        DiscoveryType.SI: re.compile(
            r"SPECIAL\s+INTERROGATORIES,?\s+SET", re.IGNORECASE
        ),
    }

    for idx, line in enumerate(lines):
        stripped = line.strip()
        if heading_pattern.search(stripped):
            start_idx = idx + 1
            break
        alt = alt_patterns.get(disc_type)
        if alt and alt.search(stripped):
            start_idx = idx + 1
            break

    # Parse blocks separated by blank lines
    requests: List[DiscoveryRequest] = []
    current_body_lines: List[str] = []
    current_def_lines: List[str] = []
    request_number = 0

    def flush_block():
        nonlocal request_number
        body = "\n".join(current_body_lines).strip()
        if body:
            request_number += 1
            definitions = [d.strip() for d in current_def_lines if d.strip()]
            requests.append(DiscoveryRequest(
                number=request_number,
                text=body,
                definitions=definitions,
            ))
        current_body_lines.clear()
        current_def_lines.clear()

    # Track whether we're inside a definition block vs body
    for line in lines[start_idx:]:
        stripped = line.strip()

        # Skip empty and /// lines (they act as separators)
        if not stripped or stripped == "///":
            if current_body_lines or current_def_lines:
                # Check if the next non-empty line is a definition continuation
                # If we have body text, this blank line might separate requests
                # We only flush if the current block has body content
                if current_body_lines:
                    flush_block()
            continue

        # Stop at signature/date blocks
        if stripped.startswith("Dated:") or stripped.startswith("Dated "):
            break

        # Definition line (starts with parenthesis)
        if stripped.startswith("(") or stripped.startswith('"'):
            # Definitions can appear after request body or between requests
            if current_body_lines:
                current_def_lines.append(stripped)
            else:
                # Definition without preceding body -- append to previous request
                if requests:
                    requests[-1].definitions.append(stripped)
            continue

        # Regular body line
        current_body_lines.append(stripped)

    # Flush any remaining block
    flush_block()

    return requests


def _parse_request_body(number: int, body_text: str) -> DiscoveryRequest:
    """Parse body text into request text and inline definitions."""
    lines = body_text.splitlines()
    text_lines: List[str] = []
    definitions: List[str] = []

    for line in lines:
        stripped = line.strip()

        # Skip empty lines and ///
        if not stripped or stripped == "///":
            continue

        # Stop at signature/date blocks
        if stripped.startswith("Dated:") or stripped.startswith("Dated "):
            break

        # Definition line (starts with parenthesis)
        if stripped.startswith("("):
            definitions.append(stripped)
        else:
            # Only add to text_lines if we haven't started collecting definitions
            # (definitions always come after body text)
            if not definitions:
                text_lines.append(stripped)
            else:
                # Non-definition line after definitions -- this is unusual
                # but could happen; treat it as continuation of body
                text_lines.append(stripped)

    text = " ".join(text_lines).strip()
    return DiscoveryRequest(number=number, text=text, definitions=definitions)


def substitute_variables(text: str, variables: Dict[str, str]) -> str:
    """Replace placeholders and contextual blanks in *text*.

    - ``{key}`` named placeholders are replaced with values from *variables*.
    - Contextual ``____`` blanks near "Responding Party" / "Plaintiff" are
      replaced with ``responding_party_name`` from *variables*.
    - Contextual ``____`` blanks near "Propounding Party" / "Defendant" are
      replaced with ``propounding_party_name`` from *variables*.
    - Unresolved ``____`` blanks are left visible as a marker.
    """
    result = text

    # 1) Replace {key} placeholders (only for keys present in variables)
    def replace_named(match: re.Match) -> str:
        key = match.group(1)
        if key in variables:
            return variables[key]
        return match.group(0)  # leave unresolved

    result = re.sub(r"\{(\w+)\}", replace_named, result)

    # 2) Contextual ____ replacements (4+ underscores)
    responding_name = variables.get("responding_party_name", "")
    propounding_name = variables.get("propounding_party_name", "")

    if responding_name:
        # "Responding Party, ____" or "Plaintiff, ____" or "Plaintiff ____"
        result = re.sub(
            r"(Responding\s+Party,?\s+)_{4,}",
            rf"\g<1>{responding_name}",
            result,
            flags=re.IGNORECASE,
        )
        result = re.sub(
            r"(Plaintiff,?\s+)_{4,}",
            rf"\g<1>{responding_name}",
            result,
            flags=re.IGNORECASE,
        )

    if propounding_name:
        # "Propounding Party, ____" or "Defendant, ____" or "Defendant ____"
        result = re.sub(
            r"(Propounding\s+Party,?\s+)_{4,}",
            rf"\g<1>{propounding_name}",
            result,
            flags=re.IGNORECASE,
        )
        result = re.sub(
            r"(Defendant,?\s+)_{4,}",
            rf"\g<1>{propounding_name}",
            result,
            flags=re.IGNORECASE,
        )

    return result


# ---------------------------------------------------------------------------
# TemplateLoader class
# ---------------------------------------------------------------------------

class TemplateLoader:
    """Load discovery templates from .docx files in a discovery directory.

    Parameters
    ----------
    discovery_dir : str
        Path to the root discovery directory containing template folders.
    """

    def __init__(self, discovery_dir: str):
        self.discovery_dir = discovery_dir

    # -- public API ---------------------------------------------------------

    def load_standard_requests(
        self, standard_type: str, disc_type: DiscoveryType
    ) -> List[DiscoveryRequest]:
        """Load requests from a standard template .docx file.

        Parameters
        ----------
        standard_type : str
            Key for the standard template folder (e.g. ``"negligence"``).
        disc_type : DiscoveryType
            The type of discovery (SI, RPD, RFA).

        Returns
        -------
        List[DiscoveryRequest]
            Parsed requests from the template.
        """
        folder = _STANDARD_TYPE_FOLDERS.get(standard_type)
        if not folder:
            raise ValueError(
                f"Unknown standard type: {standard_type!r}. "
                f"Available: {list(_STANDARD_TYPE_FOLDERS.keys())}"
            )
        filename = _TEMPLATE_FILENAMES.get(disc_type.abbreviation)
        if not filename:
            raise ValueError(
                f"No template filename configured for {disc_type.abbreviation}"
            )

        path = os.path.join(self.discovery_dir, folder, filename)
        text = self._extract_text(path)
        return extract_requests_from_text(text, disc_type)

    def load_defined_terms(self) -> str:
        """Load the standard defined terms document as plain text."""
        path = os.path.join(self.discovery_dir, "DISCOVERY DEFINED TERMS.docx")
        return self._extract_text(path)

    def load_instructions(
        self, disc_type: DiscoveryType, template_path: str
    ) -> str:
        """Extract the 'INSTRUCTIONS TO ANSWERING PARTY' section from a template.

        Returns the text between the "INSTRUCTIONS TO ANSWERING PARTY" heading
        and the first request header or section heading.
        """
        text = self._extract_text(template_path)
        return self._extract_section(
            text,
            start_marker=r"INSTRUCTIONS\s+TO\s+ANSWERING\s+PARTY",
            stop_patterns=[
                disc_type.request_header_pattern,
                re.escape(disc_type.section_heading),
                # Also stop at section heading variants with set number
                r"(?:SPECIAL\s+INTERROGATORIES|REQUESTS?\s+FOR\s+(?:PRODUCTION|ADMISSION)).*SET",
            ],
        )

    def load_preamble(
        self, disc_type: DiscoveryType, template_path: str
    ) -> str:
        """Extract the 'TO ... ATTORNEYS OF RECORD' preamble block.

        Returns the text from the "TO ..." line through the text before
        the "INSTRUCTIONS TO ANSWERING PARTY" heading.
        """
        text = self._extract_text(template_path)
        return self._extract_section(
            text,
            start_marker=r"TO\s+.+ATTORNEYS\s+OF\s+RECORD",
            stop_patterns=[
                r"INSTRUCTIONS\s+TO\s+ANSWERING\s+PARTY",
            ],
        )

    # -- private helpers ----------------------------------------------------

    @staticmethod
    def _extract_text(docx_path: str) -> str:
        """Extract all paragraph text from a .docx file."""
        if not os.path.isfile(docx_path):
            raise FileNotFoundError(f"Template not found: {docx_path}")
        doc = DocxDocument(docx_path)
        return "\n".join(p.text for p in doc.paragraphs)

    @staticmethod
    def _extract_section(
        text: str,
        start_marker: str,
        stop_patterns: List[str],
    ) -> str:
        """Extract a section of text between a start marker and stop patterns.

        The start marker line is included in the output.
        """
        lines = text.splitlines()
        start_re = re.compile(start_marker, re.IGNORECASE)
        stop_res = [re.compile(p, re.IGNORECASE) for p in stop_patterns]

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
