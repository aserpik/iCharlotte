"""Pure I/O helpers reused by the wizard page and the parse worker."""
from __future__ import annotations

import os

from icharlotte_core.doc_library.library import DocumentLibrary
from icharlotte_core.discovery.form_interrogatory_selection import (
    complete_selected_form_interrogatories,
    extract_selected_form_interrogatory_numbers,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery
from icharlotte_core.discovery.response_type_detector import normalize_discovery_type

try:
    import fitz
except ImportError:  # pragma: no cover - depends on local install
    fitz = None


def _read_cached_document_text(path: str, case_root: str | None) -> str | None:
    if not case_root:
        return None
    try:
        text, _method, error = DocumentLibrary(case_root).get_or_extract_text(path)
    except Exception:  # noqa: BLE001 - direct reader below preserves old behavior
        return None
    if error:
        return None
    return text


def read_document_text(path: str, case_root: str | None = None) -> str:
    """Extract text from a supported context or discovery file."""
    if not path or not os.path.isfile(path):
        return ""
    lower = path.lower()
    if lower.endswith((".pdf", ".docx", ".txt")):
        cached = _read_cached_document_text(path, case_root)
        if cached is not None:
            return cached
    if lower.endswith(".pdf"):
        if not fitz:
            return ""
        doc = fitz.open(path)
        try:
            return "\n".join(page.get_text() for page in doc)
        finally:
            doc.close()
    if lower.endswith(".docx"):
        from docx import Document
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    if lower.endswith(".txt"):
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return fh.read()
    return ""


def read_first_page_text(path: str) -> str:
    """Read first-page text for type detection without parsing the whole PDF."""
    if not path or not path.lower().endswith(".pdf") or not fitz:
        return ""
    if not os.path.isfile(path):
        return ""
    doc = fitz.open(path)
    try:
        if len(doc) == 0:
            return ""
        return doc[0].get_text()
    finally:
        doc.close()


def normalize_and_filter_parsed_discovery(
    parsed: ParsedDiscovery,
    detected_type: str,
    discovery_file: str,
    selected_fi_numbers: list[str] | None = None,
) -> ParsedDiscovery:
    """Canonicalize discovery type and keep only checked FROG items.

    ``selected_fi_numbers`` overrides automatic checkbox detection — pass the
    attorney-confirmed selection here so FI responses cover exactly those
    interrogatories regardless of how reliably the checkboxes auto-detected.
    """
    normalized_detected = normalize_discovery_type(detected_type)
    parsed.discovery_type = normalized_detected or normalize_discovery_type(
        parsed.discovery_type
    )
    if parsed.discovery_type != "FI":
        return parsed
    selected_numbers = (
        list(selected_fi_numbers)
        if selected_fi_numbers is not None
        else extract_selected_form_interrogatory_numbers(discovery_file)
    )
    return complete_selected_form_interrogatories(
        parsed, discovery_file, selected_numbers,
    )
