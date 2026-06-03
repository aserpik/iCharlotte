"""Selection helpers for flattened Judicial Council Form Interrogatories."""
from __future__ import annotations

from dataclasses import dataclass, field, replace
import re
from typing import Iterable

from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery,
    ParsedRequest,
    detect_compound,
    extract_defined_terms,
)
from icharlotte_core.discovery.response_type_detector import normalize_discovery_type


_NUMBER_RE = re.compile(r"^\d+\.\d+$")
_NUMBER_LINE_RE = re.compile(r"^\s*(\d+\.\d+)\b\.?\s*(.*)$")
# An interrogatory label is "N.M" with M >= 1; "N.0" is a section header.
_INTERROGATORY_RE = re.compile(r"^(\d+)\.(\d+)$")


@dataclass
class ScannedInterrogatory:
    """A Form Interrogatory found on the page, with its auto-detected state."""

    number: str
    checked: bool = False
    text: str = ""


def extract_selected_form_interrogatory_numbers(pdf_path: str) -> list[str]:
    """Return selected FI numbers from a flattened PDF checkbox form.

    The Judicial Council FROG form usually arrives as a flattened PDF in which
    checkboxes are vector drawings, not form fields. A box is drawn either as a
    single rectangle or (on the official DISC-001) as four thin edge-slivers,
    and a selection is a small filled glyph/stroke (a checkmark or X) inside the
    box. This anchors on the printed interrogatory number, then looks just to
    its left for such a mark.
    """
    return [item.number for item in scan_form_interrogatories(pdf_path) if item.checked]


def scan_form_interrogatories(pdf_path: str) -> list[ScannedInterrogatory]:
    """List every Form Interrogatory checkbox in the document.

    Returns one entry per interrogatory that has a checkbox, with ``checked``
    set from best-effort mark detection and ``text`` pulled from the text layer.
    This powers both auto-detection and the manual confirmation list, so the
    attorney can review/override what was detected.
    """
    try:
        import fitz
    except ImportError:
        return []

    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return []

    order: list[str] = []
    checked_by_number: dict[str, bool] = {}
    text_by_number: dict[str, str] = {}
    try:
        for page in doc:
            _collect_numbered_lines(
                _page_text_in_reading_order(page), None, text_by_number
            )
            drawings = page.get_drawings()
            for number, rect in _interrogatory_number_rects(page, fitz):
                region = _checkbox_region(rect, fitz)
                if not _region_has_checkbox(drawings, region, fitz):
                    continue
                is_checked = _region_has_mark(drawings, region, fitz)
                if number not in checked_by_number:
                    order.append(number)
                    checked_by_number[number] = is_checked
                elif is_checked:
                    checked_by_number[number] = True
    finally:
        doc.close()

    order.sort(key=_number_key)
    return [
        ScannedInterrogatory(
            number=number,
            checked=checked_by_number[number],
            text=text_by_number.get(number, "").strip(),
        )
        for number in order
    ]


def filter_parsed_form_interrogatories(
    parsed: ParsedDiscovery,
    selected_numbers: Iterable[str],
) -> ParsedDiscovery:
    """Filter parsed FI requests to checked-box numbers when available."""
    discovery_type = normalize_discovery_type(parsed.discovery_type)
    if discovery_type != "FI":
        return replace(parsed, discovery_type=discovery_type)

    selected = {str(number).strip() for number in selected_numbers if str(number).strip()}
    if not selected:
        return replace(parsed, discovery_type="FI")

    requests = [req for req in parsed.requests if req.number.strip() in selected]
    return replace(parsed, discovery_type="FI", requests=requests)


def complete_selected_form_interrogatories(
    parsed: ParsedDiscovery,
    pdf_path: str,
    selected_numbers: Iterable[str] | None = None,
) -> ParsedDiscovery:
    """Filter FI requests to checked boxes and fill missing selected rows."""
    selected = list(
        selected_numbers
        if selected_numbers is not None
        else extract_selected_form_interrogatory_numbers(pdf_path)
    )
    filtered = filter_parsed_form_interrogatories(parsed, selected)
    if normalize_discovery_type(filtered.discovery_type) != "FI" or not selected:
        return filtered

    llm_by_number = {
        req.number.strip(): req
        for req in filtered.requests
        if req.number and req.number.strip()
    }
    # The FROG question text is canonical and standard, so prefer the text read
    # straight from the form (column-aware) over the parse LLM's version — the
    # LLM is fed two-column text that interleaves adjacent interrogatories and
    # is told to skip unchecked rows, so its FI text is unreliable.
    extracted_by_number = {
        req.number: req
        for req in extract_form_interrogatory_requests(pdf_path, selected)
    }
    completed_requests = []
    for number in selected:
        extracted = extracted_by_number.get(number)
        if extracted is not None and extracted.text.strip():
            completed_requests.append(extracted)
        elif number in llm_by_number:
            completed_requests.append(llm_by_number[number])
        else:
            completed_requests.append(
                ParsedRequest(
                    number=number,
                    text=f"Form Interrogatory No. {number}.",
                )
            )
    return replace(filtered, discovery_type="FI", requests=completed_requests)


def extract_form_interrogatory_requests(
    pdf_path: str,
    numbers: Iterable[str],
) -> list[ParsedRequest]:
    """Extract visible FROG row text for selected numbers from the PDF text layer."""
    wanted = {str(number).strip() for number in numbers if str(number).strip()}
    if not wanted:
        return []
    try:
        import fitz
    except ImportError:
        return []

    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return []

    text_by_number: dict[str, str] = {}
    try:
        for page in doc:
            _collect_numbered_lines(
                _page_text_in_reading_order(page), wanted, text_by_number
            )
    finally:
        doc.close()

    requests = []
    for number in sorted(wanted, key=_number_key):
        text = text_by_number.get(number, "").strip()
        if text:
            requests.append(
                ParsedRequest(
                    number=number,
                    text=text,
                    is_compound=detect_compound(text),
                    defined_terms_used=extract_defined_terms(text),
                )
            )
    return requests


def _page_text_in_reading_order(page) -> str:
    """Return page text in column-aware reading order.

    The Judicial Council FROG is laid out in two columns. ``get_text("text")``
    (and even ``"blocks"``) reads roughly across the full page width, so the two
    columns interleave line-by-line — e.g. interrogatory 3.1's sub-parts (left
    column) get mixed with 4.1's (right column), and a row's text is cut short
    by a contents-list entry from the other column. fitz groups some rows into
    single full-width blocks whose *text* is already interleaved, so block-level
    sorting is not enough. Reconstruct from individual words instead: assign each
    word to a column by its own x-center, rebuild lines within each column, and
    emit the left column fully before the right.
    """
    try:
        words = page.get_text("words")  # (x0, y0, x1, y1, word, block, line, n)
    except Exception:
        return page.get_text("text")
    if not words:
        return page.get_text("text")

    mid_x = page.rect.width / 2.0
    columns: tuple[list, list] = ([], [])
    for word in words:
        center_x = (word[0] + word[2]) / 2.0
        columns[0 if center_x < mid_x else 1].append(word)

    lines: list[str] = []
    for column_words in columns:
        column_words.sort(key=lambda w: (round(w[1], 1), w[0]))
        current: list = []
        line_y: float | None = None
        for word in column_words:
            y0 = word[1]
            if line_y is not None and abs(y0 - line_y) > 3.0:
                lines.append(" ".join(w[4] for w in current))
                current = []
                line_y = None
            if line_y is None:
                line_y = y0
            current.append(word)
        if current:
            lines.append(" ".join(w[4] for w in current))
    return "\n".join(lines)


def _collect_numbered_lines(
    page_text: str,
    wanted: set[str] | None,
    text_by_number: dict[str, str],
) -> None:
    """Map each line-starting ``N.M`` to its text. ``wanted=None`` captures all."""
    current_number = ""
    current_parts: list[str] = []

    def flush() -> None:
        if (
            current_number
            and (wanted is None or current_number in wanted)
            and current_number not in text_by_number
        ):
            text = " ".join(part.strip() for part in current_parts if part.strip())
            text_by_number[current_number] = re.sub(r"\s+", " ", text).strip()

    for raw_line in (page_text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = _NUMBER_LINE_RE.match(line)
        if match:
            flush()
            current_number = match.group(1)
            current_parts = [match.group(2).strip()]
            continue
        if current_number:
            current_parts.append(line)
    flush()


def _interrogatory_number_rects(page, fitz) -> list[tuple[str, object]]:
    """Word rects for interrogatory labels (``N.M`` with M >= 1)."""
    out = []
    for word in page.get_text("words"):
        match = _INTERROGATORY_RE.match(word[4])
        if not match or match.group(2) == "0":
            continue
        out.append((word[4], fitz.Rect(word[:4])))
    return out


def _checkbox_region(number_rect, fitz):
    """The area where a checkbox sits: just left of the interrogatory number."""
    return fitz.Rect(
        number_rect.x0 - 20,
        number_rect.y0 - 3,
        number_rect.x0 - 1,
        number_rect.y1 + 3,
    )


def _region_has_checkbox(drawings, region, fitz) -> bool:
    """True if a checkbox outline sits in the region (single rect or edge-slivers).

    Used to tell a real interrogatory label apart from an inline numeric
    reference (e.g., a statute like ``2033.710``) that has no box.
    """
    expanded = fitz.Rect(region.x0 - 3, region.y0 - 2, region.x1 + 5, region.y1 + 2)
    for drawing in drawings:
        rect = fitz.Rect(drawing.get("rect"))
        if rect.is_empty or not expanded.intersects(rect):
            continue
        width, height = rect.width, rect.height
        item_kinds = {item[0] for item in drawing.get("items", [])}
        # A box drawn as a single rectangle outline.
        if 8 <= width <= 22 and 6 <= height <= 16 and "re" in item_kinds:
            return True
        # A box edge drawn as a thin sliver (DISC-001 renders four of these).
        if min(width, height) < 2.0 and 5 <= max(width, height) <= 22:
            return True
    return False


def _region_has_mark(drawings, region, fitz) -> bool:
    """True if a checkmark/X glyph sits inside the region.

    A mark is a small filled glyph (the common case) or a small stroked path of
    lines/curves — but not a plain rectangle outline (that is the empty box) and
    not a thin border sliver.
    """
    for drawing in drawings:
        rect = fitz.Rect(drawing.get("rect"))
        if rect.is_empty:
            continue
        center_x = (rect.x0 + rect.x1) / 2
        center_y = (rect.y0 + rect.y1) / 2
        if not (region.x0 <= center_x <= region.x1 and region.y0 <= center_y <= region.y1):
            continue
        width, height = rect.width, rect.height
        if min(width, height) < 2.0 or max(width, height) > 14:
            continue
        item_kinds = {item[0] for item in drawing.get("items", [])}
        if drawing.get("fill") is not None:
            return True
        if (item_kinds & {"l", "c", "qu"}) and "re" not in item_kinds:
            return True
    return False


def _number_key(number: str) -> tuple[int, ...]:
    parts = []
    for piece in str(number).split("."):
        try:
            parts.append(int(piece))
        except ValueError:
            parts.append(0)
    return tuple(parts)
