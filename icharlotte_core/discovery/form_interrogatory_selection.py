"""Selection helpers for flattened Judicial Council Form Interrogatories."""
from __future__ import annotations

from dataclasses import replace
import re
from typing import Iterable

from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_type_detector import normalize_discovery_type


_NUMBER_RE = re.compile(r"^\d+\.\d+$")
_NUMBER_LINE_RE = re.compile(r"^\s*(\d+\.\d+)\b\.?\s*(.*)$")


def extract_selected_form_interrogatory_numbers(pdf_path: str) -> list[str]:
    """Return selected FI numbers from a flattened PDF checkbox form.

    The Judicial Council FROG form often arrives as a flattened PDF. In that
    form, checkboxes and X marks are vector drawings rather than form fields.
    This reads those vector rectangles/marks and maps selected boxes to the
    interrogatory number printed on the same line.
    """
    try:
        import fitz
    except ImportError:
        return []

    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return []

    selected: set[str] = set()
    try:
        for page in doc:
            numbers = _number_words(page, fitz)
            marks = _selected_mark_rects(page, fitz)
            for box in _checkbox_rects(page, fitz):
                if not _box_is_selected(box, marks, fitz):
                    continue
                number = _number_for_box(box, numbers)
                if number:
                    selected.add(number)
    finally:
        doc.close()
    return sorted(selected, key=_number_key)


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

    existing_by_number = {
        req.number.strip(): req
        for req in filtered.requests
        if req.number and req.number.strip()
    }
    missing_numbers = [
        number for number in selected if number not in existing_by_number
    ]
    if not missing_numbers:
        return filtered

    fallback_by_number = {
        req.number: req
        for req in extract_form_interrogatory_requests(pdf_path, missing_numbers)
    }
    completed_requests = []
    for number in selected:
        if number in existing_by_number:
            completed_requests.append(existing_by_number[number])
        else:
            completed_requests.append(
                fallback_by_number.get(
                    number,
                    ParsedRequest(
                        number=number,
                        text=f"Form Interrogatory No. {number}.",
                    ),
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
            _collect_numbered_lines(page.get_text("text"), wanted, text_by_number)
    finally:
        doc.close()

    requests = []
    for number in sorted(wanted, key=_number_key):
        text = text_by_number.get(number, "").strip()
        if text:
            requests.append(ParsedRequest(number=number, text=text))
    return requests


def _collect_numbered_lines(
    page_text: str,
    wanted: set[str],
    text_by_number: dict[str, str],
) -> None:
    current_number = ""
    current_parts: list[str] = []

    def flush() -> None:
        if current_number and current_number in wanted and current_number not in text_by_number:
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


def _checkbox_rects(page, fitz) -> list:
    rects = []
    for drawing in page.get_drawings():
        rect = fitz.Rect(drawing.get("rect"))
        if not rect or rect.is_empty:
            continue
        if drawing.get("type") not in {"s", "fs"}:
            continue
        if 14 <= rect.width <= 24 and 6 <= rect.height <= 14:
            rects.append(rect)
    return rects


def _selected_mark_rects(page, fitz) -> list:
    marks = []
    for drawing in page.get_drawings():
        rect = fitz.Rect(drawing.get("rect"))
        if not rect or rect.is_empty:
            continue
        fill = drawing.get("fill")
        if fill != (0.0, 0.0, 0.0):
            continue
        if 2 <= rect.width <= 9 and 2 <= rect.height <= 9:
            marks.append(rect)
    return marks


def _number_words(page, fitz) -> list[tuple[str, object]]:
    words = []
    for word in page.get_text("words"):
        text = word[4]
        if _NUMBER_RE.match(text):
            words.append((text, fitz.Rect(word[:4])))
    return words


def _box_is_selected(box, marks, fitz) -> bool:
    expanded = fitz.Rect(box.x0 - 2, box.y0 - 2, box.x1 + 2, box.y1 + 2)
    return any(expanded.contains(mark) for mark in marks)


def _number_for_box(box, numbers: list[tuple[str, object]]) -> str:
    candidates = []
    box_center_y = (box.y0 + box.y1) / 2
    for number, rect in numbers:
        ydist = abs(((rect.y0 + rect.y1) / 2) - box_center_y)
        xdist = rect.x0 - box.x1
        if rect.x0 >= box.x1 - 2 and ydist < 12 and xdist < 90:
            candidates.append((xdist, number))
    if not candidates:
        for number, rect in numbers:
            ydist = abs(((rect.y0 + rect.y1) / 2) - box_center_y)
            xdist = abs(rect.x0 - box.x1)
            if ydist < 12 and xdist < 120:
                candidates.append((xdist, number))
    return sorted(candidates)[0][1] if candidates else ""


def _number_key(number: str) -> tuple[int, ...]:
    parts = []
    for piece in str(number).split("."):
        try:
            parts.append(int(piece))
        except ValueError:
            parts.append(0)
    return tuple(parts)
