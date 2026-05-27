"""Word assembly helpers for opposition memorandum previews."""

from __future__ import annotations

import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from icharlotte_core.word_validator import validate_opposition_docx

from .models import DraftDocument


def assemble_opposition_preview(
    *,
    draft: DraftDocument,
    output_path: str,
    caption_path: str = "",
) -> str:
    """Render an opposition draft to a .docx preview and validate the file."""
    if caption_path and os.path.isfile(caption_path):
        if _same_path(caption_path, output_path):
            raise ValueError("Output path must be different from the caption template path")
        doc = Document(caption_path)
        if not _replace_caption_markers(doc, draft.title or "Opposition"):
            _add_title(doc, draft.title or "Opposition")
    else:
        doc = Document()
        _add_title(doc, draft.title or "Opposition")

    _add_paragraphs(doc, draft.body_text)

    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    doc.save(output_path)

    validation = validate_opposition_docx(output_path)
    if validation.has_errors:
        messages = "; ".join(str(finding) for finding in validation.findings)
        raise ValueError(f"Generated opposition failed validation: {messages}")

    return output_path


def _add_title(doc, title: str) -> None:
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(title)
    run.bold = True
    _format_run(run)


def _add_paragraphs(doc, body_text: str) -> None:
    for raw_line in (body_text or "").splitlines():
        line = raw_line.strip()
        if not line:
            doc.add_paragraph("")
            continue

        paragraph = doc.add_paragraph()
        run = paragraph.add_run(line)
        _format_run(run)
        if _looks_like_heading(line):
            run.bold = True


def _format_run(run) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)


def _looks_like_heading(line: str) -> bool:
    return line.isupper() or line.startswith(("I.", "II.", "III.", "IV.", "V."))


def _same_path(first: str, second: str) -> bool:
    if not first or not second:
        return False
    return os.path.normcase(os.path.abspath(first)) == os.path.normcase(
        os.path.abspath(second)
    )


def _replace_caption_markers(doc, title: str) -> bool:
    """Replace CAPTION PAGE markers in body, headers, and footers."""
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w_t = f"{{{w_ns}}}t"
    w_p = f"{{{w_ns}}}p"
    w_r = f"{{{w_ns}}}r"
    replaced = 0

    for root in _iter_story_roots(doc):
        for paragraph_element in list(root.iter(w_p)):
            paragraph_text = "".join(
                text_element.text or "" for text_element in paragraph_element.iter(w_t)
            )
            if "CAPTION PAGE" not in paragraph_text.upper():
                continue

            for run_element in list(paragraph_element.iter(w_r)):
                parent = run_element.getparent()
                if parent is not None:
                    parent.remove(run_element)

            paragraph_element.append(_make_title_run(title))
            replaced += 1

    return replaced > 0


def _iter_story_roots(doc):
    seen: set[int] = set()
    roots = [doc.element.body]
    for section in doc.sections:
        roots.extend(
            [
                section.header._element,
                section.first_page_header._element,
                section.even_page_header._element,
                section.footer._element,
                section.first_page_footer._element,
                section.even_page_footer._element,
            ]
        )
    for root in roots:
        root_id = id(root)
        if root_id in seen:
            continue
        seen.add(root_id)
        yield root


def _make_title_run(text: str):
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    properties.append(OxmlElement("w:b"))
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    properties.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "24")
    properties.append(size)
    complex_size = OxmlElement("w:szCs")
    complex_size.set(qn("w:val"), "24")
    properties.append(complex_size)
    run.append(properties)
    text_node = OxmlElement("w:t")
    text_node.text = text
    text_node.set(qn("xml:space"), "preserve")
    run.append(text_node)
    return run
