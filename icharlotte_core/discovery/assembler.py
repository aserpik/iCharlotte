"""
Document assembler for the discovery generation pipeline.

Renders DiscoverySet objects into formatted .docx files by appending
content after a caption page template.  Uses the template's built-in
styles (Body Double, Center Double Bold Und, etc.) for double-spaced
legal document formatting.
"""
import os
from datetime import date
from typing import Optional

from docx import Document as DocxDocument
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .models import DiscoverySet, DiscoveryType, number_to_word
from .declaration import generate_declaration
from .templates import extract_requests_from_text


# ---------------------------------------------------------------------------
# Style constants — these are pre-defined in the caption page template
# ---------------------------------------------------------------------------
STYLE_BODY_DOUBLE = "Body Double"               # Standard double-spaced body text
STYLE_CENTER_DOUBLE = "Center Double"            # Centered double-spaced
STYLE_CENTER_DOUBLE_BOLD_UND = "Center Double Bold Und"  # Centered bold underline
STYLE_FLUSH_LEFT_DOUBLE = "Flush Left Double"    # Left-aligned double-spaced
STYLE_DOUBLE_SPACING = "Double Spacing"          # Generic double-spaced


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _add_styled_paragraph(doc, text="", style_name=STYLE_BODY_DOUBLE,
                          bold=False, underline=False):
    """Add a paragraph using a named template style.

    Falls back to manual formatting if the style is not found.
    Returns the paragraph.
    """
    para = doc.add_paragraph()

    # Try to apply the named style
    try:
        para.style = doc.styles[style_name]
    except KeyError:
        # Style not in template — fall back to manual double-spacing
        para.paragraph_format.line_spacing = Pt(24)

    if text:
        run = para.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        if bold:
            run.bold = True
        if underline:
            run.underline = True

    return para


def _add_empty_line(doc):
    """Add a blank double-spaced paragraph."""
    _add_styled_paragraph(doc, "")


# ---------------------------------------------------------------------------
# DiscoveryAssembler
# ---------------------------------------------------------------------------

class DiscoveryAssembler:
    """Assembles DiscoverySet objects into formatted .docx files.

    Parameters
    ----------
    caption_page_path : str
        Path to a .docx caption page template.  The assembler opens this
        file, preserves its styles and formatting, then appends discovery
        content after the existing caption.
    """

    def __init__(self, caption_page_path: str):
        if not os.path.isfile(caption_page_path):
            raise FileNotFoundError(
                f"Caption page template not found: {caption_page_path}"
            )
        self.caption_page_path = caption_page_path

    # -- public API ---------------------------------------------------------

    def assemble(
        self,
        discovery_set: DiscoverySet,
        output_path: str,
        attorney_name: str = "",
        firm_name: str = "Bordin Semmer LLP",
        date_str: str = "",
    ) -> str:
        """Render a DiscoverySet into a formatted .docx file.

        Opens the caption page template, replaces the document title
        placeholder in the caption table, then appends all discovery
        sections in order.  Creates the output directory if needed.

        Returns the output_path on success.
        """
        ds = discovery_set
        if not date_str:
            date_str = date.today().strftime("%B %d, %Y")

        doc = DocxDocument(self.caption_page_path)

        # --- (a) Replace "CAPTION PAGE" placeholder in caption table ---
        self._replace_caption_title(doc, ds)

        # --- (b) Propounding / Responding party block ---
        self._insert_party_block(doc, ds)

        # --- (c) Preamble paragraph ---
        _add_empty_line(doc)
        preamble = self._build_preamble(ds)
        _add_styled_paragraph(doc, preamble, STYLE_FLUSH_LEFT_DOUBLE)

        # --- (d) Instructions to Answering Party ---
        if ds.instructions_block:
            _add_empty_line(doc)
            self._insert_instructions(doc, ds.instructions_block)

        # --- (e) Definitions block ---
        if ds.definitions_block:
            _add_empty_line(doc)
            self._insert_definitions(doc, ds.definitions_block)

        # --- (f) Section heading ---
        _add_empty_line(doc)
        heading_text = f"{ds.discovery_type.section_heading}, SET {ds.set_word.upper()}"
        _add_styled_paragraph(doc, heading_text, STYLE_CENTER_DOUBLE_BOLD_UND,
                              bold=True, underline=True)

        # --- (g) Numbered discovery requests ---
        _add_empty_line(doc)
        self._insert_requests(doc, ds)

        # --- (h) Signature block ---
        _add_empty_line(doc)
        self._insert_signature_block(doc, ds, attorney_name, firm_name, date_str)

        # --- (i) Declaration (if needed) ---
        if ds.needs_declaration:
            self._insert_declaration(doc, ds, attorney_name, firm_name, date_str)

        # Ensure output directory exists
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        doc.save(output_path)
        return output_path

    def assemble_from_plain_text(
        self,
        plain_text: str,
        discovery_set: DiscoverySet,
        output_path: str,
        attorney_name: str = "",
        firm_name: str = "Bordin Semmer LLP",
        date_str: str = "",
    ) -> str:
        """Re-parse plain text into requests, update the set, and assemble."""
        requests = extract_requests_from_text(plain_text, discovery_set.discovery_type)
        discovery_set.requests = requests
        return self.assemble(
            discovery_set, output_path,
            attorney_name=attorney_name,
            firm_name=firm_name,
            date_str=date_str,
        )

    @staticmethod
    def find_caption_page(case_path: str) -> Optional[str]:
        """Search a case folder for a .docx with 'caption page' in the filename."""
        if not os.path.isdir(case_path):
            return None

        for entry in os.listdir(case_path):
            if entry.lower().endswith(".docx") and "caption page" in entry.lower():
                return os.path.join(case_path, entry)

        # Also check one level of subdirectories
        for entry in os.listdir(case_path):
            subdir = os.path.join(case_path, entry)
            if os.path.isdir(subdir):
                for sub_entry in os.listdir(subdir):
                    if (
                        sub_entry.lower().endswith(".docx")
                        and "caption page" in sub_entry.lower()
                    ):
                        return os.path.join(subdir, sub_entry)

        return None

    # -- private helpers ----------------------------------------------------

    def _replace_caption_title(self, doc, ds: DiscoverySet):
        """Replace the 'CAPTION PAGE' placeholder in the caption table
        with the actual document title (e.g., 'DEFENDANT ...'S SPECIAL
        INTERROGATORIES TO PLAINTIFF ..., SET ONE')."""
        title = ds.discovery_type.document_title_template.format(
            propounding=f"DEFENDANT {ds.propounding_party.name.upper()}",
            responding=(
                f"{ds.directed_to.role_label.upper()} "
                f"{ds.directed_to.name.upper()}"
            ),
            set_word=ds.set_word.upper(),
        )

        # Search through all tables for "CAPTION PAGE" text and replace it
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if "CAPTION PAGE" in para.text.upper():
                            # Clear existing runs and insert title
                            for run in para.runs:
                                run.text = ""
                            if para.runs:
                                para.runs[0].text = title
                                para.runs[0].bold = True
                                para.runs[0].underline = True
                            else:
                                run = para.add_run(title)
                                run.bold = True
                                run.underline = True
                                run.font.name = "Times New Roman"
                                run.font.size = Pt(12)
                            return  # Only replace the first match

    def _insert_party_block(self, doc, ds: DiscoverySet):
        """Insert propounding/responding party identification block."""
        _add_empty_line(doc)
        _add_empty_line(doc)
        _add_styled_paragraph(
            doc,
            f"PROPOUNDING PARTY:\t{ds.propounding_party.formal_description}",
            STYLE_BODY_DOUBLE,
        )
        _add_styled_paragraph(
            doc,
            f"RESPONDING PARTY:\t{ds.directed_to.formal_description}",
            STYLE_FLUSH_LEFT_DOUBLE,
        )
        _add_styled_paragraph(
            doc,
            f"SET NO.:\t{ds.set_word.upper()} ({ds.set_number})",
            STYLE_FLUSH_LEFT_DOUBLE,
        )

    def _insert_instructions(self, doc, instructions_block: str):
        """Insert the instructions to answering party block."""
        lines = instructions_block.strip().split("\n")
        for line in lines:
            stripped = line.strip()
            if not stripped:
                _add_empty_line(doc)
                continue
            if "INSTRUCTIONS TO ANSWERING PARTY" in stripped.upper():
                _add_styled_paragraph(
                    doc, stripped, STYLE_CENTER_DOUBLE_BOLD_UND,
                    bold=True, underline=True,
                )
            else:
                _add_styled_paragraph(doc, stripped, STYLE_BODY_DOUBLE)

    def _insert_definitions(self, doc, definitions_block: str):
        """Insert the definitions block."""
        lines = definitions_block.strip().split("\n")
        for line in lines:
            stripped = line.strip()
            if not stripped:
                _add_empty_line(doc)
                continue
            if stripped.upper() in ("DEFINITIONS", "DEFINED TERMS"):
                _add_styled_paragraph(
                    doc, stripped, STYLE_CENTER_DOUBLE_BOLD_UND,
                    bold=True, underline=True,
                )
            else:
                _add_styled_paragraph(doc, stripped, STYLE_BODY_DOUBLE)

    def _insert_requests(self, doc, ds: DiscoverySet):
        """Insert numbered discovery requests with inline definitions."""
        for req in ds.requests:
            # Request header (bold + underline)
            header = ds.discovery_type.request_header_template.format(
                number=req.number
            )
            # Add colon if not already present
            if not header.endswith(":"):
                header += ":"
            _add_styled_paragraph(doc, header, STYLE_BODY_DOUBLE,
                                  bold=True, underline=True)

            # Request body
            _add_styled_paragraph(doc, req.text, STYLE_BODY_DOUBLE)

            # Inline definitions
            for defn in req.definitions:
                _add_styled_paragraph(doc, defn, STYLE_BODY_DOUBLE)

            # Visual separator between requests
            _add_empty_line(doc)

    def _insert_signature_block(self, doc, ds: DiscoverySet, attorney_name: str,
                                firm_name: str, date_str: str):
        """Insert the signature block."""
        # /// filler lines
        for _ in range(3):
            _add_styled_paragraph(doc, "///", STYLE_BODY_DOUBLE)

        _add_empty_line(doc)
        _add_styled_paragraph(
            doc,
            f"Dated:  {date_str}\t      {firm_name.upper()}",
            STYLE_BODY_DOUBLE,
        )
        _add_empty_line(doc)
        _add_empty_line(doc)
        _add_styled_paragraph(
            doc, f"By:\t______________________________",
            STYLE_FLUSH_LEFT_DOUBLE,
        )
        if attorney_name:
            _add_styled_paragraph(doc, f"\t\t{attorney_name}", STYLE_FLUSH_LEFT_DOUBLE)
        _add_styled_paragraph(
            doc,
            f"Attorneys for {ds.propounding_party.role_label},",
            STYLE_FLUSH_LEFT_DOUBLE,
        )
        _add_styled_paragraph(
            doc, ds.propounding_party.name.upper(),
            STYLE_FLUSH_LEFT_DOUBLE,
        )

    def _insert_declaration(self, doc, ds: DiscoverySet, attorney_name: str,
                            firm_name: str, date_str: str):
        """Insert the CCP declaration of necessity (page break before)."""
        decl_text = generate_declaration(ds, attorney_name, firm_name)
        if not decl_text:
            return

        decl_text = decl_text.replace("{date}", date_str)

        # Page break before declaration
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        page_break_para = doc.add_paragraph()
        run = page_break_para.add_run()
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run._element.append(br)

        # Parse declaration lines
        lines = decl_text.split("\n")
        for i, line in enumerate(lines):
            stripped = line.strip()

            if not stripped:
                _add_empty_line(doc)
                continue

            if i == 0 and stripped.startswith("DECLARATION"):
                _add_styled_paragraph(
                    doc, stripped, STYLE_CENTER_DOUBLE_BOLD_UND,
                    bold=True, underline=True,
                )
            elif stripped.startswith("____"):
                _add_styled_paragraph(doc, stripped, STYLE_FLUSH_LEFT_DOUBLE)
            else:
                _add_styled_paragraph(doc, stripped, STYLE_BODY_DOUBLE)

    @staticmethod
    def _build_preamble(ds: DiscoverySet) -> str:
        """Return the statutory preamble text for the discovery type."""
        prop_role = ds.propounding_party.role_label
        prop_name = ds.propounding_party.name.upper()
        resp_role = ds.directed_to.role_label
        resp_name = ds.directed_to.name.upper()
        set_word = number_to_word(ds.set_number)

        if ds.discovery_type == DiscoveryType.SI:
            return (
                f"TO {resp_role.upper()} {resp_name} AND "
                f"ATTORNEYS OF RECORD:\n"
                f"\tPursuant to California Code of Civil Procedure \u00a72030.030, "
                f"{prop_role}, {prop_name} "
                f'("Propounding Party" or "{prop_role}"), hereby '
                f"propounds to {resp_role}, {resp_name} "
                f'("{resp_role}" or "Responding Party"), the following '
                f"{set_word} Set of Special Interrogatories, "
                f"each of which shall be answered fully, separately, in writing, "
                f"under oath, and within thirty (30) days as required by law."
            )
        elif ds.discovery_type == DiscoveryType.RPD:
            return (
                f"TO {resp_role.upper()} {resp_name} AND "
                f"ATTORNEYS OF RECORD:\n"
                f"\tDemand is hereby made by {prop_role}, "
                f"{prop_name} "
                f'("Propounding Party" or "{prop_role}"), pursuant to '
                f"Code of Civil Procedure section 2031.010, et seq., that "
                f"{resp_role}, {resp_name} "
                f'("{resp_role}" or "Responding Party"), produce and permit '
                f"inspection, photographing, and photocopying of the documents "
                f"and/or inspection, photographing, testing, and sampling of "
                f"other tangible things described herein."
            )
        elif ds.discovery_type == DiscoveryType.RFA:
            return (
                f"TO {resp_role.upper()} {resp_name} AND "
                f"ATTORNEYS OF RECORD:\n"
                f"\tPursuant to California Code of Civil Procedure \u00a72033.010, "
                f"{prop_role}, {prop_name} "
                f'("Propounding Party" or "{prop_role}"), hereby '
                f"requests that {resp_role}, {resp_name} "
                f'("{resp_role}" or "Responding Party"), admit the truth of '
                f"the following matters within thirty (30) days as required by law."
            )
        return ""
