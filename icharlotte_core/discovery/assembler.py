"""
Document assembler for the discovery generation pipeline.

Renders DiscoverySet objects into formatted .docx files by appending
content after a caption page template.  Handles document title, party
blocks, preamble, instructions, definitions, numbered requests, signature
block, and CCP declaration of necessity.
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
# Formatting helpers
# ---------------------------------------------------------------------------

def _add_paragraph(doc, text="", bold=False, underline=False, centered=False,
                   space_after=None, space_before=None, first_line_indent=None):
    """Add a paragraph with Times New Roman 12pt formatting.

    Returns the paragraph so callers can further customise it.
    """
    para = doc.add_paragraph()
    if centered:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if space_after is not None:
        para.paragraph_format.space_after = space_after
    if space_before is not None:
        para.paragraph_format.space_before = space_before
    if first_line_indent is not None:
        para.paragraph_format.first_line_indent = first_line_indent

    if text:
        run = para.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.bold = bold
        run.underline = underline

    return para


def _add_empty_line(doc):
    """Add a blank paragraph (visual spacer)."""
    _add_paragraph(doc, "")


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

        Opens the caption page template, inserts the document title, then
        appends all discovery sections in order.  Creates the output
        directory if needed.

        Returns the output_path on success.
        """
        ds = discovery_set
        if not date_str:
            date_str = date.today().strftime("%B %d, %Y")

        doc = DocxDocument(self.caption_page_path)

        # --- (a) Document title (after caption) ---
        self._insert_document_title(doc, ds)

        # --- (b) Propounding / Responding party block ---
        self._insert_party_block(doc, ds)

        # --- (c) Preamble paragraph ---
        preamble = self._build_preamble(ds)
        _add_empty_line(doc)
        _add_paragraph(doc, preamble)

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
        _add_paragraph(doc, heading_text, bold=True, underline=True, centered=True)

        # --- (g) Numbered discovery requests ---
        _add_empty_line(doc)
        self._insert_requests(doc, ds)

        # --- (h) Signature block ---
        _add_empty_line(doc)
        self._insert_signature_block(doc, attorney_name, firm_name, date_str)

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
        """Re-parse plain text into requests, update the set, and assemble.

        Uses ``extract_requests_from_text`` to parse *plain_text* into
        DiscoveryRequest objects, replaces discovery_set.requests, then
        delegates to :meth:`assemble`.
        """
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
        """Search a case folder for a .docx with 'caption page' in the filename.

        Returns the full path if found, otherwise None.
        """
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

    def _insert_document_title(self, doc, ds: DiscoverySet):
        """Insert the document title after the caption page content."""
        propounding_label = ds.propounding_party.formal_description
        responding_label = ds.directed_to.formal_description

        # Use the propounding party's name for the title (uppercase)
        propounding_name = ds.propounding_party.name.upper()
        responding_name = ds.directed_to.name.upper()

        title = ds.discovery_type.document_title_template.format(
            propounding=propounding_name,
            responding=responding_name,
            set_word=ds.set_word.upper(),
        )

        _add_empty_line(doc)
        _add_paragraph(doc, title, bold=True, underline=True, centered=True)

    def _insert_party_block(self, doc, ds: DiscoverySet):
        """Insert propounding/responding party identification block."""
        _add_empty_line(doc)
        _add_paragraph(
            doc,
            f"PROPOUNDING PARTY: {ds.propounding_party.formal_description}",
            bold=True,
        )
        _add_paragraph(
            doc,
            f"RESPONDING PARTY: {ds.directed_to.formal_description}",
            bold=True,
        )

    def _insert_instructions(self, doc, instructions_block: str):
        """Insert the instructions to answering party block."""
        lines = instructions_block.strip().split("\n")
        for line in lines:
            stripped = line.strip()
            if not stripped:
                _add_empty_line(doc)
                continue
            # Check if this is the heading line
            if "INSTRUCTIONS TO ANSWERING PARTY" in stripped.upper():
                _add_paragraph(
                    doc, stripped, bold=True, underline=True, centered=True,
                )
            else:
                _add_paragraph(doc, stripped)

    def _insert_definitions(self, doc, definitions_block: str):
        """Insert the definitions block."""
        lines = definitions_block.strip().split("\n")
        for line in lines:
            stripped = line.strip()
            if not stripped:
                _add_empty_line(doc)
                continue
            # Check if this is the heading line
            if stripped.upper() in ("DEFINITIONS", "DEFINED TERMS"):
                _add_paragraph(
                    doc, stripped, bold=True, underline=True, centered=True,
                )
            else:
                _add_paragraph(doc, stripped)

    def _insert_requests(self, doc, ds: DiscoverySet):
        """Insert numbered discovery requests with inline definitions."""
        for req in ds.requests:
            # Request header (bold + underline)
            header = ds.discovery_type.request_header_template.format(
                number=req.number
            ) + ":"
            _add_paragraph(doc, header, bold=True, underline=True)
            _add_empty_line(doc)

            # Request body (indented via tab)
            _add_paragraph(doc, f"\t{req.text}")

            # Inline definitions
            for defn in req.definitions:
                _add_empty_line(doc)
                _add_paragraph(doc, f"\t{defn}")

            # Visual separator between requests
            _add_empty_line(doc)

    def _insert_signature_block(self, doc, attorney_name: str, firm_name: str,
                                date_str: str):
        """Insert the signature block with firm name, attorney, and date."""
        # /// filler lines
        for _ in range(3):
            _add_paragraph(doc, "///", centered=True)

        _add_empty_line(doc)
        _add_paragraph(doc, f"Dated: {date_str}")
        _add_empty_line(doc)

        # Firm name and attorney on the right side
        _add_paragraph(doc, firm_name, bold=True)
        _add_empty_line(doc)
        _add_empty_line(doc)
        _add_paragraph(doc, "____________________________")
        _add_paragraph(doc, attorney_name)
        if attorney_name:
            _add_paragraph(doc, f"Attorneys for {firm_name}")

    def _insert_declaration(self, doc, ds: DiscoverySet, attorney_name: str,
                            firm_name: str, date_str: str):
        """Insert the CCP declaration of necessity (page break before)."""
        decl_text = generate_declaration(ds, attorney_name, firm_name)
        if not decl_text:
            return

        # Replace the {date} placeholder with the actual date
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

            # Title line (first non-empty line — bold + underline + centered)
            if i == 0 and stripped.startswith("DECLARATION"):
                _add_paragraph(
                    doc, stripped, bold=True, underline=True, centered=True,
                )
            elif stripped.startswith("____"):
                # Signature line
                _add_paragraph(doc, stripped)
            else:
                _add_paragraph(doc, stripped)

    @staticmethod
    def _build_preamble(ds: DiscoverySet) -> str:
        """Return the statutory preamble text for the discovery type.

        Each type cites the relevant CCP section and identifies
        propounding/responding parties.
        """
        prop_name = ds.propounding_party.name
        resp_name = ds.directed_to.name
        resp_role = ds.directed_to.role_label

        if ds.discovery_type == DiscoveryType.SI:
            return (
                f"Pursuant to CCP \u00a72030.030, {prop_name} hereby propounds "
                f"the following Special Interrogatories to {resp_role}, "
                f"{resp_name}, and requests that they be answered under oath "
                f"within thirty (30) days of service."
            )
        elif ds.discovery_type == DiscoveryType.RPD:
            return (
                f"Demand is hereby made by {prop_name} pursuant to CCP "
                f"\u00a72031.010 et seq. that {resp_role}, {resp_name}, produce "
                f"and permit inspection and copying of the following documents "
                f"and tangible things within thirty (30) days of service."
            )
        elif ds.discovery_type == DiscoveryType.RFA:
            return (
                f"Pursuant to CCP \u00a72033.010, {prop_name} hereby requests "
                f"that {resp_role}, {resp_name}, admit the truth of the "
                f"following matters within thirty (30) days of service."
            )
        else:
            return ""
