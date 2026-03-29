"""Combine all .msg files in a folder into a single PDF."""

import os
import glob
import extract_msg
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, HRFlowable
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_LEFT
import html
import re
from datetime import datetime


def clean_text(text):
    """Clean text for reportlab - escape HTML entities and handle encoding."""
    if not text:
        return ""
    text = html.escape(str(text))
    # Replace newlines with <br/> for reportlab
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\n", "<br/>")
    # Remove any control characters
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    return text


def parse_msg_file(filepath):
    """Extract metadata and body from a .msg file."""
    try:
        msg = extract_msg.Message(filepath)
        data = {
            "filename": os.path.basename(filepath),
            "subject": msg.subject or "(no subject)",
            "sender": msg.sender or "Unknown",
            "to": msg.to or "Unknown",
            "cc": msg.cc or "",
            "date": str(msg.date) if msg.date else "Unknown",
            "body": msg.body or "(no body text)",
            "attachments": [att.longFilename or att.shortFilename or "unnamed" for att in msg.attachments] if msg.attachments else [],
        }
        msg.close()
        return data
    except Exception as e:
        return {
            "filename": os.path.basename(filepath),
            "subject": f"ERROR reading file: {e}",
            "sender": "", "to": "", "cc": "", "date": "",
            "body": "", "attachments": [],
        }


def build_pdf(emails, output_path):
    """Build a PDF from a list of parsed email dicts."""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "EmailTitle", parent=styles["Heading2"],
        fontSize=13, spaceAfter=4, textColor=HexColor("#1a1a1a"),
    )
    meta_style = ParagraphStyle(
        "EmailMeta", parent=styles["Normal"],
        fontSize=9, spaceAfter=2, textColor=HexColor("#555555"),
        leftIndent=10,
    )
    body_style = ParagraphStyle(
        "EmailBody", parent=styles["Normal"],
        fontSize=10, spaceAfter=6, leading=14,
    )
    toc_style = ParagraphStyle(
        "TOCEntry", parent=styles["Normal"],
        fontSize=9, spaceAfter=2,
    )

    story = []

    # Title page / TOC
    story.append(Paragraph("Combined Email Evidence", styles["Title"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Source: EVIDENCE folder — {len(emails)} emails", styles["Normal"]))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 20))
    story.append(Paragraph("<b>Table of Contents</b>", styles["Heading3"]))
    story.append(Spacer(1, 6))

    for i, email in enumerate(emails, 1):
        subj = clean_text(email["subject"])
        date = clean_text(email["date"])
        story.append(Paragraph(f"{i}. {subj} — {date}", toc_style))

    story.append(PageBreak())

    # Each email
    for i, email in enumerate(emails, 1):
        story.append(Paragraph(f"{i}. {clean_text(email['subject'])}", title_style))
        story.append(HRFlowable(width="100%", thickness=1, color=HexColor("#cccccc")))
        story.append(Spacer(1, 4))

        story.append(Paragraph(f"<b>From:</b> {clean_text(email['sender'])}", meta_style))
        story.append(Paragraph(f"<b>To:</b> {clean_text(email['to'])}", meta_style))
        if email["cc"]:
            story.append(Paragraph(f"<b>CC:</b> {clean_text(email['cc'])}", meta_style))
        story.append(Paragraph(f"<b>Date:</b> {clean_text(email['date'])}", meta_style))

        if email["attachments"]:
            att_list = ", ".join(clean_text(a) for a in email["attachments"])
            story.append(Paragraph(f"<b>Attachments:</b> {att_list}", meta_style))

        story.append(Spacer(1, 10))

        # Body - split into paragraphs to avoid giant blocks
        body = clean_text(email["body"])
        story.append(Paragraph(body, body_style))

        if i < len(emails):
            story.append(PageBreak())

    doc.build(story)
    print(f"PDF created: {output_path}")


def main():
    folder = r"Z:\Shared\Current Clients\3800- NATIONWIDE\230 - Bhatti\EVIDENCE"
    output = os.path.join(folder, "Combined_Email_Evidence.pdf")

    msg_files = sorted(glob.glob(os.path.join(folder, "*.msg")))
    print(f"Found {len(msg_files)} .msg files")

    emails = []
    for f in msg_files:
        print(f"  Reading: {os.path.basename(f)}")
        emails.append(parse_msg_file(f))

    # Sort by date if possible
    def sort_key(e):
        try:
            return str(e["date"])
        except:
            return ""
    emails.sort(key=sort_key)

    build_pdf(emails, output)
    print("Done!")


if __name__ == "__main__":
    main()
