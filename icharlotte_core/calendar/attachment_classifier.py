"""
Attachment Classifier for iCharlotte Calendar Agent.

Uses LLM to classify legal document types and extract relevant information
like hearing dates and motion types.
"""

import os
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional, Tuple, Any

from icharlotte_core.ui.logs_tab import LogManager
from icharlotte_core.llm import LLMHandler
from icharlotte_core.config import API_KEYS


class AttachmentClassifier:
    """
    LLM-based classifier for legal document attachments.

    Classifies documents into:
    - correspondence: Letters, emails, general correspondence
    - discovery_request: Interrogatories, RFPs, RFAs, deposition notices
    - discovery_response: Responses to discovery requests
    - motion: Motions filed with the court
    - opposition: Oppositions to motions
    - reply: Reply briefs

    Also extracts:
    - Motion type (MSJ, MSA, demurrer, standard, etc.)
    - Hearing date (if present)
    - Key dates mentioned in the document

    Usage:
        classifier = AttachmentClassifier()
        result = classifier.classify_attachment(file_path)
        # result = {
        #     'doc_type': 'motion',
        #     'motion_type': 'msj',
        #     'hearing_date': datetime(2026, 3, 20),
        #     'dates_found': [...],
        #     'summary': 'Motion for Summary Judgment...'
        # }
    """

    # Document type patterns for fallback detection
    DOC_TYPE_PATTERNS = {
        'deposition_notice': [
            r'deposition\s+notice',
            r'notice\s+of\s+(?:taking\s+)?deposition',
            r'notice\s+of\s+(?:video(?:taped)?|remote)?\s*deposition',
            r'notice\s+to\s+take\s+deposition',
        ],
        # NOTE: discovery_response MUST be checked before discovery_request
        # because response documents often contain the request text they're responding to
        'discovery_response': [
            r'response\s+to\s+(?:form\s+)?interrogator',
            r'response\s+to\s+(?:special\s+)?interrogator',
            r'response\s+to\s+request\s+for\s+production',
            r'response\s+to\s+request\s+for\s+admission',
            r'response\s+to\s+demand',
            r"plaintiff's\s+response\s+to",
            r"defendant's\s+response\s+to",
            r"responses?\s+to\s+(?:form\s+)?interrogator",
            r"responses?\s+to\s+request",
        ],
        'discovery_request': [
            r'form\s+interrogator',
            r'special\s+interrogator',
            r'request\s+for\s+production',
            r'request\s+for\s+admission',
            r'demand\s+for\s+inspection',
            r'set\s+(?:one|two|three|four|five|\d+)',
        ],
        # NOTE: opposition and reply MUST be checked before motion
        # because opposition/reply documents often contain "motion" in their text
        'opposition': [
            r'opposition\s+to',
            r'opposing\s+(?:the\s+)?motion',
            r'in\s+opposition',
            r'memorandum\s+of\s+points\s+and\s+authorities\s+in\s+opposition',
        ],
        'reply': [
            r'reply\s+(?:brief|memorandum|to\s+opposition)',
            r'reply\s+in\s+support',
            r'moving\s+party.{0,20}reply',
        ],
        'motion': [
            r'motion\s+for\s+summary\s+judgment',
            r'motion\s+for\s+summary\s+adjudication',
            r'notice\s+of\s+motion',
            r'motion\s+to\s+compel',
            r'motion\s+to\s+strike',
            r'demurrer',
            r'motion\s+for\s+protective\s+order',
            r'motion\s+in\s+limine',
            r'ex\s+parte\s+application',
        ],
        'correspondence': [
            r'^dear\s+',
            r'sincerely,',
            r'best\s+regards,',
            r'very\s+truly\s+yours,',
        ],
    }

    # Motion type patterns
    # NOTE: Order matters - anti_slapp must be checked before motion_to_strike
    # because anti-SLAPP is a "special motion to strike"
    MOTION_TYPE_PATTERNS = {
        'msj': [r'summary\s+judgment', r'\bmsj\b'],
        'msa': [r'summary\s+adjudication', r'\bmsa\b'],
        'demurrer': [r'\bdemurrer\b'],
        'anti_slapp': [r'anti-?slapp', r'special\s+motion\s+to\s+strike', r'425\.16', r'ccp\s*425\.16'],
        'motion_to_strike': [r'(?<!special\s)motion\s+to\s+strike'],  # Negative lookbehind to skip anti-SLAPP
        'motion_to_compel': [r'motion\s+to\s+compel'],
        'protective_order': [r'protective\s+order'],
        'in_limine': [r'in\s+limine', r'\bmil\b'],
        'ex_parte': [r'ex\s+parte'],
    }

    # Date patterns
    DATE_PATTERNS = [
        # Month DD, YYYY
        r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}',
        # MM/DD/YYYY or MM-DD-YYYY
        r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',
        # Month DD
        r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?(?:,?\s+\d{4})?',
    ]

    def __init__(self):
        """Initialize the classifier."""
        self.log = LogManager()

    def classify_attachment(self, file_path: str, email_body: str = "") -> Dict[str, Any]:
        """
        Classify an attachment and extract relevant information.

        Args:
            file_path: Path to the attachment file
            email_body: Optional email body for additional context

        Returns:
            Classification result dict with keys:
            - doc_type: str (correspondence, discovery_request, discovery_response, motion, opposition, reply)
            - motion_type: str (msj, msa, demurrer, standard, etc.) - only if doc_type is motion/opposition/reply
            - hearing_date: Optional[datetime] - extracted hearing date
            - dates_found: List[datetime] - all dates found in document
            - summary: str - brief summary of the document
            - confidence: float - confidence score (0-1)
        """
        result = {
            'doc_type': 'correspondence',
            'motion_type': None,
            'hearing_date': None,
            'deposition_date': None,
            'deposition_time': None,
            'deponent_name': None,
            'dates_found': [],
            'summary': '',
            'confidence': 0.0,
        }

        try:
            # Extract text from attachment
            text = self._extract_text(file_path)

            if not text:
                self.log.add_log("Calendar", f"Could not extract text from {file_path}")
                return result

            # Try LLM classification first
            if API_KEYS.get("Gemini"):
                llm_result = self._classify_with_llm(text, email_body)
                if llm_result:
                    result.update(llm_result)
                    return result

            # Fallback to pattern-based classification
            pattern_result = self._classify_with_patterns(text)
            result.update(pattern_result)

            # Extract dates
            result['dates_found'] = self._extract_dates(text)

            # Try to find hearing date
            result['hearing_date'] = self._find_hearing_date(text)

            return result

        except Exception as e:
            self.log.add_log("Calendar", f"Classification error: {e}")
            return result

    def _extract_text(self, file_path: str) -> str:
        """
        Extract text content from a file.

        Supports PDF, DOC, DOCX, TXT.
        """
        path = Path(file_path)
        ext = path.suffix.lower()

        try:
            if ext == '.txt':
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()

            elif ext == '.pdf':
                return self._extract_pdf_text(file_path)

            elif ext in ['.doc', '.docx']:
                return self._extract_word_text(file_path)

            else:
                self.log.add_log("Calendar", f"Unsupported file type: {ext}")
                return ""

        except Exception as e:
            self.log.add_log("Calendar", f"Text extraction error: {e}")
            return ""

    def _extract_pdf_text(self, file_path: str) -> str:
        """Extract text from PDF using PyMuPDF or pdfplumber."""
        text = ""

        # Try PyMuPDF first
        try:
            import fitz  # PyMuPDF
            with fitz.open(file_path) as doc:
                for page in doc:
                    text += page.get_text()
            return text
        except ImportError:
            pass
        except Exception as e:
            self.log.add_log("Calendar", f"PyMuPDF error: {e}")

        # Try pdfplumber as fallback
        try:
            import pdfplumber
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
            return text
        except ImportError:
            self.log.add_log("Calendar", "Neither PyMuPDF nor pdfplumber available")
        except Exception as e:
            self.log.add_log("Calendar", f"pdfplumber error: {e}")

        return ""

    def _extract_word_text(self, file_path: str) -> str:
        """Extract text from Word documents."""
        try:
            import docx
            doc = docx.Document(file_path)
            return "\n".join(para.text for para in doc.paragraphs)
        except ImportError:
            self.log.add_log("Calendar", "python-docx not available")
        except Exception as e:
            self.log.add_log("Calendar", f"Word extraction error: {e}")

        return ""

    def _classify_with_llm(self, text: str, email_body: str = "") -> Optional[Dict]:
        """
        Use LLM to classify the document.
        """
        # Truncate text if too long (first 8000 chars should be enough for classification)
        text_sample = text[:8000] if len(text) > 8000 else text

        system_prompt = """You are a legal document classifier. Analyze the provided document and classify it.

Return a JSON object with these fields:
{
    "doc_type": "correspondence" | "deposition_notice" | "discovery_request" | "discovery_response" | "motion" | "opposition" | "reply",
    "motion_type": null | "msj" | "msa" | "demurrer" | "motion_to_strike" | "motion_to_compel" | "anti_slapp" | "standard",
    "hearing_date": null | "YYYY-MM-DD",
    "deposition_date": null | "YYYY-MM-DD",
    "deposition_time": null | "HH:MM" (24-hour format, e.g., "09:00", "14:30"),
    "deponent_name": null | "Name of person being deposed",
    "summary": "Brief 1-2 sentence summary of the document",
    "confidence": 0.0-1.0
}

Classification rules:
- correspondence: Letters, emails, cover letters, transmittals
- deposition_notice: Notice of Deposition, Notice of Taking Deposition (NOT other discovery requests)
- discovery_request: Interrogatories, Requests for Production (RFPs), Requests for Admission (RFAs) (NOT deposition notices)
- discovery_response: Responses to any discovery request
- motion: Any motion filed with the court (includes Notice of Motion, Memorandum of Points and Authorities)
- opposition: Opposition papers to a motion
- reply: Reply brief in support of a motion

IMPORTANT: Deposition notices are NOT discovery_request. They must be classified as "deposition_notice".

For deposition_notice documents:
- Extract deponent_name: The name of the person being deposed (e.g., "John Smith", "Jane Doe, PMK")
- Extract deposition_date: The date the deposition is scheduled for
- Extract deposition_time: The time the deposition is scheduled for in 24-hour format (e.g., "09:00", "10:30", "14:00")

For motion_type, only set if doc_type is motion, opposition, or reply:
- msj: Motion for Summary Judgment
- msa: Motion for Summary Adjudication
- demurrer: Demurrer
- motion_to_strike: Motion to Strike
- motion_to_compel: Motion to Compel
- anti_slapp: Anti-SLAPP / Special Motion to Strike
- standard: Any other motion type

For hearing_date, extract if mentioned (e.g., "Hearing Date: March 20, 2026" or "set for hearing on...")

Return ONLY the JSON object, no other text."""

        user_prompt = f"Document text:\n\n{text_sample}"
        if email_body:
            user_prompt += f"\n\nEmail context:\n{email_body[:1000]}"

        try:
            result = LLMHandler.generate(
                provider="Gemini",
                model="gemini-2.0-flash",
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                file_contents="",
                settings={"temperature": 0.1, "max_tokens": 500, "stream": False}
            )

            # Parse JSON response
            import json

            # Clean up response (remove markdown code blocks if present)
            result = result.strip()
            if result.startswith('```'):
                result = result.split('\n', 1)[1]  # Remove first line
                result = result.rsplit('```', 1)[0]  # Remove last ```

            parsed = json.loads(result)

            # Convert hearing_date string to datetime if present
            if parsed.get('hearing_date'):
                try:
                    parsed['hearing_date'] = datetime.strptime(parsed['hearing_date'], '%Y-%m-%d')
                except:
                    parsed['hearing_date'] = None

            # Convert deposition_date string to datetime if present
            if parsed.get('deposition_date'):
                try:
                    parsed['deposition_date'] = datetime.strptime(parsed['deposition_date'], '%Y-%m-%d')
                except:
                    parsed['deposition_date'] = None

            # Extract all dates from text
            parsed['dates_found'] = self._extract_dates(text)

            return parsed

        except Exception as e:
            self.log.add_log("Calendar", f"LLM classification error: {e}")
            return None

    def _classify_with_patterns(self, text: str) -> Dict:
        """
        Fallback pattern-based classification.
        """
        text_lower = text.lower()
        result = {
            'doc_type': 'correspondence',
            'motion_type': None,
            'confidence': 0.5,
            'summary': '',
        }

        # Check each document type
        for doc_type, patterns in self.DOC_TYPE_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, text_lower, re.IGNORECASE):
                    result['doc_type'] = doc_type
                    result['confidence'] = 0.7
                    break
            if result['doc_type'] != 'correspondence':
                break

        # If it's a motion/opposition/reply, determine motion type
        if result['doc_type'] in ['motion', 'opposition', 'reply']:
            for motion_type, patterns in self.MOTION_TYPE_PATTERNS.items():
                for pattern in patterns:
                    if re.search(pattern, text_lower, re.IGNORECASE):
                        result['motion_type'] = motion_type
                        break
                if result['motion_type']:
                    break

            # Default to standard if no specific type found
            if not result['motion_type']:
                result['motion_type'] = 'standard'

        return result

    def _extract_dates(self, text: str) -> list:
        """
        Extract all dates from text.
        """
        dates = []

        for pattern in self.DATE_PATTERNS:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                parsed = self._parse_date(match)
                if parsed and parsed not in dates:
                    dates.append(parsed)

        return sorted(dates)

    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """
        Parse a date string into a datetime object.
        """
        formats = [
            '%B %d, %Y',      # January 15, 2026
            '%B %d %Y',       # January 15 2026
            '%m/%d/%Y',       # 01/15/2026
            '%m-%d-%Y',       # 01-15-2026
            '%m/%d/%y',       # 01/15/26
            '%m-%d-%y',       # 01-15-26
        ]

        # Clean up the string
        date_str = date_str.strip()
        date_str = re.sub(r'(\d+)(?:st|nd|rd|th)', r'\1', date_str)  # Remove ordinals

        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue

        return None

    def _find_hearing_date(self, text: str) -> Optional[datetime]:
        """
        Find the hearing date specifically mentioned in the document.
        """
        # Look for explicit hearing date mentions
        hearing_patterns = [
            r'hearing\s+(?:date|is\s+set\s+for|on)[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})',
            r'set\s+for\s+hearing\s+on\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})',
            r'hearing[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
            r'(?:will\s+be\s+)?heard\s+on\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})',
        ]

        for pattern in hearing_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                parsed = self._parse_date(match.group(1))
                if parsed:
                    return parsed

        return None

    def classify_from_outlook_attachment(self, attachment) -> Tuple[str, Dict[str, Any]]:
        """
        Classify an Outlook attachment object.

        Saves the attachment to a temp file, classifies it, then cleans up.

        Args:
            attachment: Outlook attachment COM object

        Returns:
            Tuple of (temp_file_path, classification_result)
        """
        temp_path = None

        try:
            # Save attachment to temp file
            filename = attachment.FileName
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"calendar_agent_{filename}")
            attachment.SaveAsFile(temp_path)

            # Classify
            result = self.classify_attachment(temp_path)

            return temp_path, result

        except Exception as e:
            self.log.add_log("Calendar", f"Attachment processing error: {e}")
            return temp_path, {
                'doc_type': 'correspondence',
                'motion_type': None,
                'hearing_date': None,
                'deposition_date': None,
                'deposition_time': None,
                'deponent_name': None,
                'dates_found': [],
                'summary': '',
                'confidence': 0.0,
            }
