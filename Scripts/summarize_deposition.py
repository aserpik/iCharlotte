"""
Summarize Deposition Agent for iCharlotte

Summarizes deposition transcripts using LLM with multi-pass processing.

Features:
- Three-pass processing: Extraction -> Narrative Summary -> Cross-Check
- Dynamic topic count based on transcript length
- Automatic deponent name/date extraction from content
- Impeachment material flagging
- Exhibit reference tracking
- Multi-provider LLM fallback (Gemini, Claude, OpenAI)
- Memory monitoring for large transcripts
- Structured progress reporting for UI integration
"""

import os
import sys
import re
import subprocess
import datetime
from docx import Document
from docx.shared import Pt, Inches

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import shared infrastructure
from icharlotte_core.document_processor import DocumentProcessor, OCRConfig
from icharlotte_core.agent_logger import AgentLogger, create_legacy_log_event
from icharlotte_core.llm_config import LLMCaller, LLMConfig
from icharlotte_core.memory_monitor import MemoryMonitor, track_memory
from icharlotte_core.exceptions import (
    ExtractionError, LLMError, PassFailedError,
    SummaryPassError, CrossCheckPassError, MemoryLimitError
)

# Import Case Data Manager
try:
    from case_data_manager import CaseDataManager
except ImportError:
    sys.path.append(os.path.join(os.getcwd(), 'Scripts'))
    from case_data_manager import CaseDataManager


# =============================================================================
# Configuration
# =============================================================================

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
EXTRACTION_PROMPT_FILE = os.path.join(SCRIPTS_DIR, "DEPOSITION_EXTRACTION_PROMPT.txt")
NARRATIVE_PROMPT_FILE = os.path.join(SCRIPTS_DIR, "SUMMARIZE_DEPOSITION_PROMPT.txt")
CROSS_CHECK_PROMPT_FILE = os.path.join(SCRIPTS_DIR, "DEPOSITION_CROSS_CHECK_PROMPT.txt")

# Legacy log file for backward compatibility
LEGACY_LOG_FILE = r"C:\GeminiTerminal\Summarize_Deposition_activity.log"


# =============================================================================
# Deponent Information Extraction
# =============================================================================

class DeponentExtractor:
    """Extracts deponent information from transcript text."""

    # Common deposition header patterns
    DEPONENT_PATTERNS = [
        # "DEPOSITION OF JOHN SMITH" or "DEPOSITION OF JOHN SMITH, M.D."
        r"DEPOSITION\s+OF\s+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+(?:,?\s*(?:M\.?D\.?|D\.?O\.?|Ph\.?D\.?|D\.?C\.?))?)",
        # "Deponent: John Smith"
        r"[Dd]eponent[:\s]+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+)",
        # "THE VIDEOTAPED DEPOSITION OF JOHN SMITH"
        r"(?:VIDEOTAPED\s+)?DEPOSITION\s+OF\s+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+)",
        # "Examination of John Smith"
        r"[Ee]xamination\s+of\s+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+)",
        # "WITNESS: JOHN SMITH" (in caption)
        r"WITNESS[:\s]+([A-Z][A-Za-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][A-Za-z]+)+)",
    ]

    # Date patterns
    DATE_PATTERNS = [
        # "January 15, 2024" or "January 15th, 2024"
        r"([A-Z][a-z]+\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})",
        # "01/15/2024" or "1/15/2024"
        r"(\d{1,2}/\d{1,2}/\d{4})",
        # "2024-01-15"
        r"(\d{4}-\d{2}-\d{2})",
        # "15 January 2024"
        r"(\d{1,2}\s+[A-Z][a-z]+\s+\d{4})",
    ]

    # Doctor patterns
    DOCTOR_PATTERNS = [
        r"(?:Dr\.?|Doctor)\s+([A-Z][A-Za-z]+)",
        r"([A-Z][A-Za-z]+),?\s+(?:M\.?D\.?|D\.?O\.?|Ph\.?D\.?)",
    ]

    @classmethod
    def extract_deponent_name(cls, text: str, first_n_chars: int = 5000) -> str:
        """
        Extract deponent name from transcript text.

        Args:
            text: Full transcript text.
            first_n_chars: Number of characters to search (deponent info usually at start).

        Returns:
            Deponent name or empty string if not found.
        """
        search_text = text[:first_n_chars]

        for pattern in cls.DEPONENT_PATTERNS:
            match = re.search(pattern, search_text, re.IGNORECASE)
            if match:
                name = match.group(1).strip()
                # Clean up the name
                name = re.sub(r'\s+', ' ', name)
                # Remove trailing punctuation
                name = name.rstrip('.,;:')
                return name

        return ""

    @classmethod
    def extract_deposition_date(cls, text: str, first_n_chars: int = 3000) -> str:
        """
        Extract deposition date from transcript text.

        Args:
            text: Full transcript text.
            first_n_chars: Number of characters to search.

        Returns:
            Date string or empty string if not found.
        """
        search_text = text[:first_n_chars]

        # Look for date near "taken on" or "held on" or just in the header
        date_context_patterns = [
            r"(?:taken|held|conducted)\s+(?:on\s+)?(" + "|".join(cls.DATE_PATTERNS) + ")",
            r"(?:Date|DATE)[:\s]+(" + "|".join(cls.DATE_PATTERNS) + ")",
        ]

        for pattern in date_context_patterns:
            match = re.search(pattern, search_text, re.IGNORECASE)
            if match:
                return match.group(1).strip()

        # Fallback: just find any date in the header area
        for pattern in cls.DATE_PATTERNS:
            match = re.search(pattern, search_text)
            if match:
                return match.group(1).strip()

        return ""

    @classmethod
    def detect_deponent_type(cls, text: str, name: str = "") -> str:
        """
        Detect if deponent is plaintiff, defendant, doctor, expert, etc.

        Args:
            text: Transcript text (first few thousand chars).
            name: Deponent name if already extracted.

        Returns:
            Deponent type string.
        """
        search_text = text[:10000].upper()
        name_upper = name.upper() if name else ""

        # Check for doctor
        if name:
            for pattern in cls.DOCTOR_PATTERNS:
                if re.search(pattern, text[:5000], re.IGNORECASE):
                    return "Treating Physician"

        if "M.D." in name_upper or "D.O." in name_upper or "PH.D." in name_upper:
            return "Expert/Physician"

        # Check context clues
        if "PLAINTIFF" in search_text and name_upper in search_text:
            # Check if this name appears near "plaintiff"
            if re.search(rf"PLAINTIFF\s*{re.escape(name_upper)}", search_text):
                return "Plaintiff"

        if "DEFENDANT" in search_text and name_upper in search_text:
            if re.search(rf"DEFENDANT\s*{re.escape(name_upper)}", search_text):
                return "Defendant"

        if "EXPERT" in search_text:
            return "Expert Witness"

        return "Witness"

    @classmethod
    def estimate_page_count(cls, text: str, chars_per_page: int = 3000) -> int:
        """Estimate page count from text length."""
        return max(1, len(text) // chars_per_page)

    @classmethod
    def calculate_topic_count(cls, text: str) -> int:
        """
        Calculate appropriate number of topics based on transcript length.

        Returns:
            Recommended number of topics (5-20).
        """
        pages = cls.estimate_page_count(text)

        if pages < 50:
            return min(8, max(5, pages // 6))
        elif pages < 150:
            return min(12, max(8, pages // 12))
        else:
            return min(20, max(12, pages // 15))


# =============================================================================
# Exhibit Extraction
# =============================================================================

class ExhibitExtractor:
    """Extracts exhibit references from deposition transcripts."""

    # Exhibit patterns
    EXHIBIT_PATTERNS = [
        # "Exhibit A" or "Exhibit 1" or "Exhibit A-1"
        r"[Ee]xhibit\s+([A-Z](?:-?\d+)?|\d+(?:-?[A-Z])?)",
        # "marked as Exhibit A"
        r"marked\s+(?:as\s+)?[Ee]xhibit\s+([A-Z](?:-?\d+)?|\d+)",
        # "Deposition Exhibit 1"
        r"[Dd]eposition\s+[Ee]xhibit\s+(\d+|[A-Z])",
        # "Plaintiff's Exhibit 1" or "Defendant's Exhibit A"
        r"(?:Plaintiff|Defendant)'s\s+[Ee]xhibit\s+([A-Z](?:-?\d+)?|\d+)",
    ]

    @classmethod
    def extract_exhibits(cls, text: str) -> list:
        """
        Extract all exhibit references from transcript.

        Args:
            text: Full transcript text.

        Returns:
            List of unique exhibit designations.
        """
        exhibits = set()

        for pattern in cls.EXHIBIT_PATTERNS:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                # Normalize exhibit designation
                exhibit = match.upper().strip()
                exhibits.add(exhibit)

        # Sort exhibits (numbers first, then letters)
        def sort_key(e):
            if e.isdigit():
                return (0, int(e), "")
            elif e[0].isdigit():
                return (0, int(re.match(r'\d+', e).group()), e)
            else:
                return (1, 0, e)

        return sorted(list(exhibits), key=sort_key)

    @classmethod
    def get_exhibit_context(cls, text: str, exhibit: str, context_chars: int = 500) -> str:
        """
        Get context around an exhibit reference.

        Args:
            text: Full transcript text.
            exhibit: Exhibit designation to find.
            context_chars: Characters of context to extract.

        Returns:
            Context string around exhibit reference.
        """
        pattern = rf"[Ee]xhibit\s+{re.escape(exhibit)}"
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            start = max(0, match.start() - context_chars // 2)
            end = min(len(text), match.end() + context_chars // 2)
            return text[start:end]

        return ""


# =============================================================================
# Impeachment Detection
# =============================================================================

class ImpeachmentDetector:
    """Detects potential impeachment material in deposition testimony."""

    # Patterns that suggest prior inconsistent statements
    IMPEACHMENT_PATTERNS = [
        # Prior testimony references
        (r"(?:prior|previous|earlier)\s+(?:testimony|deposition|statement)", "prior_testimony"),
        (r"(?:did you|didn't you)\s+(?:testify|state|say)\s+(?:before|previously|earlier)", "prior_statement"),
        (r"(?:at your|in your)\s+(?:earlier|prior|previous)\s+deposition", "prior_deposition"),

        # Contradiction indicators
        (r"(?:isn't it true|isn't that correct)\s+that", "contradiction_setup"),
        (r"(?:but|however)\s+(?:earlier|before|previously)\s+you\s+(?:said|stated|testified)", "contradiction"),
        (r"(?:that's not what you|that contradicts what you)\s+(?:said|stated|testified)", "direct_contradiction"),

        # Evasion indicators
        (r"I\s+(?:don't|do not)\s+(?:recall|remember)", "memory_failure"),
        (r"I\s+(?:can't|cannot)\s+(?:recall|remember)", "memory_failure"),
        (r"(?:I'm not sure|I'm unsure|I don't know)\s+(?:if|whether|about)", "uncertainty"),

        # Document contradictions
        (r"(?:this document|this exhibit)\s+(?:shows|indicates|says)\s+(?:something different|otherwise)", "document_contradiction"),
        (r"(?:according to|based on)\s+(?:this|the)\s+(?:document|exhibit|record)", "document_reference"),
    ]

    @classmethod
    def detect_impeachment_material(cls, text: str) -> list:
        """
        Detect potential impeachment material in transcript.

        Args:
            text: Full transcript text.

        Returns:
            List of dicts with impeachment findings.
        """
        findings = []

        for pattern, category in cls.IMPEACHMENT_PATTERNS:
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            for match in matches[:5]:  # Limit to 5 per category
                start = max(0, match.start() - 200)
                end = min(len(text), match.end() + 200)
                context = text[start:end]

                findings.append({
                    "category": category,
                    "match": match.group(),
                    "context": context,
                    "position": match.start()
                })

        # Sort by position in document
        findings.sort(key=lambda x: x["position"])

        return findings

    @classmethod
    def summarize_impeachment(cls, findings: list) -> str:
        """
        Create a summary of impeachment findings.

        Args:
            findings: List of impeachment findings from detect_impeachment_material.

        Returns:
            Formatted summary string.
        """
        if not findings:
            return ""

        # Group by category
        categories = {}
        for finding in findings:
            cat = finding["category"]
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(finding)

        lines = ["## Potential Impeachment Material\n"]

        category_names = {
            "prior_testimony": "Prior Testimony References",
            "prior_statement": "Prior Statement References",
            "prior_deposition": "Prior Deposition References",
            "contradiction_setup": "Contradiction Setups",
            "contradiction": "Contradictions",
            "direct_contradiction": "Direct Contradictions",
            "memory_failure": "Memory Failures",
            "uncertainty": "Uncertainty/Evasion",
            "document_contradiction": "Document Contradictions",
            "document_reference": "Document References",
        }

        for cat, cat_findings in categories.items():
            cat_name = category_names.get(cat, cat.replace("_", " ").title())
            lines.append(f"### {cat_name} ({len(cat_findings)} instances)")

            for i, finding in enumerate(cat_findings[:3], 1):  # Show top 3 per category
                context = finding["context"].replace("\n", " ").strip()
                if len(context) > 300:
                    context = context[:300] + "..."
                lines.append(f"- Instance {i}: \"...{context}...\"")

            lines.append("")

        return "\n".join(lines)


# =============================================================================
# Document Output Functions
# =============================================================================

def add_markdown_to_doc(doc, content):
    """Parses basic Markdown and adds it to the docx Document with specific formatting."""
    lines = content.split('\n')

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Headings (## or #)
        if stripped.startswith('#'):
            text = stripped.lstrip('#').strip()
            if not text.endswith('.'):
                text += "."
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.first_line_indent = Inches(0.5)
            run = p.add_run(text)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            continue

        # List items
        if stripped.startswith('* ') or stripped.startswith('- '):
            text = stripped[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.25)

            run = p.add_run("\t")
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

            # Support bold parsing within list items
            parts = re.split(r'(\*\*.*?\*\*)', text)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    r = p.add_run(part[2:-2])
                    r.bold = True
                else:
                    r = p.add_run(part)
                r.font.name = 'Times New Roman'
                r.font.size = Pt(12)
            continue

        # Normal text (with bold support)
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.first_line_indent = Inches(0.5)

        parts = re.split(r'(\*\*.*?\*\*)', stripped)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part[2:-2])
                run.bold = True
            else:
                run = p.add_run(part)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)


def save_to_docx(content: str, output_path: str, deponent_name: str,
                 deposition_date: str, logger: AgentLogger) -> bool:
    """
    Saves the content to a DOCX file with deposition formatting.

    Args:
        content: Summary content to save.
        output_path: Path to output file (will append if exists).
        deponent_name: Name of the deponent.
        deposition_date: Date of the deposition.
        logger: AgentLogger instance.

    Returns:
        True if successful.
    """
    base_name, ext = os.path.splitext(output_path)
    counter = 1
    current_output_path = output_path

    while True:
        try:
            if os.path.exists(current_output_path):
                try:
                    doc = Document(current_output_path)
                    doc.add_page_break()
                    logger.info(f"Appending to existing file: {current_output_path}")
                except Exception:
                    raise PermissionError("Cannot open existing file.")
            else:
                doc = Document()
                logger.info(f"Creating new file: {current_output_path}")

            # Apply styles
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)
            style.paragraph_format.line_spacing = 1.0

            # Count existing depositions for numbering
            next_num = 1
            for para in doc.paragraphs:
                if re.match(r'^\d+\.\tDeposition of', para.text):
                    next_num += 1

            # Title paragraph with numbering
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.left_indent = Inches(1.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(1.5))

            run_num = p.add_run(f"{next_num}.\t")
            run_num.bold = True
            run_num.font.name = 'Times New Roman'
            run_num.font.size = Pt(12)

            # Format title with deponent name
            title_name = deponent_name if deponent_name else "Unknown Deponent"
            run_title = p.add_run(f"Deposition of {title_name}")
            run_title.bold = True
            run_title.underline = True
            run_title.font.name = 'Times New Roman'
            run_title.font.size = Pt(12)

            # Add date if available
            if deposition_date:
                p2 = doc.add_paragraph()
                p2.paragraph_format.line_spacing = 1.0
                run_date = p2.add_run(f"Date: {deposition_date}")
                run_date.font.name = 'Times New Roman'
                run_date.font.size = Pt(12)
                run_date.italic = True

            # Add summary content
            add_markdown_to_doc(doc, content)

            doc.save(current_output_path)
            logger.output_file(current_output_path)
            return True

        except (PermissionError, IOError) as e:
            logger.warning(f"File locked: {current_output_path}. Trying next version.")
            counter += 1
            current_output_path = f"{base_name} v.{counter}{ext}"

            if counter > 10:
                logger.error("Failed to save after 10 attempts.")
                return False

        except Exception as e:
            logger.error(f"Error saving to DOCX: {e}")
            return False


def get_output_directory(input_path: str) -> str:
    """Determine the output directory based on input path."""
    parts = input_path.split(os.sep)
    output_dir = None
    case_root_parts = None

    # Priority 1: Find folder starting with exactly 3 digits (Case Folder)
    for i in range(len(parts) - 1, -1, -1):
        if re.match(r'^\d{3}(\D|$)', parts[i]):
            case_root_parts = parts[:i+1]
            break

    # Priority 2: Standard "Current Clients" structure
    if not case_root_parts:
        for i, part in enumerate(parts):
            if part.lower() == "current clients":
                if i + 2 < len(parts):
                    case_root_parts = parts[:i+3]
                break

    if case_root_parts:
        output_dir = os.sep.join(case_root_parts + ["NOTES", "AI OUTPUT"])

    if not output_dir:
        # Fallback 1: NOTES already in path
        for i in range(len(parts) - 1, -1, -1):
            if parts[i].upper() == "NOTES":
                output_dir = os.path.join(os.sep.join(parts[:i+1]), "AI OUTPUT")
                break

    if not output_dir:
        # Fallback 2: Sibling NOTES folder
        input_dir = os.path.dirname(input_path)
        parent_dir = os.path.dirname(input_dir)
        output_dir = os.path.join(parent_dir, "NOTES", "AI OUTPUT")

    return output_dir


# =============================================================================
# Main Processing Functions
# =============================================================================

def build_narrative_prompt(base_prompt: str, topic_count: int, deponent_name: str,
                          deponent_type: str) -> str:
    """
    Build the narrative summary prompt with dynamic topic count.

    Args:
        base_prompt: Base prompt from file.
        topic_count: Number of topics to request.
        deponent_name: Name of the deponent.
        deponent_type: Type of deponent (Plaintiff, Doctor, etc.).

    Returns:
        Modified prompt string.
    """
    # Replace the fixed "10 key topics" with dynamic count
    prompt = re.sub(
        r"identify\s+\d+\s+key topics",
        f"identify {topic_count} key topics",
        base_prompt,
        flags=re.IGNORECASE
    )

    prompt = re.sub(
        r"using the\s+\d+\s+key topics",
        f"using the {topic_count} key topics",
        prompt,
        flags=re.IGNORECASE
    )

    # Add deponent info if available
    if deponent_name:
        prompt += f"\n\nNote: The deponent is {deponent_name}"
        if deponent_type:
            prompt += f", who is the {deponent_type}"
        prompt += "."

    return prompt


def process_document(input_path: str, logger: AgentLogger) -> bool:
    """
    Process a single deposition transcript through the 3-pass pipeline.

    Pipeline:
    1. Extract text
    2. Pass 1: Structured extraction (deponent, topics, key Q&A, exhibits)
    3. Pass 2: Narrative summary (dynamic topic count)
    4. Pass 3: Cross-check with impeachment detection
    5. Save to DOCX

    Args:
        input_path: Path to the deposition file.
        logger: AgentLogger instance.

    Returns:
        True if successful.
    """
    # Initialize components
    memory_monitor = MemoryMonitor(warn_threshold_mb=1500, abort_threshold_mb=2000, logger=logger.info)
    llm_caller = LLMCaller(logger=logger)

    # Load prompts
    extraction_prompt = None
    narrative_prompt = None
    cross_check_prompt = None

    try:
        if os.path.exists(EXTRACTION_PROMPT_FILE):
            with open(EXTRACTION_PROMPT_FILE, "r", encoding="utf-8") as f:
                extraction_prompt = f.read()

        with open(NARRATIVE_PROMPT_FILE, "r", encoding="utf-8") as f:
            narrative_prompt = f.read()

        if os.path.exists(CROSS_CHECK_PROMPT_FILE):
            with open(CROSS_CHECK_PROMPT_FILE, "r", encoding="utf-8") as f:
                cross_check_prompt = f.read()

    except Exception as e:
        logger.error(f"Error reading prompt files: {e}")
        return False

    # Determine number of passes
    total_passes = 2  # Extraction + Narrative minimum
    if extraction_prompt:
        total_passes += 1  # Add extraction pass
    if cross_check_prompt:
        total_passes += 1  # Add cross-check pass

    pass_number = 0

    # ==========================================================================
    # Pass 1: Text Extraction
    # ==========================================================================
    pass_number += 1
    logger.pass_start("Text Extraction", pass_number, total_passes)

    try:
        with memory_monitor.track_operation("Text Extraction"):
            processor = DocumentProcessor(
                ocr_config=OCRConfig(adaptive=True),
                logger=logger
            )
            result = processor.extract_with_dynamic_ocr(input_path)

            if not result.success:
                raise ExtractionError(f"Failed to extract text: {result.error}", file_path=input_path)

            text = result.text
            logger.info(f"Extracted {result.char_count} chars from {result.page_count} pages")
            logger.info(f"OCR used on {len(result.ocr_pages)} pages ({result.ocr_percentage:.1f}%)")

    except MemoryLimitError as e:
        logger.pass_failed("Text Extraction", str(e), recoverable=False)
        return False
    except Exception as e:
        logger.pass_failed("Text Extraction", str(e), recoverable=False)
        return False

    logger.pass_complete("Text Extraction", success=True)

    # ==========================================================================
    # Extract Deponent Information from Content
    # ==========================================================================
    deponent_name = DeponentExtractor.extract_deponent_name(text)
    deposition_date = DeponentExtractor.extract_deposition_date(text)
    deponent_type = DeponentExtractor.detect_deponent_type(text, deponent_name)
    topic_count = DeponentExtractor.calculate_topic_count(text)

    if deponent_name:
        logger.info(f"Extracted deponent: {deponent_name} ({deponent_type})")
    else:
        # Fallback: extract from filename
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        for skip in ["deposition", "depo", "transcript", ".pdf", ".docx"]:
            base_name = base_name.replace(skip, "").replace(skip.upper(), "")
        deponent_name = base_name.strip(" _-")
        logger.info(f"Using filename for deponent: {deponent_name}")

    if deposition_date:
        logger.info(f"Extracted date: {deposition_date}")

    logger.info(f"Recommended topic count: {topic_count} (based on ~{DeponentExtractor.estimate_page_count(text)} pages)")

    # Extract exhibits
    exhibits = ExhibitExtractor.extract_exhibits(text)
    if exhibits:
        logger.info(f"Found {len(exhibits)} exhibits: {', '.join(exhibits[:10])}{'...' if len(exhibits) > 10 else ''}")

    # Detect impeachment material
    impeachment_findings = ImpeachmentDetector.detect_impeachment_material(text)
    if impeachment_findings:
        logger.info(f"Found {len(impeachment_findings)} potential impeachment items")

    # ==========================================================================
    # Pass 2: Structured Extraction (if prompt available)
    # ==========================================================================
    extraction_result = None

    if extraction_prompt:
        pass_number += 1
        logger.pass_start("Structured Extraction", pass_number, total_passes)

        try:
            with memory_monitor.track_operation("Structured Extraction"):
                extraction_result = llm_caller.call(extraction_prompt, text, task_type="extraction")

                if not extraction_result:
                    logger.warning("Extraction pass returned empty. Continuing without extraction.")
                else:
                    logger.info(f"Extraction complete: {len(extraction_result)} chars")

        except Exception as e:
            logger.pass_failed("Structured Extraction", str(e), recoverable=True)
            logger.warning("Continuing with narrative pass only")

        logger.pass_complete("Structured Extraction", success=bool(extraction_result))

    # ==========================================================================
    # Pass 3: Narrative Summary
    # ==========================================================================
    pass_number += 1
    logger.pass_start("Narrative Summary", pass_number, total_passes)

    try:
        with memory_monitor.track_operation("Narrative Summary"):
            # Build prompt with dynamic topic count
            modified_prompt = build_narrative_prompt(
                narrative_prompt,
                topic_count,
                deponent_name,
                deponent_type
            )

            summary = llm_caller.call(modified_prompt, text, task_type="summary")

            if not summary:
                raise SummaryPassError("LLM returned empty response for narrative summary")

            logger.info(f"Generated narrative summary: {len(summary)} chars")

    except Exception as e:
        logger.pass_failed("Narrative Summary", str(e), recoverable=True)
        return False

    logger.pass_complete("Narrative Summary", success=True)

    # ==========================================================================
    # Pass 4: Cross-Check with Impeachment Enhancement (if prompt available)
    # ==========================================================================
    if cross_check_prompt and extraction_result:
        pass_number += 1
        logger.pass_start("Cross-Check", pass_number, total_passes)

        try:
            with memory_monitor.track_operation("Cross-Check"):
                # Format the cross-check prompt
                formatted_prompt = cross_check_prompt.replace(
                    "{extraction}", extraction_result
                ).replace(
                    "{summary}", summary
                ).replace(
                    "{original}", text[:75000]  # Limit original text
                )

                verified_summary = llm_caller.call(formatted_prompt, "", task_type="cross_check")

                if verified_summary and len(verified_summary) > len(summary) * 0.8:
                    summary = verified_summary
                    logger.info("Cross-check completed with enhancements")
                else:
                    logger.warning("Cross-check returned shorter/empty result. Using original summary.")

        except Exception as e:
            logger.pass_failed("Cross-Check", str(e), recoverable=True)
            logger.warning("Continuing with unverified summary")

        logger.pass_complete("Cross-Check", success=True)

    # ==========================================================================
    # Add Exhibit List and Impeachment Summary
    # ==========================================================================

    # Add exhibit list if we have exhibits
    if exhibits:
        exhibit_section = "\n\n## Exhibits Referenced\n"
        for exhibit in exhibits:
            context = ExhibitExtractor.get_exhibit_context(text, exhibit, 200)
            # Extract a brief description from context
            brief = context[:100].replace("\n", " ").strip() + "..."
            exhibit_section += f"- **Exhibit {exhibit}**: {brief}\n"
        summary += exhibit_section

    # Add impeachment summary if we have findings
    if impeachment_findings:
        impeachment_summary = ImpeachmentDetector.summarize_impeachment(impeachment_findings)
        if impeachment_summary:
            summary += "\n\n" + impeachment_summary

    # ==========================================================================
    # Save Output
    # ==========================================================================
    output_dir = get_output_directory(input_path)

    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            logger.info(f"Created directory: {output_dir}")
        except Exception as e:
            logger.error(f"Error creating directory: {e}")
            return False

    output_file = os.path.join(output_dir, "Deposition_summaries.docx")

    if not save_to_docx(summary, output_file, deponent_name, deposition_date, logger):
        return False

    # ==========================================================================
    # Save to Case Data
    # ==========================================================================
    try:
        data_manager = CaseDataManager()
        file_num_match = re.search(r"(\d{4}\.\d{3})", input_path)

        if file_num_match:
            file_num = file_num_match.group(1)
            clean_name = re.sub(r"[^a-zA-Z0-9_]", "_", deponent_name.lower()) if deponent_name else "unknown"
            var_key = f"depo_summary_{clean_name}"

            logger.info(f"Saving to case data: {var_key} for {file_num}")
            data_manager.save_variable(
                file_num,
                var_key,
                summary,
                source="deposition_agent",
                extra_tags=["Deposition", deponent_type] if deponent_type else ["Deposition"]
            )

            # Also save structured data if available
            if extraction_result:
                data_manager.save_variable(
                    file_num,
                    f"depo_extraction_{clean_name}",
                    extraction_result,
                    source="deposition_agent",
                    extra_tags=["Deposition", "Extraction"]
                )

    except Exception as e:
        logger.warning(f"Could not save to case data: {e}")

    return True


def process_directory(dir_path: str, logger: AgentLogger):
    """Process all deposition documents in a directory."""
    files_to_process = []

    for root, _, files in os.walk(dir_path):
        for file in files:
            if file.lower().endswith(('.pdf', '.docx')):
                # Skip output files
                if "Deposition_summaries" in file:
                    continue
                if "AI_OUTPUT" in file:
                    continue
                files_to_process.append(os.path.join(root, file))

    if not files_to_process:
        logger.info("No suitable deposition files found in directory.")
        return

    logger.info(f"Found {len(files_to_process)} files to process.")

    for file_path in files_to_process:
        logger.info(f"Processing: {file_path}")
        try:
            subprocess.run([sys.executable, sys.argv[0], file_path], check=True)
        except subprocess.CalledProcessError as e:
            logger.error(f"Failed to process {file_path}: {e}")


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        print("Error: No file path provided.", flush=True)
        sys.exit(1)

    # Parse arguments
    raw_args = sys.argv[1:]

    # Handle quoted paths with spaces
    combined_arg = " ".join(raw_args)
    clean_combined = combined_arg.strip().strip('"').strip("'")

    if os.path.exists(clean_combined):
        file_paths = [clean_combined]
    else:
        file_paths = raw_args

    # Dispatcher mode: multiple files
    if len(file_paths) > 1:
        print(f"Detected {len(file_paths)} files. Launching separate agents...", flush=True)
        for path in file_paths:
            try:
                if os.name == 'nt':
                    subprocess.Popen([sys.executable, sys.argv[0], path], creationflags=0x08000000)
                else:
                    subprocess.Popen([sys.executable, sys.argv[0], path])
            except Exception as e:
                print(f"Failed to spawn agent for {path}: {e}", flush=True)
        sys.exit(0)

    # Single file/directory mode
    input_path = file_paths[0].strip().strip('"').strip("'")
    input_path = os.path.abspath(input_path)

    if not os.path.exists(input_path):
        print(f"Error: File not found: {input_path}", flush=True)
        sys.exit(1)

    # Extract file number for logger context
    file_num_match = re.search(r"(\d{4}\.\d{3})", input_path)
    file_number = file_num_match.group(1) if file_num_match else None

    # Initialize logger
    logger = AgentLogger("Deposition", file_number=file_number)

    # Also create legacy log_event for backward compatibility
    log_event = create_legacy_log_event("Deposition", LEGACY_LOG_FILE)

    logger.info(f"Starting deposition agent for: {input_path}")

    # Directory handling
    if os.path.isdir(input_path):
        logger.info(f"Input is a directory. Scanning for files...")
        process_directory(input_path, logger)
        logger.info("Directory processing complete.")
        sys.exit(0)

    # Single file processing
    success = process_document(input_path, logger)

    if success:
        logger.info("Deposition agent finished successfully.")
        sys.exit(0)
    else:
        logger.error("Deposition agent finished with errors.")
        sys.exit(1)


if __name__ == '__main__':
    main()
