"""
Summarize Agent for iCharlotte

Summarizes documents (PDF, DOCX, TXT) using LLM with cross-check verification.

Features:
- Dynamic OCR threshold for text extraction
- Two-pass processing: Summary + Cross-Check
- Multi-provider LLM fallback (Gemini, Claude, OpenAI)
- Memory monitoring for large documents
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
PROMPT_FILE = os.path.join(SCRIPTS_DIR, "SUMMARIZE_PROMPT.txt")
CROSS_CHECK_PROMPT_FILE = os.path.join(SCRIPTS_DIR, "SUMMARIZE_CROSS_CHECK_PROMPT.txt")

# Legacy log file for backward compatibility
LEGACY_LOG_FILE = r"C:\GeminiTerminal\Summarize_activity.log"


# =============================================================================
# Document Output Functions
# =============================================================================

def add_markdown_to_doc(doc, content):
    """Parses basic Markdown and adds it to the docx Document."""
    lines = content.split('\n')
    active_paragraph = None

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Headings: Convert to bold text at start of paragraph, ending with period
        if stripped.startswith('#'):
            text = stripped.lstrip('#').strip()
            if not text.endswith('.'):
                text += "."
            active_paragraph = doc.add_paragraph()
            run = active_paragraph.add_run(text + " ")
            run.bold = True
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

            active_paragraph = None
            continue

        # Normal text (with bold support)
        p = active_paragraph if active_paragraph else doc.add_paragraph()

        parts = re.split(r'(\*\*.*?\*\*)', stripped)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part[2:-2])
                run.bold = True
            else:
                p.add_run(part)

        active_paragraph = None


def save_to_docx(content: str, output_path: str, title_text: str, logger: AgentLogger) -> bool:
    """Saves the content to a DOCX file. Appends if exists. Handles locking."""
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

            # Apply Styles
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)
            style.paragraph_format.line_spacing = 1.0

            # Update Heading styles
            for i in range(1, 10):
                if f'Heading {i}' in doc.styles:
                    h_style = doc.styles[f'Heading {i}']
                    h_style.font.name = 'Times New Roman'
                    h_style.font.size = Pt(12)
                    h_style.paragraph_format.line_spacing = 1.0

            # Title
            p = doc.add_paragraph()
            run = p.add_run(title_text)
            run.bold = True
            run.underline = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

            doc.add_paragraph(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

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

def process_document(input_path: str, logger: AgentLogger) -> bool:
    """
    Process a single document through the summarization pipeline.

    Pipeline:
    1. Extract text (with dynamic OCR)
    2. Pass 1: Initial summary
    3. Pass 2: Cross-check verification
    4. Save to DOCX

    Args:
        input_path: Path to the document.
        logger: AgentLogger instance.

    Returns:
        True if successful.
    """
    # Initialize components
    memory_monitor = MemoryMonitor(warn_threshold_mb=1500, abort_threshold_mb=2000, logger=logger.info)
    llm_caller = LLMCaller(logger=logger)

    # Load prompts
    try:
        with open(PROMPT_FILE, "r", encoding="utf-8") as f:
            summary_prompt = f.read()
    except Exception as e:
        logger.error(f"Error reading prompt file: {e}")
        return False

    cross_check_prompt = None
    if os.path.exists(CROSS_CHECK_PROMPT_FILE):
        try:
            with open(CROSS_CHECK_PROMPT_FILE, "r", encoding="utf-8") as f:
                cross_check_prompt = f.read()
        except Exception:
            logger.warning("Could not load cross-check prompt. Skipping cross-check pass.")

    total_passes = 3 if cross_check_prompt else 2

    # ==========================================================================
    # Pass 1: Text Extraction
    # ==========================================================================
    logger.pass_start("Extraction", 1, total_passes)

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
        logger.pass_failed("Extraction", str(e), recoverable=False)
        return False
    except Exception as e:
        logger.pass_failed("Extraction", str(e), recoverable=False)
        return False

    logger.pass_complete("Extraction", success=True)

    # ==========================================================================
    # Pass 2: Summary Generation
    # ==========================================================================
    logger.pass_start("Summary", 2, total_passes)

    try:
        with memory_monitor.track_operation("Summary Generation"):
            summary = llm_caller.call(summary_prompt, text, task_type="summary")

            if not summary:
                raise SummaryPassError("LLM returned empty response")

            logger.info(f"Generated summary: {len(summary)} chars")

    except Exception as e:
        logger.pass_failed("Summary", str(e), recoverable=True)
        return False

    logger.pass_complete("Summary", success=True)

    # ==========================================================================
    # Pass 3: Cross-Check Verification (Optional)
    # ==========================================================================
    if cross_check_prompt:
        logger.pass_start("CrossCheck", 3, total_passes)

        try:
            with memory_monitor.track_operation("Cross-Check"):
                # Format the cross-check prompt with summary and original
                formatted_prompt = cross_check_prompt.replace("{summary}", summary).replace("{original}", text[:50000])

                verified_summary = llm_caller.call(formatted_prompt, "", task_type="cross_check")

                if verified_summary:
                    # Check if cross-check made meaningful changes
                    if len(verified_summary) > len(summary) * 0.9:
                        summary = verified_summary
                        logger.info("Cross-check completed with improvements")
                    else:
                        logger.warning("Cross-check returned shorter result. Using original summary.")
                else:
                    logger.warning("Cross-check returned empty. Using original summary.")

        except Exception as e:
            logger.pass_failed("CrossCheck", str(e), recoverable=True)
            logger.warning("Continuing with unverified summary")

        logger.pass_complete("CrossCheck", success=True)

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

    output_file = os.path.join(output_dir, "AI_OUTPUT.docx")
    title = os.path.basename(input_path)

    if not save_to_docx(summary, output_file, title, logger):
        return False

    # ==========================================================================
    # Save to Case Data
    # ==========================================================================
    try:
        data_manager = CaseDataManager()
        file_num_match = re.search(r"(\d{4}\.\d{3})", input_path)

        if file_num_match:
            file_num = file_num_match.group(1)
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            clean_var_name = re.sub(r"[^a-zA-Z0-9_]", "_", base_name.lower())
            var_key = f"summary_{clean_var_name}"

            logger.info(f"Saving to case data: {var_key} for {file_num}")
            data_manager.save_variable(file_num, var_key, summary, source="summarize_agent")
    except Exception as e:
        logger.warning(f"Could not save to case data: {e}")

    return True


def process_directory(dir_path: str, logger: AgentLogger):
    """Process all documents in a directory."""
    files_to_process = []

    for root, _, files in os.walk(dir_path):
        for file in files:
            if file.lower().endswith(('.pdf', '.docx')):
                if "AI_OUTPUT" in file:
                    continue
                files_to_process.append(os.path.join(root, file))

    if not files_to_process:
        logger.info("No suitable files found in directory.")
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
    logger = AgentLogger("Summarize", file_number=file_number)

    # Also create legacy log_event for backward compatibility
    log_event = create_legacy_log_event("Summarize", LEGACY_LOG_FILE)

    logger.info(f"Starting agent for: {input_path}")

    # Directory handling
    if os.path.isdir(input_path):
        logger.info(f"Input is a directory. Scanning for files...")
        process_directory(input_path, logger)
        logger.info("Directory processing complete.")
        sys.exit(0)

    # Single file processing
    success = process_document(input_path, logger)

    if success:
        logger.info("Agent finished successfully.")
        sys.exit(0)
    else:
        logger.error("Agent finished with errors.")
        sys.exit(1)


if __name__ == '__main__':
    main()
