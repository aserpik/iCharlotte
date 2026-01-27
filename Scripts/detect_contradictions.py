"""
Contradiction Detection Agent for iCharlotte

Analyzes multiple document summaries to identify factual contradictions
and inconsistencies across the case file.

Features:
- Multi-document comparison
- Severity classification
- Suggested resolution actions
- Integration with CaseDataManager
"""

import os
import sys
import re
import json
import datetime
import argparse
from typing import List, Dict, Optional, Any
from dataclasses import dataclass, field, asdict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_COLOR_INDEX

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import shared infrastructure
from icharlotte_core.agent_logger import AgentLogger
from icharlotte_core.llm_config import LLMCaller
from icharlotte_core.memory_monitor import MemoryMonitor
from icharlotte_core.exceptions import LLMError

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
PROMPT_FILE = os.path.join(SCRIPTS_DIR, "CONTRADICTION_DETECTION_PROMPT.txt")
LEGACY_LOG_FILE = r"C:\GeminiTerminal\Contradiction_activity.log"


# =============================================================================
# Data Classes
# =============================================================================

@dataclass
class Claim:
    """A factual claim from a document."""
    source: str
    text: str
    context: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class Contradiction:
    """A contradiction between two or more claims."""
    id: int
    topic: str
    claim_1: Claim
    claim_2: Claim
    additional_claims: List[Claim] = field(default_factory=list)
    severity: str = "medium"  # high, medium, low
    analysis: str = ""
    resolution_needed: bool = True
    suggested_action: str = ""

    def to_dict(self) -> dict:
        result = {
            'id': self.id,
            'topic': self.topic,
            'claim_1': self.claim_1.to_dict(),
            'claim_2': self.claim_2.to_dict(),
            'additional_claims': [c.to_dict() for c in self.additional_claims],
            'severity': self.severity,
            'analysis': self.analysis,
            'resolution_needed': self.resolution_needed,
            'suggested_action': self.suggested_action
        }
        return result


@dataclass
class ContradictionReport:
    """Report containing all detected contradictions."""
    contradictions: List[Contradiction] = field(default_factory=list)
    file_number: str = ""
    generated: str = ""
    documents_analyzed: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        high = sum(1 for c in self.contradictions if c.severity == "high")
        medium = sum(1 for c in self.contradictions if c.severity == "medium")
        low = sum(1 for c in self.contradictions if c.severity == "low")

        key_concerns = [
            c.topic for c in self.contradictions
            if c.severity == "high"
        ][:5]

        return {
            'file_number': self.file_number,
            'generated': self.generated,
            'documents_analyzed': self.documents_analyzed,
            'contradictions': [c.to_dict() for c in self.contradictions],
            'summary': {
                'total_contradictions': len(self.contradictions),
                'high_severity': high,
                'medium_severity': medium,
                'low_severity': low,
                'key_concerns': key_concerns
            }
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ContradictionReport':
        report = cls(
            file_number=data.get('file_number', ''),
            generated=data.get('generated', ''),
            documents_analyzed=data.get('documents_analyzed', [])
        )

        for c_data in data.get('contradictions', []):
            claim_1 = Claim(**c_data.get('claim_1', {}))
            claim_2 = Claim(**c_data.get('claim_2', {}))
            additional = [Claim(**c) for c in c_data.get('additional_claims', [])]

            contradiction = Contradiction(
                id=c_data.get('id', 0),
                topic=c_data.get('topic', ''),
                claim_1=claim_1,
                claim_2=claim_2,
                additional_claims=additional,
                severity=c_data.get('severity', 'medium'),
                analysis=c_data.get('analysis', ''),
                resolution_needed=c_data.get('resolution_needed', True),
                suggested_action=c_data.get('suggested_action', '')
            )
            report.contradictions.append(contradiction)

        return report


# =============================================================================
# Document Output
# =============================================================================

def save_report_to_docx(report: ContradictionReport, output_path: str, logger: AgentLogger) -> bool:
    """Save contradiction report to DOCX."""
    try:
        doc = Document()

        # Title
        title = doc.add_paragraph()
        run = title.add_run("CONTRADICTION ANALYSIS REPORT")
        run.bold = True
        run.underline = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)

        # Metadata
        doc.add_paragraph(f"File Number: {report.file_number}")
        doc.add_paragraph(f"Generated: {report.generated}")
        doc.add_paragraph(f"Documents Analyzed: {len(report.documents_analyzed)}")

        # Summary
        doc.add_heading("Executive Summary", level=1)

        summary = report.to_dict()['summary']
        p = doc.add_paragraph()
        p.add_run(f"Total Contradictions Found: {summary['total_contradictions']}\n")
        p.add_run(f"High Severity: ").bold = True
        p.add_run(f"{summary['high_severity']}\n")
        p.add_run(f"Medium Severity: {summary['medium_severity']}\n")
        p.add_run(f"Low Severity: {summary['low_severity']}\n")

        if summary['key_concerns']:
            p.add_run("\nKey Concerns:\n").bold = True
            for concern in summary['key_concerns']:
                p.add_run(f"  - {concern}\n")

        # Detailed Contradictions
        doc.add_heading("Detailed Findings", level=1)

        # Group by severity
        by_severity = {'high': [], 'medium': [], 'low': []}
        for c in report.contradictions:
            by_severity[c.severity].append(c)

        severity_headers = {
            'high': 'High Severity Issues',
            'medium': 'Medium Severity Issues',
            'low': 'Low Severity Issues'
        }

        for severity in ['high', 'medium', 'low']:
            contradictions = by_severity[severity]
            if not contradictions:
                continue

            doc.add_heading(severity_headers[severity], level=2)

            for c in contradictions:
                # Topic heading
                p = doc.add_paragraph()
                run = p.add_run(f"#{c.id}: {c.topic}")
                run.bold = True
                run.font.size = Pt(12)

                if severity == 'high':
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

                # Claims
                table = doc.add_table(rows=2, cols=2)
                table.style = 'Table Grid'

                # Header row
                table.rows[0].cells[0].text = "Source"
                table.rows[0].cells[1].text = "Claim"

                # Claim 1
                table.rows[0].cells[0].paragraphs[0].add_run(c.claim_1.source).bold = True
                table.rows[0].cells[1].text = c.claim_1.text

                # Claim 2
                table.rows[1].cells[0].paragraphs[0].add_run(c.claim_2.source).bold = True
                table.rows[1].cells[1].text = c.claim_2.text

                # Analysis
                p = doc.add_paragraph()
                p.add_run("Analysis: ").bold = True
                p.add_run(c.analysis)

                # Suggested action
                if c.suggested_action:
                    p = doc.add_paragraph()
                    p.add_run("Suggested Action: ").bold = True
                    p.add_run(c.suggested_action)

                doc.add_paragraph()  # Spacing

        doc.save(output_path)
        logger.output_file(output_path)
        return True

    except Exception as e:
        logger.error(f"Error saving report to DOCX: {e}")
        return False


# =============================================================================
# Main Processing
# =============================================================================

def gather_case_summaries(file_number: str, logger: AgentLogger,
                          selected_summaries: List[str] = None) -> Dict[str, str]:
    """
    Gather all summaries for a case from CaseDataManager.

    Args:
        file_number: The case file number.
        logger: AgentLogger instance.
        selected_summaries: Optional list of specific summary names to include.
                           If None, includes all available summaries.

    Returns:
        Dictionary of document_name -> summary_text.
    """
    summaries = {}

    try:
        data_manager = CaseDataManager()

        # Get all variables for the case
        variables = data_manager.get_all_variables(file_number)

        for var_name, var_data in variables.items():
            # Filter for summary-type variables
            if any(tag in var_name.lower() for tag in ['summary', 'depo', 'extraction']):
                # If selected_summaries is provided, only include those
                if selected_summaries is not None and var_name not in selected_summaries:
                    continue

                value = var_data.get('value', '')
                source = var_data.get('source', var_name)

                if value and len(value) > 100:  # Skip very short entries
                    summaries[source] = value
                    logger.info(f"Loaded summary: {var_name}")

    except Exception as e:
        logger.warning(f"Error loading case summaries: {e}")

    return summaries


def detect_contradictions(
    summaries: Dict[str, str],
    file_number: str,
    logger: AgentLogger
) -> Optional[ContradictionReport]:
    """
    Detect contradictions between document summaries.

    Args:
        summaries: Dictionary of document_name -> summary_text.
        file_number: Case file number.
        logger: AgentLogger instance.

    Returns:
        ContradictionReport or None on failure.
    """
    if len(summaries) < 2:
        logger.warning("Need at least 2 summaries for contradiction detection")
        return None

    memory_monitor = MemoryMonitor(warn_threshold_mb=1500, abort_threshold_mb=2000, logger=logger.info)
    llm_caller = LLMCaller(logger=logger)

    # Load prompt
    try:
        with open(PROMPT_FILE, "r", encoding="utf-8") as f:
            prompt = f.read()
    except Exception as e:
        logger.error(f"Error reading prompt file: {e}")
        return None

    # Build document content
    doc_content = ""
    for name, text in summaries.items():
        doc_content += f"\n\n=== {name} ===\n{text[:20000]}"  # Limit per doc

    logger.pass_start("Contradiction Analysis", 1, 1)

    try:
        with memory_monitor.track_operation("Contradiction Analysis"):
            # Use agent_contradict config from LLM settings (default: gemini-3-pro-preview)
            response = llm_caller.call(prompt, doc_content, agent_id="agent_contradict")

            if not response:
                raise LLMError("LLM returned empty response")

            # Parse JSON from response
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                data = json.loads(json_match.group())
            else:
                raise ValueError("No JSON found in response")

            # Build report
            report = ContradictionReport(
                file_number=file_number,
                generated=datetime.datetime.now().isoformat(),
                documents_analyzed=list(summaries.keys())
            )

            for c_data in data.get('contradictions', []):
                claim_1_data = c_data.get('claim_1', {})
                claim_2_data = c_data.get('claim_2', {})

                claim_1 = Claim(
                    source=claim_1_data.get('source', ''),
                    text=claim_1_data.get('text', ''),
                    context=claim_1_data.get('context', '')
                )
                claim_2 = Claim(
                    source=claim_2_data.get('source', ''),
                    text=claim_2_data.get('text', ''),
                    context=claim_2_data.get('context', '')
                )

                additional = []
                for ac in c_data.get('additional_claims', []):
                    additional.append(Claim(
                        source=ac.get('source', ''),
                        text=ac.get('text', ''),
                        context=ac.get('context', '')
                    ))

                contradiction = Contradiction(
                    id=c_data.get('id', len(report.contradictions) + 1),
                    topic=c_data.get('topic', ''),
                    claim_1=claim_1,
                    claim_2=claim_2,
                    additional_claims=additional,
                    severity=c_data.get('severity', 'medium'),
                    analysis=c_data.get('analysis', ''),
                    resolution_needed=c_data.get('resolution_needed', True),
                    suggested_action=c_data.get('suggested_action', '')
                )
                report.contradictions.append(contradiction)

            logger.info(f"Found {len(report.contradictions)} contradictions")

    except json.JSONDecodeError as e:
        logger.pass_failed("Contradiction Analysis", f"JSON parse error: {e}", recoverable=True)
        return None
    except Exception as e:
        logger.pass_failed("Contradiction Analysis", str(e), recoverable=True)
        return None

    logger.pass_complete("Contradiction Analysis", success=True)
    return report


def get_output_directory(file_number: str) -> str:
    """Get output directory for the case."""
    # Default path structure
    base = r"C:\Current Clients"

    # Try to find case folder
    if os.path.exists(base):
        for client_folder in os.listdir(base):
            client_path = os.path.join(base, client_folder)
            if os.path.isdir(client_path):
                for case_folder in os.listdir(client_path):
                    if file_number in case_folder:
                        return os.path.join(client_path, case_folder, "NOTES", "AI OUTPUT")

    # Fallback
    return os.path.join(os.getcwd(), "output")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Detect contradictions between document summaries in a case."
    )
    parser.add_argument("file_number", help="The case file number (e.g., 2024.123)")
    parser.add_argument("--summaries", type=str, default=None,
                        help="Comma-separated list of summary names to include (default: all)")

    args = parser.parse_args()

    file_number = args.file_number.strip()

    # Parse selected summaries if provided
    selected_summaries = None
    if args.summaries:
        selected_summaries = [s.strip() for s in args.summaries.split(",") if s.strip()]

    # Validate file number format
    if not re.match(r'\d{4}\.\d{3}', file_number):
        print(f"Warning: File number '{file_number}' may not be in expected format (YYYY.NNN)")

    # Initialize logger
    logger = AgentLogger("Contradiction", file_number=file_number)
    logger.info(f"Starting contradiction detection for case: {file_number}")

    if selected_summaries:
        logger.info(f"Filtering to {len(selected_summaries)} selected summaries")

    # Gather summaries
    summaries = gather_case_summaries(file_number, logger, selected_summaries)

    if not summaries:
        logger.error("No summaries found for this case. Run summarization agents first.")
        sys.exit(1)

    logger.info(f"Found {len(summaries)} document summaries")

    # Detect contradictions (uses agent_contradict config from LLM settings)
    report = detect_contradictions(summaries, file_number, logger)

    if not report:
        logger.error("Failed to detect contradictions")
        sys.exit(1)

    # Get output path
    output_dir = get_output_directory(file_number)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Save outputs
    docx_path = os.path.join(output_dir, "Contradiction_Report.docx")
    json_path = os.path.join(output_dir, "Contradiction_Report.json")

    # Save DOCX
    save_report_to_docx(report, docx_path, logger)

    # Save JSON
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(report.to_dict(), f, indent=2)
        logger.info(f"Saved JSON to: {json_path}")
    except Exception as e:
        logger.warning(f"Could not save JSON: {e}")

    # Save to CaseDataManager
    try:
        data_manager = CaseDataManager()
        data_manager.save_variable(
            file_number,
            "contradiction_report",
            json.dumps(report.to_dict()),
            source="contradiction_agent",
            extra_tags=["Contradictions", "Analysis"]
        )
        logger.info("Saved report to case data")
    except Exception as e:
        logger.warning(f"Could not save to case data: {e}")

    # Summary output
    summary = report.to_dict()['summary']
    logger.info(f"Analysis complete: {summary['total_contradictions']} contradictions found")
    logger.info(f"  High: {summary['high_severity']}, Medium: {summary['medium_severity']}, Low: {summary['low_severity']}")

    sys.exit(0)


if __name__ == '__main__':
    main()
