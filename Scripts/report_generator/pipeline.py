"""
Report Generator Pipeline - Main orchestrator for the 5-stage report generation process.

Stages:
    1. GATHER  - Collect case metadata and agent outputs
    2. REFINE  - Per-section LLM refinement with style matching
    3. ASSEMBLE - Render sections into docxtpl Word template
    4. POLISH  - Final LLM pass for transitions and consistency
    5. OUTPUT  - Save final report to case folder

Usage:
    python -m Scripts.report_generator.pipeline --file-number 1280.001
    python -m Scripts.report_generator.pipeline --file-number 1280.001 --type STATUS --prior-report path/to/prior.docx
    python -m Scripts.report_generator.pipeline --file-number 1280.001 --skip-polish
    python -m Scripts.report_generator.pipeline --file-number 1280.001 --skip-refine  # raw content, no LLM
"""

import argparse
import logging
import os
import sys
import time
from typing import Dict, Optional

logger = logging.getLogger(__name__)

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def run_pipeline(
    file_number: str,
    report_type: str = "FSR",
    prior_report_path: str = None,
    template_path: str = None,
    output_path: str = None,
    skip_refine: bool = False,
    skip_polish: bool = False,
    llm_caller=None,
) -> str:
    """
    Run the full report generation pipeline.

    Args:
        file_number: Case file number (e.g., "1280.001")
        report_type: "FSR" or "STATUS"
        prior_report_path: Path to prior report (for STATUS reports)
        template_path: Override template .docx path
        output_path: Override output .docx path
        skip_refine: Skip LLM refinement (use raw agent output)
        skip_polish: Skip the final LLM polish pass
        llm_caller: Pre-configured LLMCaller (creates one if None)

    Returns:
        Path to the generated report .docx
    """
    start_time = time.time()

    print(f"\n{'='*60}")
    print(f"  Report Generator Pipeline")
    print(f"  File: {file_number}  |  Type: {report_type}")
    print(f"{'='*60}\n")

    # Initialize LLM caller if needed (shared across stages)
    if llm_caller is None and not (skip_refine and skip_polish):
        llm_caller = _get_llm_caller()

    # ── Stage 1: GATHER ──────────────────────────────────────────
    print("[1/5] GATHER - Collecting case data and agent outputs...")
    from Scripts.report_generator.gather import gather_case_data

    gathered = gather_case_data(file_number, report_type, prior_report_path)

    print(f"  Case path: {gathered['case_path']}")
    print(f"  Metadata fields: {len(gathered['metadata'])}")
    print(f"  Sections with content: {len(gathered['sections'])}")
    print(f"  Sections to include: {gathered['included_sections']}")

    if not gathered["sections"] and not gathered["included_sections"]:
        print("\n  WARNING: No agent outputs found. Generating placeholder report.")

    # ── Stage 2: REFINE ──────────────────────────────────────────
    if skip_refine:
        print("[2/5] REFINE - SKIPPED (--skip-refine)")
        refined = _passthrough_sections(gathered)
    else:
        print("[2/5] REFINE - Reshaping agent outputs into report prose...")
        from Scripts.report_generator.refine import refine_sections

        refined = refine_sections(gathered, llm_caller=llm_caller)

        for section, text in refined.items():
            print(f"  {section}: {len(text)} chars")

    # ── Stage 3: ASSEMBLE ────────────────────────────────────────
    print("[3/5] ASSEMBLE - Rendering template...")
    from Scripts.report_generator.assemble import assemble_report

    assembled_path = assemble_report(
        gathered, refined,
        template_path=template_path,
        output_path=output_path,
    )
    print(f"  Assembled: {assembled_path}")

    # ── Stage 4: POLISH ──────────────────────────────────────────
    if skip_polish:
        print("[4/5] POLISH - SKIPPED (--skip-polish)")
        final_path = assembled_path
    else:
        print("[4/5] POLISH - Smoothing transitions and cross-references...")
        from Scripts.report_generator.polish import polish_report

        final_path = polish_report(
            assembled_path, gathered, refined, llm_caller=llm_caller
        )
        print(f"  Polished: {final_path}")

    # ── Stage 5: OUTPUT ──────────────────────────────────────────
    elapsed = time.time() - start_time
    print(f"[5/5] OUTPUT - Done!")
    print(f"\n{'='*60}")
    print(f"  Report saved: {final_path}")
    print(f"  Elapsed: {elapsed:.1f}s")
    print(f"{'='*60}\n")

    # Log the generation in CaseDataManager
    _log_generation(file_number, final_path, report_type, gathered)

    return final_path


def _passthrough_sections(gathered: Dict) -> Dict[str, str]:
    """Use raw agent output without LLM refinement (for --skip-refine mode)."""
    from Scripts.report_generator.refine import PLACEHOLDER_TEMPLATES

    refined = {}
    for section in gathered["included_sections"]:
        raw = gathered["sections"].get(section)
        if raw:
            refined[section] = raw
        else:
            refined[section] = PLACEHOLDER_TEMPLATES.get(
                section, "This section will be updated in a subsequent report."
            )
    return refined


def _get_llm_caller():
    """Create an LLMCaller instance."""
    try:
        from icharlotte_core.llm_config import LLMCaller
        caller = LLMCaller()
        print("  LLM caller initialized")
        return caller
    except ImportError:
        print("  WARNING: LLMCaller not available. Refinement/polish will use raw content.")
        return None


def _log_generation(file_number: str, output_path: str, report_type: str,
                    gathered: Dict):
    """Log the report generation to CaseDataManager."""
    try:
        from Scripts.case_data_manager import CaseDataManager
        from datetime import datetime

        mgr = CaseDataManager()
        mgr.save_variable(
            file_number,
            "last_generated_report",
            {
                "path": output_path,
                "type": report_type,
                "date": datetime.now().isoformat(),
                "sections_included": gathered.get("included_sections", []),
            },
            source="report_generator",
            auto_tag=False,
        )
    except Exception as e:
        logger.debug(f"Could not log generation: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="Generate a litigation report from iCharlotte agent outputs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Full Status Report (ILP):
    python -m Scripts.report_generator.pipeline --file-number 1280.001

  Status Report with prior report for delta:
    python -m Scripts.report_generator.pipeline --file-number 1280.001 --type STATUS --prior-report path/to/prior.docx

  Quick test (no LLM, raw content):
    python -m Scripts.report_generator.pipeline --file-number 1280.001 --skip-refine --skip-polish
        """
    )

    parser.add_argument(
        "--file-number", "-f", required=True,
        help="Case file number (e.g., 1280.001)"
    )
    parser.add_argument(
        "--type", "-t", default="FSR", choices=["FSR", "STATUS"],
        help="Report type: FSR (Full Status Report / ILP) or STATUS"
    )
    parser.add_argument(
        "--prior-report", "-p", default=None,
        help="Path to prior report .docx (for STATUS reports)"
    )
    parser.add_argument(
        "--template", default=None,
        help="Override path to .docx template"
    )
    parser.add_argument(
        "--output", "-o", default=None,
        help="Override output .docx path"
    )
    parser.add_argument(
        "--skip-refine", action="store_true",
        help="Skip LLM refinement (use raw agent output)"
    )
    parser.add_argument(
        "--skip-polish", action="store_true",
        help="Skip the final LLM polish pass"
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Enable verbose logging"
    )

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    try:
        result = run_pipeline(
            file_number=args.file_number,
            report_type=args.type,
            prior_report_path=args.prior_report,
            template_path=args.template,
            output_path=args.output,
            skip_refine=args.skip_refine,
            skip_polish=args.skip_polish,
        )
        sys.exit(0)
    except FileNotFoundError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
    except Exception as e:
        logger.exception(f"Pipeline failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
