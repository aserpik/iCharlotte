"""
Daily Report Generator.

Run: python -m sideprojects.work_monitor.reports.daily [YYYY-MM-DD]
Scheduled: 7 PM daily via Windows Task Scheduler.
"""

import os
import sys
import logging
from datetime import datetime

# Add project root to path for icharlotte_core imports
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))
)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from sideprojects.work_monitor import config
from sideprojects.work_monitor.db import ActivityDB
from sideprojects.work_monitor.reports.notifier import send_toast

logger = logging.getLogger("work_monitor.daily")


def generate_daily_report(date_str: str = None):
    """Generate a daily summary report for the given date."""
    config.ensure_dirs()
    db = ActivityDB()

    date_str = date_str or datetime.now().strftime("%Y-%m-%d")
    events = db.get_daily_events(date_str)

    # Check if there's enough data
    total_events = sum(len(v) for v in events.values())
    if total_events < 5:
        print(f"Not enough data for {date_str} ({total_events} events)")
        return

    # Build structured data summary for the LLM
    data_summary = _build_data_summary(events, date_str)

    # Import LLMCaller from icharlotte_core
    try:
        from icharlotte_core.llm_config import call_llm
    except ImportError:
        logger.error("icharlotte_core not available, cannot generate report")
        return

    prompt = """You are a work activity analyst. Given the structured activity data below,
generate a concise daily work summary report in markdown format.

Include these sections:

## Time Breakdown
- Top applications by active time (with percentages of total active time)
- Per-case breakdown of time spent (use case numbers where available)

## Document Sessions
- Files worked on, duration, and save counts
- Organized by case number when available

## Email Activity
- Sent/received counts
- Top recipients
- Key subjects and thread patterns

## Workflow Patterns
- Common app-switching sequences (e.g., Outlook -> Word -> Outlook repeated 12x)
- Clipboard transfer patterns between apps
- Screenshot insights for repetitive UI actions

## Key Observations
- Top 3 notable findings or patterns
- Any inefficiencies or repetitive manual tasks spotted
- Time sinks worth investigating

Keep the report factual and concise (1-2 pages). Use bullet points.
Format durations as hours and minutes (e.g., 1h 23m)."""

    result = call_llm(prompt, data_summary, task_type="summary")

    if not result:
        logger.error("LLM generation failed for daily report")
        return

    # Save report
    report_path = os.path.join(config.DAILY_REPORT_DIR, f"{date_str}.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"# Daily Work Activity Report - {date_str}\n\n")
        f.write(f"*Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}*\n\n")
        f.write(result)

    # Record in metadata table
    db.insert_report_metadata(
        report_type="daily",
        report_date=date_str,
        file_path=report_path,
        generated_at=datetime.now().isoformat(),
        summary=result[:200],
    )

    # Send toast notification
    findings = _extract_top_findings(result)
    send_toast("Daily Work Report", findings, report_path)

    print(f"Daily report saved to {report_path}")
    return report_path


def _build_data_summary(events: dict, date_str: str) -> str:
    """Convert raw event dicts into a structured text summary for the LLM."""
    parts = [f"Activity Data for: {date_str}\n"]

    # Window events summary
    window_events = events.get('window_events', [])
    if window_events:
        app_time: dict[str, float] = {}
        case_time: dict[str, float] = {}
        for e in window_events:
            app = e.get('app_name', 'unknown')
            dur = e.get('duration_seconds') or 0
            app_time[app] = app_time.get(app, 0) + dur
            cn = e.get('case_number')
            if cn:
                case_time[cn] = case_time.get(cn, 0) + dur

        total_active = sum(app_time.values())
        parts.append(f"## Active Window Time (total: {total_active/60:.0f} min)")
        for app, seconds in sorted(app_time.items(), key=lambda x: -x[1])[:15]:
            pct = (seconds / total_active * 100) if total_active else 0
            parts.append(f"- {app}: {seconds/60:.1f} min ({pct:.0f}%)")

        if case_time:
            parts.append("\n## Time by Case Number")
            for cn, seconds in sorted(case_time.items(), key=lambda x: -x[1]):
                parts.append(f"- Case {cn}: {seconds/60:.1f} min")

    # Screenshot descriptions
    screenshots = events.get('screenshots', [])
    described = [s for s in screenshots if s.get('vision_description')]
    if described:
        parts.append(f"\n## Screenshot Descriptions ({len(described)}/{len(screenshots)} analyzed)")
        for s in described[:20]:
            parts.append(
                f"- [{s['timestamp'][:19]}] {s.get('app_name', '')}: "
                f"{s.get('vision_description', '')}"
            )

    # Email events
    emails = events.get('email_events', [])
    if emails:
        sent = [e for e in emails if e['event_type'] in ('sent', 'replied', 'forwarded')]
        received = [e for e in emails if e['event_type'] == 'received']
        parts.append(f"\n## Email Events ({len(sent)} sent, {len(received)} received)")
        for e in emails[:30]:
            parts.append(
                f"- [{e['event_type']}] {e.get('subject', '')[:80]} | "
                f"To: {e.get('recipients', '')[:60]} | "
                f"Case: {e.get('case_number', 'N/A')}"
            )

    # Document sessions
    doc_sessions = events.get('document_sessions', [])
    if doc_sessions:
        parts.append(f"\n## Document Sessions ({len(doc_sessions)})")
        for d in doc_sessions:
            dur = d.get('duration_seconds') or 0
            parts.append(
                f"- {d.get('file_path', 'unknown')}: {dur/60:.1f} min, "
                f"{d.get('save_count', 0)} saves (case: {d.get('case_number', 'N/A')})"
            )

    # File events
    file_events = events.get('file_events', [])
    if file_events:
        parts.append(f"\n## File Events ({len(file_events)} total)")
        for e in file_events[:30]:
            parts.append(
                f"- [{e['event_type']}] {e.get('file_path', '')} "
                f"(case: {e.get('case_number', 'N/A')})"
            )

    # Clipboard events
    clips = events.get('clipboard_events', [])
    if clips:
        parts.append(f"\n## Clipboard Events ({len(clips)} total)")
        for c in clips[:20]:
            preview = (c.get('content_preview', '') or '')[:80]
            parts.append(
                f"- From {c.get('source_app', '?')} "
                f"({c.get('source_window', '')[:40]}): {preview}"
            )

    # Idle periods
    idles = events.get('idle_periods', [])
    if idles:
        total_idle = sum(i.get('duration_seconds', 0) for i in idles)
        parts.append(f"\n## Idle Periods ({len(idles)} total, {total_idle/60:.0f} min)")
        for i in idles:
            dur = i.get('duration_seconds', 0)
            parts.append(
                f"- {i['idle_type']}: {dur/60:.1f} min "
                f"({i.get('start_time', '')[:19]} to {i.get('end_time', '')[:19]})"
            )

    return "\n".join(parts)


def _extract_top_findings(report_text: str) -> str:
    """Extract the Key Observations section for toast notification."""
    lines = report_text.split("\n")
    in_section = False
    findings = []
    for line in lines:
        if "key observation" in line.lower():
            in_section = True
            continue
        if in_section:
            if line.startswith("##"):
                break
            stripped = line.strip()
            if stripped.startswith("-") or stripped.startswith("*"):
                findings.append(stripped[2:].strip())

    return "\n".join(findings[:3]) if findings else "Daily report generated."


if __name__ == "__main__":
    date_arg = sys.argv[1] if len(sys.argv) > 1 else None
    generate_daily_report(date_arg)
