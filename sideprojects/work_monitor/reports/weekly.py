"""
Weekly Automation Analyzer.

Run: python -m sideprojects.work_monitor.reports.weekly [YYYY-Www]
Scheduled: Friday 7 PM via Windows Task Scheduler.
"""

import os
import sys
import logging
from datetime import datetime, timedelta

# Add project root to path for icharlotte_core imports
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))
)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from sideprojects.work_monitor import config
from sideprojects.work_monitor.db import ActivityDB
from sideprojects.work_monitor.reports.notifier import send_toast

logger = logging.getLogger("work_monitor.weekly")


def generate_weekly_report(week_str: str = None):
    """Generate weekly automation analysis."""
    config.ensure_dirs()
    db = ActivityDB()

    today = datetime.now()
    week_str = week_str or today.strftime("%Y-W%V")

    # Calculate Monday-Sunday of the target week
    # Parse week string if provided, otherwise use current week
    if week_str and week_str != today.strftime("%Y-W%V"):
        year, week_num = week_str.split("-W")
        monday = datetime.strptime(f"{year}-W{week_num}-1", "%Y-W%V-%u")
    else:
        monday = today - timedelta(days=today.weekday())

    monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
    sunday = monday + timedelta(days=7)

    start_str = monday.strftime("%Y-%m-%d")
    end_str = sunday.strftime("%Y-%m-%d")

    # Gather weekly event data
    events = db.get_weekly_events(start_str, end_str)
    total_events = sum(len(v) for v in events.values())

    if total_events < 20:
        print(f"Not enough data for {week_str} ({total_events} events)")
        return

    # Load daily reports for the week
    daily_reports = _load_daily_reports(monday, sunday)

    # Load previous weekly report for progress tracking
    prev_report = _load_previous_weekly_report(week_str)

    # Build comprehensive data package
    data_text = _build_weekly_data(events, daily_reports, prev_report, start_str, end_str)

    # Import LLMCaller
    try:
        from icharlotte_core.llm_config import call_llm
    except ImportError:
        logger.error("icharlotte_core not available, cannot generate report")
        return

    prompt = """You are a workplace automation analyst reviewing a week of work activity.
Given the week's activity data, daily reports, and previous recommendations,
produce a ranked automation opportunity analysis.

Include these sections:

## Weekly Summary
- Total active hours, breakdown by day
- Most-used applications and documents
- Cases worked on and time per case

## Top 5 Time Sinks
Ranked by estimated weekly hours. For each:
- Description of the repetitive pattern observed
- Estimated hours per week spent on this pattern
- How many days it appeared (consistency)
- Specific examples from the data

## Automation Proposals
For each time sink, provide a concrete proposal:
1. What to build (be specific: "Python script that...", "Word macro that...")
2. Technology/tool recommendation
3. Estimated implementation effort (hours)
4. Estimated weekly time savings after implementation
5. ROI: how many weeks until the implementation time is recovered

## Progress Report
If previous recommendations exist, compare:
- What was implemented since last week?
- What remains from last week?
- New patterns discovered this week

## Priority Matrix
Final ranked list combining:
- Quick wins (< 2 hours to build, saves 30+ min/week)
- High impact (saves 2+ hours/week regardless of effort)
- Low-hanging fruit for next week

Be specific and actionable. Reference actual application names, file paths, and case numbers.
Format time estimates consistently (e.g., "2h 15m/week")."""

    result = call_llm(prompt, data_text, task_type="summary")

    if not result:
        logger.error("LLM generation failed for weekly report")
        return

    # Save report
    report_path = os.path.join(config.WEEKLY_REPORT_DIR, f"{week_str}.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"# Weekly Automation Analysis - {week_str}\n")
        f.write(f"# Period: {start_str} to {end_str}\n\n")
        f.write(f"*Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}*\n\n")
        f.write(result)

    db.insert_report_metadata(
        report_type="weekly",
        report_date=week_str,
        file_path=report_path,
        generated_at=datetime.now().isoformat(),
        summary=result[:200],
    )

    # Send toast with top proposal
    top_proposal = _extract_top_proposal(result)
    send_toast("Weekly Automation Report", top_proposal, report_path)

    print(f"Weekly report saved to {report_path}")
    return report_path


def _load_daily_reports(monday: datetime, sunday: datetime) -> str:
    """Load all daily reports for the given week."""
    parts = []
    current = monday
    while current < sunday:
        date_str = current.strftime("%Y-%m-%d")
        path = os.path.join(config.DAILY_REPORT_DIR, f"{date_str}.md")
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                parts.append(f"--- Daily Report: {date_str} ---\n{content}\n")
            except Exception:
                pass
        current += timedelta(days=1)

    return "\n".join(parts) if parts else "No daily reports available for this week."


def _load_previous_weekly_report(current_week: str) -> str:
    """Load the previous week's report for progress tracking."""
    try:
        year, week_num = current_week.split("-W")
        prev_week_num = int(week_num) - 1
        if prev_week_num < 1:
            prev_year = int(year) - 1
            prev_week_str = f"{prev_year}-W52"
        else:
            prev_week_str = f"{year}-W{prev_week_num:02d}"

        path = os.path.join(config.WEEKLY_REPORT_DIR, f"{prev_week_str}.md")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return f"--- Previous Week Report ({prev_week_str}) ---\n{f.read()}"
    except Exception:
        pass
    return "No previous weekly report available."


def _build_weekly_data(events: dict, daily_reports: str,
                       prev_report: str, start_str: str, end_str: str) -> str:
    """Build comprehensive data package for the LLM."""
    parts = [f"Week: {start_str} to {end_str}\n"]

    # Aggregate window events by app and case
    window_events = events.get('window_events', [])
    if window_events:
        app_time: dict[str, float] = {}
        case_time: dict[str, float] = {}
        day_hours: dict[str, float] = {}

        for e in window_events:
            app = e.get('app_name', 'unknown')
            dur = e.get('duration_seconds') or 0
            app_time[app] = app_time.get(app, 0) + dur

            cn = e.get('case_number')
            if cn:
                case_time[cn] = case_time.get(cn, 0) + dur

            day = e.get('timestamp', '')[:10]
            day_hours[day] = day_hours.get(day, 0) + dur

        total = sum(app_time.values())
        parts.append(f"## Weekly App Usage (total: {total/3600:.1f} hours)")
        for app, secs in sorted(app_time.items(), key=lambda x: -x[1])[:15]:
            parts.append(f"- {app}: {secs/3600:.1f}h ({secs/total*100:.0f}%)")

        if case_time:
            parts.append("\n## Time by Case")
            for cn, secs in sorted(case_time.items(), key=lambda x: -x[1]):
                parts.append(f"- Case {cn}: {secs/3600:.1f}h")

        if day_hours:
            parts.append("\n## Active Hours by Day")
            for day, secs in sorted(day_hours.items()):
                parts.append(f"- {day}: {secs/3600:.1f}h")

    # Email summary
    emails = events.get('email_events', [])
    if emails:
        sent = len([e for e in emails if e['event_type'] in ('sent', 'replied', 'forwarded')])
        received = len([e for e in emails if e['event_type'] == 'received'])

        # Top recipients
        recipient_count: dict[str, int] = {}
        for e in emails:
            if e['event_type'] in ('sent', 'replied', 'forwarded'):
                for r in (e.get('recipients', '') or '').split(', '):
                    r = r.strip()
                    if r:
                        recipient_count[r] = recipient_count.get(r, 0) + 1

        parts.append(f"\n## Email Summary ({sent} sent, {received} received)")
        if recipient_count:
            parts.append("Top recipients:")
            for r, count in sorted(recipient_count.items(), key=lambda x: -x[1])[:10]:
                parts.append(f"  - {r}: {count} emails")

    # Clipboard patterns
    clips = events.get('clipboard_events', [])
    if clips:
        source_counts: dict[str, int] = {}
        for c in clips:
            src = c.get('source_app', 'unknown')
            source_counts[src] = source_counts.get(src, 0) + 1

        parts.append(f"\n## Clipboard Activity ({len(clips)} copies)")
        for app, count in sorted(source_counts.items(), key=lambda x: -x[1])[:10]:
            parts.append(f"- From {app}: {count} copies")

    # File activity
    file_events = events.get('file_events', [])
    if file_events:
        type_counts: dict[str, int] = {}
        for f in file_events:
            et = f.get('event_type', 'unknown')
            type_counts[et] = type_counts.get(et, 0) + 1

        parts.append(f"\n## File Activity ({len(file_events)} events)")
        for et, count in sorted(type_counts.items(), key=lambda x: -x[1]):
            parts.append(f"- {et}: {count}")

    # Idle summary
    idles = events.get('idle_periods', [])
    if idles:
        total_idle = sum(i.get('duration_seconds', 0) for i in idles) / 3600
        parts.append(f"\n## Idle Time: {total_idle:.1f} hours total ({len(idles)} periods)")

    # Include daily reports
    parts.append(f"\n\n{daily_reports}")

    # Include previous week report
    parts.append(f"\n\n{prev_report}")

    return "\n".join(parts)


def _extract_top_proposal(report_text: str) -> str:
    """Extract the top automation proposal for toast notification."""
    lines = report_text.split("\n")
    in_section = False
    proposals = []
    for line in lines:
        if "automation proposal" in line.lower() or "priority matrix" in line.lower():
            in_section = True
            continue
        if in_section:
            if line.startswith("##") and "proposal" not in line.lower():
                break
            stripped = line.strip()
            if stripped.startswith(("1.", "-", "*")):
                proposals.append(stripped.lstrip("1.-* ").strip())

    if proposals:
        return f"Top recommendation: {proposals[0]}"
    return "Weekly automation analysis complete."


if __name__ == "__main__":
    week_arg = sys.argv[1] if len(sys.argv) > 1 else None
    generate_weekly_report(week_arg)
