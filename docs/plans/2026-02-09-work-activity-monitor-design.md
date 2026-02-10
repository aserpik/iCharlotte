# Work Activity Monitor - Design Document

**Date:** 2026-02-09
**Status:** Approved

## Overview

A Python background service that monitors daily work activity on Windows, generates daily summaries, and produces weekly "automation opportunity" reports ranked by estimated time savings with concrete implementation proposals.

## Architecture

Three independent processes coordinated through a shared SQLite database:

### 1. Tracker (continuous, launched on logon)
Captures window titles, screenshots, clipboard events, file system events, and Outlook email activity. Writes structured data to SQLite. Sends screenshots to Gemini Vision for text descriptions, stores descriptions in DB, deletes image files after 48 hours.

### 2. Daily Reporter (triggered 7 PM daily)
Reads the day's structured data from SQLite, sends to LLMCaller for a markdown summary: top apps by time, workflow patterns, repetitive actions, per-case breakdowns. Saves to `reports/daily/`. Fires a Windows toast notification with top 3 findings.

### 3. Weekly Analyzer (triggered Friday 7 PM)
Reads 7 days of data, identifies recurring patterns, produces ranked automation opportunities with estimated time savings and specific implementation proposals. Saves to `reports/weekly/`. Fires a toast notification.

All three are separate Python invocations. Task Scheduler handles scheduling and crash recovery.

## Data Model (SQLite)

### `window_events`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| timestamp | DATETIME | Event time |
| app_name | TEXT | Process name |
| window_title | TEXT | Full window title |
| duration_seconds | REAL | Time in focus (calculated on next switch) |
| case_number | TEXT | Extracted case number (nullable) |

### `screenshots`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| timestamp | DATETIME | Capture time |
| file_path | TEXT | Image path (null after 48h cleanup) |
| vision_description | TEXT | Gemini Vision analysis |
| app_name | TEXT | Active app at capture time |
| window_title | TEXT | Active window at capture time |
| case_number | TEXT | Extracted case number (nullable) |

### `clipboard_events`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| timestamp | DATETIME | Event time |
| content_type | TEXT | Text, image, etc. |
| content_preview | TEXT | First 500 chars |
| source_app | TEXT | App on copy |
| source_window | TEXT | Window on copy |
| destination_app | TEXT | App on paste |
| destination_window | TEXT | Window on paste |
| case_number | TEXT | Extracted case number (nullable) |

### `file_events`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| timestamp | DATETIME | Event time |
| event_type | TEXT | open, save, rename, delete |
| file_path | TEXT | Full file path |
| app_name | TEXT | App that triggered event |
| case_number | TEXT | Extracted from file path (nullable) |

### `document_sessions`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| timestamp_start | DATETIME | Session start |
| timestamp_end | DATETIME | Session end |
| file_path | TEXT | Document path |
| app_name | TEXT | Application |
| duration_seconds | REAL | Total active time |
| save_count | INTEGER | Number of saves during session |
| case_number | TEXT | Extracted case number (nullable) |

### `email_events`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| timestamp | DATETIME | Event time |
| event_type | TEXT | sent, received, read, replied, forwarded |
| subject | TEXT | Email subject |
| recipients | TEXT | Comma-separated recipients |
| folder | TEXT | Outlook folder |
| has_attachment | BOOLEAN | Whether email has attachments |
| body_preview | TEXT | First 500 chars |
| case_number | TEXT | Extracted from subject (nullable) |

### `idle_periods`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| start_time | DATETIME | Idle start |
| end_time | DATETIME | Idle end |
| idle_type | TEXT | locked or inactive |
| duration_seconds | REAL | Total idle time |

### `report_metadata`
| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER PK | Auto-increment |
| report_type | TEXT | daily or weekly |
| report_date | TEXT | Date or week identifier |
| file_path | TEXT | Path to markdown file |
| generated_at | DATETIME | Generation timestamp |
| summary | TEXT | Brief summary of findings |

No foreign keys -- tables are independent event streams joined by timestamp ranges at query time.

## Case Number Extraction

A utility function tries multiple strategies in order:

1. **File path parsing** -- Regex `\d{4}\.\d{3}` from the full file path
2. **Window title parsing** -- Same regex against active window title
3. **Email subject line** -- Regex against Outlook subject from COM event
4. **Clipboard content** -- Regex against copied text
5. **Screenshot vision description** -- Gemini Vision prompt includes instruction to extract visible case numbers
6. **Session inference** -- If no case number found but file was opened from a folder containing one in its path, inherit it

## Capture Components

Five concurrent threads in the tracker process:

### WindowTracker
- Polls `GetForegroundWindow()` every 1 second
- On window change, logs previous window's duration, starts new event
- Extracts app name from process, title from window handle
- Runs case number extractor against title

### ScreenshotCapture
- Fixed 45-second interval (v1; adaptive in v2)
- Captures via `mss` or `Pillow`, saves as compressed JPEG
- Skips during idle periods
- Queues for async Gemini Vision analysis
- Stores vision description in SQLite
- Cleanup deletes images older than 48 hours

### ClipboardMonitor
- Watches via `AddClipboardFormatListener`
- On copy: records content preview + source app/window
- On paste: records destination app/window
- Runs case number extractor on content

### FileWatcher
- Uses `watchdog` library on case folders, Desktop, Documents, Downloads
- Logs open/save/rename/delete events
- Feeds document_sessions aggregation
- Case number extracted from file path

### OutlookMonitor
- Hooks into Outlook COM events (builds on existing `sent_items_monitor.py`)
- Captures send/receive/reply/forward
- Records subject, recipients, folder, attachment flag, body preview
- Case number extracted from subject line

All five write through a single thread-safe writer queue to avoid SQLite lock contention.

## Idle Detection

Sixth lightweight thread:

- Calls `GetLastInputInfo()` every 5 seconds
- **5-minute threshold**: checks `WTSQuerySessionInformation` for lock state
  - Locked screen → `locked` idle type
  - Unlocked, no input → `inactive` idle type (reading/thinking)
- **On idle start**: pauses screenshot capture; window/file watchers continue but flag events
- **On return from idle**: closes idle record, triggers burst of 3 screenshots at 10-second intervals
- **30+ minute idle**: treated as away (meeting/lunch), shown as gap in daily report

## Report Generation

### Daily Report (7 PM)
LLMCaller receives structured data and produces:
- Time breakdown by app and case number
- Document sessions with duration and save counts per case
- Email summary: sent/received counts, top recipients, average compose time
- Workflow sequences: most common app-switching patterns
- Clipboard transfers between apps
- Aggregated screenshot insights for repetitive UI actions

Output: 1-2 page markdown. Toast notification with top 3 findings.

### Weekly Report (Friday 7 PM)
LLMCaller receives 5 daily reports + raw event data and produces:
- Top 5 time sinks ranked by estimated weekly hours
- Repetitive patterns appearing 3+ days
- Concrete automation proposals with specific implementation details
- Estimated time savings per proposal
- Progress tracking against previous weeks' recommendations

Output: Markdown with ranked proposals. Toast notification.

## Retention Policy
- **Screenshots**: 48 hours (vision descriptions kept permanently in SQLite)
- **Structured logs**: 30 days
- **Reports**: Kept indefinitely

## LLM Strategy
- **Screenshot analysis**: Gemini Vision API (direct)
- **Report generation**: iCharlotte `LLMCaller` with automatic provider fallback

## File Structure

```
Scripts/work_monitor/
├── __init__.py
├── service.py              # Main entry point, launches all capture threads
├── config.py               # Settings (intervals, paths, retention, DB path)
├── db.py                   # SQLite schema, thread-safe writer queue, queries
├── case_extractor.py       # Case number extraction (path, title, regex, inference)
├── capture/
│   ├── __init__.py
│   ├── window_tracker.py   # Active window polling
│   ├── screenshot.py       # Screenshot capture + Gemini Vision analysis
│   ├── clipboard.py        # Clipboard copy/paste monitoring
│   ├── file_watcher.py     # File system events via watchdog
│   └── outlook_monitor.py  # Outlook COM event hooks
├── idle/
│   ├── __init__.py
│   └── detector.py         # Input monitoring, lock detection, burst-on-return
├── reports/
│   ├── __init__.py
│   ├── daily.py            # Daily summary generation
│   ├── weekly.py           # Weekly automation analysis
│   └── notifier.py         # Windows toast notifications
└── cleanup.py              # Screenshot deletion (48h), log pruning (30d)
```

Data output:
```
data/work_monitor/
├── db/
│   └── activity.db         # SQLite database
├── screenshots/            # Temp screenshot storage (48h retention)
└── reports/
    ├── daily/
    │   └── 2026-02-09.md
    └── weekly/
        └── 2026-W07.md
```

## Task Scheduler Entries
- **Tracker**: `python -m Scripts.work_monitor.service` -- on logon, restart on failure
- **Daily report**: `python -m Scripts.work_monitor.reports.daily` -- 7 PM daily
- **Weekly report**: `python -m Scripts.work_monitor.reports.weekly` -- Friday 7 PM

## Future (v2)
- Adaptive screenshot intervals (increase during rapid window switching)
- Keystroke logging for data-entry pattern detection
- Mouse click heatmaps for UI workflow analysis
- iCharlotte UI tab integration
- Cross-week trend analysis and seasonal pattern detection
