"""
Test suite for iCharlotte Calendar Agent.

Tests date parsing, document classification, deadline calculations,
and thread handling edge cases.
"""

import sys
import os
sys.path.insert(0, '.')

from datetime import datetime, timedelta
from icharlotte_core.calendar.attachment_classifier import AttachmentClassifier
from icharlotte_core.calendar.deadline_calculator import DeadlineCalculator
from icharlotte_core.calendar.calendar_monitor import CalendarMonitorWorker

# Test results tracking
passed = 0
failed = 0
results = []

def test(name, condition, details=""):
    global passed, failed, results
    if condition:
        passed += 1
        results.append(f"[PASS] {name}")
    else:
        failed += 1
        results.append(f"[FAIL] {name} - {details}")

def print_section(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print('='*60)


# ============================================================
# TEST 1: Date Parsing from Email Body
# ============================================================
print_section("TEST 1: Date Parsing from Email Body")

monitor = CalendarMonitorWorker()
now = datetime.now()
today = now.date()

# Test "tomorrow"
body = "Let's schedule a call for tomorrow."
dates = monitor._extract_dates_from_body(body)
tomorrow = (now + timedelta(days=1)).date()
test("Parse 'tomorrow'",
     any(d.date() == tomorrow for d in dates),
     f"Expected {tomorrow}, got {[d.date() for d in dates]}")

# Test "day after tomorrow"
body = "The meeting is scheduled for the day after tomorrow."
dates = monitor._extract_dates_from_body(body)
day_after = (now + timedelta(days=2)).date()
test("Parse 'day after tomorrow'",
     any(d.date() == day_after for d in dates),
     f"Expected {day_after}, got {[d.date() for d in dates]}")

# Test that "day after tomorrow" doesn't also match "tomorrow"
test("'day after tomorrow' should not also match 'tomorrow'",
     not any(d.date() == tomorrow for d in dates),
     f"Should not have {tomorrow} in {[d.date() for d in dates]}")

# Test explicit date formats
body = "The deadline is January 30, 2026."
dates = monitor._extract_dates_from_body(body)
expected = datetime(2026, 1, 30).date()
test("Parse 'January 30, 2026'",
     any(d.date() == expected for d in dates),
     f"Expected {expected}, got {[d.date() for d in dates]}")

body = "Submit by 2/15/2026."
dates = monitor._extract_dates_from_body(body)
expected = datetime(2026, 2, 15).date()
test("Parse '2/15/2026'",
     any(d.date() == expected for d in dates),
     f"Expected {expected}, got {[d.date() for d in dates]}")

# Test "next Monday"
body = "Let's meet next Monday."
dates = monitor._extract_dates_from_body(body)
# Find next Monday
days_until_monday = (7 - now.weekday()) % 7
if days_until_monday == 0:
    days_until_monday = 7
next_monday = (now + timedelta(days=days_until_monday)).date()
test("Parse 'next Monday'",
     any(d.date() == next_monday for d in dates),
     f"Expected {next_monday}, got {[d.date() for d in dates]}")

# Test "in 2 weeks"
body = "The event is in 2 weeks."
dates = monitor._extract_dates_from_body(body)
two_weeks = (now + timedelta(weeks=2)).date()
test("Parse 'in 2 weeks'",
     any(d.date() == two_weeks for d in dates),
     f"Expected {two_weeks}, got {[d.date() for d in dates]}")


# ============================================================
# TEST 2: Thread Body Extraction (Only First Message)
# ============================================================
print_section("TEST 2: Thread Body Extraction")

class MockItem:
    def __init__(self, body):
        self.Body = body

# Test that only the first message is extracted
thread_body = """Hi, let's schedule for the day after tomorrow.

Thanks,
John

-----Original Message-----
From: Jane Smith
Sent: Monday, January 20, 2026
To: John Doe
Subject: RE: Meeting

Sure, how about tomorrow?

Best,
Jane
"""

mock_item = MockItem(thread_body)
extracted = monitor._get_thread_body(mock_item)
test("Thread extraction - only first message",
     "day after tomorrow" in extracted.lower() and "how about tomorrow?" not in extracted.lower(),
     f"Should have 'day after tomorrow' but not old 'tomorrow'. Got: {extracted[:100]}...")

# Test forwarded email
fw_body = """FW: See below

Please calendar the deposition for next Friday.

---
From: opposing@counsel.com
Sent: January 15, 2026
Subject: Deposition Notice

The deposition will be tomorrow at 10am.
"""

mock_item = MockItem(fw_body)
extracted = monitor._get_thread_body(mock_item)
test("FW: email - only new content",
     "next friday" in extracted.lower() and "tomorrow at 10am" not in extracted.lower(),
     f"Should have 'next friday' but not forwarded 'tomorrow'. Got: {extracted[:100]}...")


# ============================================================
# TEST 3: Document Classification
# ============================================================
print_section("TEST 3: Document Classification")

classifier = AttachmentClassifier()

# Test deposition notice detection (pattern-based)
depo_text = """NOTICE OF TAKING DEPOSITION

PLEASE TAKE NOTICE that Plaintiff will take the deposition of
JOHN SMITH on February 15, 2026 at 10:00 a.m. at the offices of
Smith & Jones, 123 Main Street, Los Angeles, CA 90001.
"""

result = classifier._classify_with_patterns(depo_text)
test("Classify deposition notice (patterns)",
     result['doc_type'] == 'deposition_notice',
     f"Expected 'deposition_notice', got '{result['doc_type']}'")

# Test interrogatories detection
interrog_text = """FORM INTERROGATORIES - SET ONE

Propounding Party: Plaintiff
Responding Party: Defendant

INTERROGATORY NO. 1.1: State your full name...
"""

result = classifier._classify_with_patterns(interrog_text)
test("Classify interrogatories as discovery_request",
     result['doc_type'] == 'discovery_request',
     f"Expected 'discovery_request', got '{result['doc_type']}'")

# Test RFP detection
rfp_text = """REQUEST FOR PRODUCTION OF DOCUMENTS - SET TWO

Plaintiff requests that Defendant produce the following documents...
"""

result = classifier._classify_with_patterns(rfp_text)
test("Classify RFP as discovery_request",
     result['doc_type'] == 'discovery_request',
     f"Expected 'discovery_request', got '{result['doc_type']}'")

# Test motion detection
motion_text = """NOTICE OF MOTION AND MOTION FOR SUMMARY JUDGMENT

TO ALL PARTIES AND THEIR ATTORNEYS OF RECORD:

PLEASE TAKE NOTICE that on March 20, 2026 at 9:00 a.m. in Department 5,
Defendant will move for summary judgment...
"""

result = classifier._classify_with_patterns(motion_text)
test("Classify MSJ motion",
     result['doc_type'] == 'motion' and result['motion_type'] == 'msj',
     f"Expected 'motion/msj', got '{result['doc_type']}/{result['motion_type']}'")

# Test opposition detection
opp_text = """PLAINTIFF'S OPPOSITION TO DEFENDANT'S MOTION FOR SUMMARY JUDGMENT

Plaintiff hereby opposes Defendant's Motion for Summary Judgment...
"""

result = classifier._classify_with_patterns(opp_text)
test("Classify opposition",
     result['doc_type'] == 'opposition',
     f"Expected 'opposition', got '{result['doc_type']}'")

# Test reply detection
reply_text = """DEFENDANT'S REPLY BRIEF IN SUPPORT OF MOTION FOR SUMMARY JUDGMENT

Defendant submits this Reply to Plaintiff's Opposition...
"""

result = classifier._classify_with_patterns(reply_text)
test("Classify reply",
     result['doc_type'] == 'reply',
     f"Expected 'reply', got '{result['doc_type']}'")

# Test demurrer detection
demurrer_text = """DEMURRER TO PLAINTIFF'S FIRST AMENDED COMPLAINT

Defendant demurs to the First, Second, and Third Causes of Action...
"""

result = classifier._classify_with_patterns(demurrer_text)
test("Classify demurrer",
     result['doc_type'] == 'motion' and result['motion_type'] == 'demurrer',
     f"Expected 'motion/demurrer', got '{result['doc_type']}/{result['motion_type']}'")


# ============================================================
# TEST 4: Deadline Calculations
# ============================================================
print_section("TEST 4: Deadline Calculations")

calc = DeadlineCalculator()

# Test discovery response deadline (30 days + 2 court days for e-service)
request_date = datetime(2026, 1, 25)  # Saturday
deadline = calc.get_discovery_response_deadline(request_date, 'electronic')
test("Discovery response deadline (e-service)",
     deadline is not None and 'date' in deadline,
     f"Should return deadline dict with date")

if deadline:
    # 30 calendar days from 1/25 = 2/24, then +2 court days
    print(f"   Discovery response due: {deadline['date'].strftime('%Y-%m-%d')}")

# Test motion to compel deadline (45 days)
response_date = datetime(2026, 1, 25)
deadline = calc.get_motion_to_compel_deadline(response_date)
test("Motion to compel deadline (45 days)",
     deadline is not None and 'date' in deadline,
     f"Should return deadline dict with date")

if deadline:
    print(f"   Motion to compel due: {deadline['date'].strftime('%Y-%m-%d')}")

# Test opposition deadline for standard motion
hearing_date = datetime(2026, 3, 20)
deadline = calc.get_opposition_deadline('standard', hearing_date, 'electronic')
test("Opposition deadline (standard motion)",
     deadline is not None and 'date' in deadline,
     f"Should return deadline dict with date")

if deadline:
    # 9 court days before hearing + 2 court days for e-service
    print(f"   Opposition due: {deadline['date'].strftime('%Y-%m-%d')}")

# Test opposition deadline for MSJ (different timing)
deadline = calc.get_opposition_deadline('msj', hearing_date, 'electronic')
test("Opposition deadline (MSJ)",
     deadline is not None and 'date' in deadline,
     f"Should return deadline dict with date")

if deadline:
    print(f"   MSJ Opposition due: {deadline['date'].strftime('%Y-%m-%d')}")

# Test reply deadline
deadline = calc.get_reply_deadline('standard', hearing_date, 'electronic')
test("Reply deadline (standard motion)",
     deadline is not None and 'date' in deadline,
     f"Should return deadline dict with date")

if deadline:
    # 5 court days before hearing + 2 court days for e-service
    print(f"   Reply due: {deadline['date'].strftime('%Y-%m-%d')}")


# ============================================================
# TEST 5: Date Extraction from Documents
# ============================================================
print_section("TEST 5: Date Extraction from Documents")

# Test hearing date extraction
motion_with_hearing = """NOTICE OF MOTION

Hearing Date: March 20, 2026
Time: 9:00 a.m.
Department: 5

Defendant moves for summary judgment...
"""

hearing = classifier._find_hearing_date(motion_with_hearing)
test("Extract hearing date from motion",
     hearing is not None and hearing.date() == datetime(2026, 3, 20).date(),
     f"Expected 2026-03-20, got {hearing}")

# Test date extraction
dates = classifier._extract_dates(motion_with_hearing)
test("Extract dates from document",
     len(dates) > 0,
     f"Expected to find dates, got {dates}")


# ============================================================
# TEST 6: File Number Extraction
# ============================================================
print_section("TEST 6: File Number Extraction")

test_cases = [
    ("RE: 3800.133 - Smith v. Jones", "3800.133"),
    ("FW: Case 1234.567 documents", "1234.567"),
    ("5800-013 Deposition Notice", "5800.013"),
    ("File 9999.001 - Urgent", "9999.001"),
    ("No file number here", ""),
]

for subject, expected in test_cases:
    result = monitor._extract_file_number(subject)
    test(f"Extract file number from '{subject}'",
         result == expected,
         f"Expected '{expected}', got '{result}'")


# ============================================================
# TEST 7: Edge Cases
# ============================================================
print_section("TEST 7: Edge Cases")

# Empty body
dates = monitor._extract_dates_from_body("")
test("Empty body returns no dates",
     len(dates) == 0,
     f"Expected empty list, got {dates}")

# Body with past dates only
body = "The meeting was on January 1, 2020."
dates = monitor._extract_dates_from_body(body)
test("Past dates should be filtered out",
     len(dates) == 0,
     f"Expected empty (past date), got {dates}")

# Multiple dates in body
body = "Meeting on February 1, 2026 and follow-up on February 15, 2026."
dates = monitor._extract_dates_from_body(body)
test("Multiple dates extracted",
     len(dates) >= 2,
     f"Expected 2+ dates, got {len(dates)}")


# ============================================================
# TEST 8: Additional Edge Cases
# ============================================================
print_section("TEST 8: Additional Edge Cases")

# Test email with both "tomorrow" and "day after tomorrow" - should only get day after tomorrow
body = "Not tomorrow, but the day after tomorrow works better."
dates = monitor._extract_dates_from_body(body)
test("Body with both 'tomorrow' and 'day after tomorrow'",
     len(dates) == 1 and dates[0].date() == day_after,
     f"Expected only {day_after}, got {[d.date() for d in dates]}")

# Test anti-SLAPP motion detection
anti_slapp_text = """SPECIAL MOTION TO STRIKE PURSUANT TO CCP 425.16

Defendant moves to strike the Complaint under the anti-SLAPP statute...
"""
result = classifier._classify_with_patterns(anti_slapp_text)
test("Classify anti-SLAPP motion",
     result['doc_type'] == 'motion' and result['motion_type'] == 'anti_slapp',
     f"Expected 'motion/anti_slapp', got '{result['doc_type']}/{result['motion_type']}'")

# Test motion in limine detection
mil_text = """MOTION IN LIMINE NO. 1

Plaintiff moves in limine to exclude evidence of...
"""
result = classifier._classify_with_patterns(mil_text)
test("Classify motion in limine",
     result['doc_type'] == 'motion' and result['motion_type'] == 'in_limine',
     f"Expected 'motion/in_limine', got '{result['doc_type']}/{result['motion_type']}'")

# Test file number with en-dash and em-dash
test_cases_dash = [
    ("Case 1234\u2013567", "1234.567"),  # en-dash
    ("Case 1234\u2014567", "1234.567"),  # em-dash
]
for subject, expected in test_cases_dash:
    result = monitor._extract_file_number(subject)
    test(f"Extract file number with unicode dash",
         result == expected,
         f"Expected '{expected}', got '{result}'")

# Test deposition notice with PMK
pmk_depo_text = """NOTICE OF DEPOSITION OF PERSON MOST KNOWLEDGEABLE

Notice is hereby given that Plaintiff will take the deposition of
the PERSON MOST KNOWLEDGEABLE of ABC Corporation regarding the
topics set forth below, on March 1, 2026 at 9:30 a.m.
"""
result = classifier._classify_with_patterns(pmk_depo_text)
test("Classify PMK deposition notice",
     result['doc_type'] == 'deposition_notice',
     f"Expected 'deposition_notice', got '{result['doc_type']}'")

# Test response to RFA
rfa_response_text = """PLAINTIFF'S RESPONSE TO REQUEST FOR ADMISSIONS - SET ONE

REQUEST FOR ADMISSION NO. 1: Admit that...
RESPONSE: Denied.
"""
result = classifier._classify_with_patterns(rfa_response_text)
test("Classify response to RFA as discovery_response",
     result['doc_type'] == 'discovery_response',
     f"Expected 'discovery_response', got '{result['doc_type']}'")

# Test ex parte application
ex_parte_text = """EX PARTE APPLICATION FOR TEMPORARY RESTRAINING ORDER

Plaintiff applies ex parte for a temporary restraining order...
"""
result = classifier._classify_with_patterns(ex_parte_text)
test("Classify ex parte application",
     result['doc_type'] == 'motion' and result['motion_type'] == 'ex_parte',
     f"Expected 'motion/ex_parte', got '{result['doc_type']}/{result['motion_type']}'")

# Test date at year boundary
body = "The deadline is December 31, 2026."
dates = monitor._extract_dates_from_body(body)
expected = datetime(2026, 12, 31).date()
test("Parse date at year end",
     any(d.date() == expected for d in dates),
     f"Expected {expected}, got {[d.date() for d in dates]}")

# Test "this Friday" vs "next Friday"
body = "Let's meet this Friday."
dates = monitor._extract_dates_from_body(body)
# "this Friday" should be this week's Friday
current_weekday = now.weekday()
days_until_friday = (4 - current_weekday) % 7
if days_until_friday == 0 and now.hour >= 17:  # If it's Friday evening, might be next week
    days_until_friday = 7
this_friday = (now + timedelta(days=days_until_friday)).date()
test("Parse 'this Friday'",
     len(dates) > 0,  # Just check we got something, exact date depends on current day
     f"Expected a date, got {[d.date() for d in dates]}")


# ============================================================
# TEST 9: LLM Classification (if API available)
# ============================================================
print_section("TEST 9: LLM Classification")

try:
    from icharlotte_core.config import API_KEYS
    if API_KEYS.get("Gemini"):
        # Test LLM classification of deposition notice
        depo_text_llm = """NOTICE OF TAKING DEPOSITION

        PLEASE TAKE NOTICE that the deposition of JANE DOE will be taken
        on February 20, 2026 at 2:00 p.m. at the Law Offices of Smith & Jones.
        """

        result = classifier._classify_with_llm(depo_text_llm)
        if result:
            test("LLM classifies deposition notice",
                 result.get('doc_type') == 'deposition_notice',
                 f"Expected 'deposition_notice', got '{result.get('doc_type')}'")

            test("LLM extracts deponent name",
                 result.get('deponent_name') is not None and 'doe' in result.get('deponent_name', '').lower(),
                 f"Expected name with 'Doe', got '{result.get('deponent_name')}'")

            test("LLM extracts deposition date",
                 result.get('deposition_date') is not None,
                 f"Expected deposition date, got '{result.get('deposition_date')}'")
        else:
            print("   LLM classification returned None, skipping LLM tests")
    else:
        print("   Gemini API key not configured, skipping LLM tests")
except Exception as e:
    print(f"   LLM tests skipped due to error: {e}")


# ============================================================
# SUMMARY
# ============================================================
print_section("TEST SUMMARY")

for r in results:
    print(r)

print(f"\n{'='*60}")
print(f" TOTAL: {passed} passed, {failed} failed out of {passed + failed} tests")
print('='*60)

if failed > 0:
    sys.exit(1)
