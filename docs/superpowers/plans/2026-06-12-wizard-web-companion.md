# Wizard Web Companion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A standalone FastAPI server on the desktop that lets an iPhone (over Tailscale) run the seven script-based wizard tasks: pick case → pick files → settings → run with progress → answer mid-run prompts → view the finished .docx.

**Architecture:** New `webcompanion/` package, fully Qt-free. A `JobManager` launches `python -u Scripts/<script>.py …` subprocesses and parses the existing stdout protocol (`PROGRESS:` / `AWAITING_INPUT:` / `OUTPUT:`). Two-phase tasks reuse the *existing* session-manager modules (`icharlotte_core.deposition.session_manager`, `icharlotte_core.med_chron.session_manager`) so session semantics stay identical to the desktop wizard. Server-rendered Jinja2 mobile pages, no SPA.

**Tech Stack:** FastAPI + uvicorn + Jinja2 (jinja2 already installed), `subprocess.Popen` + threads, pytest.

**Spec:** `docs/superpowers/specs/2026-06-12-wizard-web-companion-design.md`

---

## Investigation findings (verified against the codebase — trust these, don't re-derive)

**Per-task invocation catalogue:**

| task_id | Phase 1 argv | Pause | Phase 2 argv |
|---|---|---|---|
| `summarize_documents` | `Scripts/summarize.py <file>` | no | — |
| `summarize_discovery` | `Scripts/summarize_discovery.py <file>` | no | — |
| `medical_records` | `Scripts/med_record.py <file>` | no | — |
| `separate` | `Scripts/separate.py <file>` | no | — |
| `summarize_depositions` | `Scripts/summarize_deposition.py <file>` | yes | `… --phase=summary <session.json>` |
| `med_chron_analysis` | `Scripts/med_chron.py --phase=prep <file>` | yes | `… --phase=run <session.json>` |
| `depo_prep` | `Scripts/depo_prep.py --phase=analyze <config.json>` | yes | `… --phase=generate <session.json>` |

**Settings finding:** the desktop base `SettingsPage.to_dict()` is a placeholder — the four single-phase tasks have **no settings** beyond file choice. All real settings live in (a) depo_prep's pre-run `config.json` and (b) the post-phase-1 session edits of the three two-phase tasks. So "key settings per task" = one pre-run form (depo_prep) + three awaiting-input forms.

**Stdout protocol** (from `icharlotte_core/ui/wizard/runners/subprocess_worker.py`):
- `PROGRESS:<int>` or `PROGRESS:<int>:<message>`
- `AWAITING_INPUT:<session_path>` — captured but only acted on after the process exits with code 0
- `OUTPUT:<path>` — authoritative output declaration; fallback is an mtime-diff scan of `<case>/NOTES/AI Output/**/*.docx`
- anything else → status log line

**Deposition session JSON** (read fields): `deponent_name`, `deponent_type`, `deposition_date`, `topics` (list of dicts with `title`), `input_path`. User config written via `session_manager.update_user_config(session_path, cfg)` which sets `phase="ready_for_summary"`. The cfg dict (from `DepoSummaryConfigForm.build_user_config`, `icharlotte_core/ui/depo_summary_config_form.py:266-306`):
```python
{"selected_topics": [...], "added_topics": [...], "bullets_per_topic": 5,
 "deponent_label": "Deponent", "custom_rules": "", "cross_check_enabled": False,
 "context_doc_paths": [], "audience": "neutral", "audience_custom": "",
 "tone": "recitation", "tone_custom": ""}
```
Audience values: `neutral`, `pro_plaintiff`, `pro_defense`, `custom`. Tone values: `recitation`, `editorial`, `custom`. Bullets spinbox: range 1–15, default 5.

**Med chron session JSON** (read fields): `provider_name`, `narrative_missing`, `catalog` (list of `{id, label, …}`). User config: `{"selected_catalog_ids": [...], "custom_analyses": [{"label","instruction","context_files":[]}]}` via `icharlotte_core.med_chron.session_manager.update_user_config`.

**Depo prep:** pre-run `config.json` (written to a temp dir, passed as the positional arg — mirrors `DepoPrepSettingsPage._on_analyze_clicked`):
```python
{"deponent_name": "", "deponent_role": "", "deponent_sources": [abs paths],
 "context_sources": [], "style": "discovery", "free_text_notes": "",
 "per_topic_flags": {"strategic_note": bool, "source_facts": bool,
                     "impeachment_hook": bool, "objection_alts": bool},
 "case_root": "<case path>"}
```
Style keys: `discovery`, `lockdown`, `expert`, `friendly`. After phase 1, topics live in `topics.json` **next to** the session.json: `{"topics": [{"id","title","strategic_note","relevant_digest_refs":[], "default_checked": bool, "lawyer_added": bool}]}`. Phase 2 reads session.json + the mutated topics.json.

**Other facts:** `MasterCaseDatabase()` takes no args; `get_all_cases()` returns `list[dict]` with keys incl. `file_number`, `plaintiff_last_name`, `case_path`; `get_case(file_number)` returns one dict or None. `summarize.py` spawns detached children when given >1 file — always pass exactly one file per process (the JobManager runs multi-file jobs sequentially anyway). `jinja2` is installed; `fastapi`/`uvicorn` are NOT.

**Simplifications locked in:** two-phase tasks accept exactly **one** input file per job (matches desktop, which preps `files[0]`). Web depo_prep v1 uses picker files as `deponent_sources` and leaves `context_sources` empty. Mobile awaiting-forms expose 3 blank custom/added rows rather than dynamic add buttons.

**Commit hygiene:** the working tree has unrelated uncommitted changes. Every commit step lists exact paths — `git add` ONLY those paths, never `git add -A`.

---

## File structure

```
webcompanion/
├── __init__.py          # empty
├── protocol.py          # stdout-line parsing (pure functions)
├── jobs.py              # Job dataclass + JobStore (jobs.json persistence)
├── runner.py            # ScriptRunner: Popen + reader thread
├── task_defs.py         # 7 TaskDefs, argv builders, session bridging
├── cases.py             # case list + safe in-case path browsing
├── job_manager.py       # queue, lifecycle, two-phase pause/resume
├── server.py            # FastAPI app factory, routes, main()
└── templates/
    ├── base.html  home.html  case.html  picker.html
    ├── depo_prep_settings.html  job.html
    ├── awaiting_deposition.html  awaiting_med_chron.html
    └── awaiting_depo_prep.html
run_webcompanion.bat
tests/test_webcompanion/
├── __init__.py  test_protocol.py  test_jobs.py  test_runner.py
├── test_task_defs.py  test_cases.py  test_job_manager.py  test_server.py
```

Job persistence: `logs/webcompanion/jobs.json` (the `logs/` dir already hosts `depo_sessions/`).

---

### Task 1: Dependencies + package scaffold

**Files:**
- Modify: `E:\geminiterminal2\requirements.txt`
- Create: `webcompanion/__init__.py`, `tests/test_webcompanion/__init__.py`

- [ ] **Step 1: Install and pin dependencies**

Run: `pip install fastapi uvicorn python-multipart`
Expected: successful install (jinja2 already present).

Append to `requirements.txt` (after the `# Testing` block):

```
# Web companion (mobile wizard access)
fastapi>=0.115.0
uvicorn>=0.30.0
python-multipart>=0.0.9
```

- [ ] **Step 2: Create empty package files**

Create `webcompanion/__init__.py` and `tests/test_webcompanion/__init__.py`, both containing only:

```python
```

- [ ] **Step 3: Verify imports work**

Run: `python -c "import fastapi, uvicorn, multipart, webcompanion; print('ok')"`
Expected: `ok`

- [ ] **Step 4: Commit**

```bash
git add requirements.txt webcompanion/__init__.py tests/test_webcompanion/__init__.py
git commit -m "chore(webcompanion): add package scaffold and web dependencies"
```

---

### Task 2: Protocol parsing

**Files:**
- Create: `webcompanion/protocol.py`
- Test: `tests/test_webcompanion/test_protocol.py`

- [ ] **Step 1: Write the failing tests**

```python
"""Tests for webcompanion.protocol — wizard stdout line parsing."""
from webcompanion.protocol import parse_line


def test_progress_plain():
    p = parse_line("PROGRESS: 42")
    assert p.kind == "progress" and p.pct == 42 and p.message == ""


def test_progress_with_message():
    p = parse_line("PROGRESS:7:Extracting text")
    assert p.kind == "progress" and p.pct == 7 and p.message == "Extracting text"


def test_progress_clamped():
    assert parse_line("PROGRESS:150").pct == 100


def test_awaiting_input():
    p = parse_line(r"AWAITING_INPUT:C:\logs\depo_sessions\abc.json")
    assert p.kind == "awaiting" and p.path == r"C:\logs\depo_sessions\abc.json"


def test_output():
    p = parse_line(r"OUTPUT:E:\case\NOTES\AI Output\summary.docx")
    assert p.kind == "output" and p.path.endswith("summary.docx")


def test_status_fallthrough():
    p = parse_line("Loading model...")
    assert p.kind == "status" and p.message == "Loading model..."
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_protocol.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'webcompanion.protocol'`

- [ ] **Step 3: Implement**

`webcompanion/protocol.py`:

```python
"""Parsing for the wizard agent stdout protocol.

Same line grammar as icharlotte_core/ui/wizard/runners/subprocess_worker.py:
  PROGRESS:<int>[:<message>]   AWAITING_INPUT:<path>   OUTPUT:<path>
Anything else is a plain status line.
"""
import re
from dataclasses import dataclass
from typing import Optional

_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*(?::(.*))?$")
_AWAITING_RE = re.compile(r"^AWAITING_INPUT:(.+)$")
_OUTPUT_RE = re.compile(r"^OUTPUT:(.+)$")


@dataclass(frozen=True)
class ParsedLine:
    kind: str  # 'progress' | 'awaiting' | 'output' | 'status'
    pct: Optional[int] = None
    message: str = ""
    path: str = ""


def parse_line(line: str) -> ParsedLine:
    m = _PROGRESS_RE.match(line)
    if m:
        pct = max(0, min(100, int(m.group(1))))
        return ParsedLine(kind="progress", pct=pct, message=(m.group(2) or "").strip())
    m = _AWAITING_RE.match(line)
    if m:
        return ParsedLine(kind="awaiting", path=m.group(1).strip())
    m = _OUTPUT_RE.match(line)
    if m:
        return ParsedLine(kind="output", path=m.group(1).strip())
    return ParsedLine(kind="status", message=line)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_protocol.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```bash
git add webcompanion/protocol.py tests/test_webcompanion/test_protocol.py
git commit -m "feat(webcompanion): stdout protocol parser"
```

---

### Task 3: Job model + persistent store

**Files:**
- Create: `webcompanion/jobs.py`
- Test: `tests/test_webcompanion/test_jobs.py`

- [ ] **Step 1: Write the failing tests**

```python
"""Tests for webcompanion.jobs — Job model + JobStore persistence."""
import json

from webcompanion import jobs as J
from webcompanion.jobs import Job, JobStore, new_job


def _mk(task_id="summarize_documents"):
    return new_job(task_id, r"E:\cases\1234", "1234", [r"E:\cases\1234\doc.pdf"])


def test_new_job_defaults():
    job = _mk()
    assert job.state == J.QUEUED and job.progress == 0 and len(job.id) == 12


def test_log_capped():
    job = _mk()
    for i in range(250):
        job.add_log(f"line {i}")
    assert len(job.log) == 200 and job.log[-1] == "line 249"


def test_store_roundtrip(tmp_path):
    path = tmp_path / "jobs.json"
    store = JobStore(path)
    job = _mk()
    job.state = J.DONE
    store.add(job)
    store2 = JobStore(path)
    loaded = store2.get(job.id)
    assert loaded is not None and loaded.state == J.DONE
    assert loaded.files == job.files


def test_active_jobs_marked_interrupted_on_load(tmp_path):
    path = tmp_path / "jobs.json"
    store = JobStore(path)
    running, queued, awaiting = _mk(), _mk(), _mk()
    running.state = J.RUNNING
    queued.state = J.QUEUED
    awaiting.state = J.AWAITING_INPUT
    awaiting.session_path = r"C:\logs\s.json"
    for j in (running, queued, awaiting):
        store.add(j)
    store2 = JobStore(path)
    assert store2.get(running.id).state == J.INTERRUPTED
    assert store2.get(queued.id).state == J.INTERRUPTED
    # awaiting survives restarts — session file is on disk, phase 2 can run
    assert store2.get(awaiting.id).state == J.AWAITING_INPUT


def test_all_sorted_newest_first(tmp_path):
    store = JobStore(tmp_path / "jobs.json")
    a, b = _mk(), _mk()
    a.created_at, b.created_at = 100.0, 200.0
    store.add(a)
    store.add(b)
    assert [j.id for j in store.all()] == [b.id, a.id]


def test_corrupt_file_tolerated(tmp_path):
    path = tmp_path / "jobs.json"
    path.write_text("{not json", encoding="utf-8")
    store = JobStore(path)  # must not raise
    assert store.all() == []
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_jobs.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'webcompanion.jobs'`

- [ ] **Step 3: Implement**

`webcompanion/jobs.py`:

```python
"""Job model + JSON-file persistence for the wizard web companion."""
import json
import threading
import time
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

QUEUED = "queued"
RUNNING = "running"
AWAITING_INPUT = "awaiting_input"
DONE = "done"
FAILED = "failed"
CANCELLED = "cancelled"
INTERRUPTED = "interrupted"

ACTIVE_STATES = {QUEUED, RUNNING, AWAITING_INPUT}
_LOG_CAP = 200


@dataclass
class Job:
    id: str
    task_id: str
    case_path: str
    file_number: str
    files: List[str]
    state: str = QUEUED
    progress: int = 0
    log: List[str] = field(default_factory=list)
    session_path: str = ""
    output_path: str = ""
    error: str = ""
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)

    def add_log(self, line: str) -> None:
        self.log.append(line)
        if len(self.log) > _LOG_CAP:
            del self.log[: len(self.log) - _LOG_CAP]
        self.touch()

    def touch(self) -> None:
        self.updated_at = time.time()


def new_job(task_id: str, case_path: str, file_number: str, files: List[str]) -> Job:
    return Job(
        id=uuid.uuid4().hex[:12],
        task_id=task_id,
        case_path=case_path,
        file_number=file_number,
        files=list(files),
    )


class JobStore:
    """Thread-safe in-memory job table persisted to a JSON file.

    On load, jobs that were `queued` or `running` when the server died are
    marked `interrupted` (the child process died with the parent or is
    orphaned — we report honestly rather than pretend to reattach).
    `awaiting_input` jobs survive restarts: their session file is on disk
    and phase 2 can still be launched.
    """

    def __init__(self, path):
        self._path = Path(path)
        self._lock = threading.RLock()
        self._jobs: Dict[str, Job] = {}
        self._load()

    def _load(self) -> None:
        if not self._path.exists():
            return
        try:
            raw = json.loads(self._path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        for item in raw.get("jobs", []):
            job = Job(**item)
            if job.state in (QUEUED, RUNNING):
                job.state = INTERRUPTED
                job.error = job.error or "Server restarted while the job was active."
            self._jobs[job.id] = job
        self.save()

    def save(self) -> None:
        with self._lock:
            data = {"jobs": [asdict(j) for j in self._jobs.values()]}
            self._path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self._path.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
            tmp.replace(self._path)

    def add(self, job: Job) -> None:
        with self._lock:
            self._jobs[job.id] = job
            self.save()

    def get(self, job_id: str) -> Optional[Job]:
        with self._lock:
            return self._jobs.get(job_id)

    def all(self) -> List[Job]:
        with self._lock:
            return sorted(self._jobs.values(), key=lambda j: j.created_at, reverse=True)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_jobs.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```bash
git add webcompanion/jobs.py tests/test_webcompanion/test_jobs.py
git commit -m "feat(webcompanion): job model and persistent job store"
```

---

### Task 4: ScriptRunner (Popen + reader thread)

**Files:**
- Create: `webcompanion/runner.py`
- Test: `tests/test_webcompanion/test_runner.py`

- [ ] **Step 1: Write the failing tests** (fake protocol-emitting scripts in tmp_path)

```python
"""Tests for webcompanion.runner — subprocess driver."""
import textwrap
import time

from webcompanion.runner import ScriptRunner


def _write_script(tmp_path, body):
    p = tmp_path / "fake_agent.py"
    p.write_text(textwrap.dedent(body), encoding="utf-8")
    return str(p)


def _run_collect(argv, timeout=15.0):
    events, exits = [], []
    r = ScriptRunner(argv, on_event=events.append, on_exit=exits.append)
    r.start()
    deadline = time.time() + timeout
    while not exits and time.time() < deadline:
        time.sleep(0.05)
    assert exits, "script did not exit in time"
    return events, exits[0]


def test_events_and_clean_exit(tmp_path):
    script = _write_script(tmp_path, """
        print("PROGRESS:10:starting")
        print("hello status")
        print("OUTPUT:E:/out/result.docx")
    """)
    events, rc = _run_collect([script])
    assert rc == 0
    kinds = [e.kind for e in events]
    assert "progress" in kinds and "status" in kinds and "output" in kinds
    out = next(e for e in events if e.kind == "output")
    assert out.path == "E:/out/result.docx"


def test_nonzero_exit(tmp_path):
    script = _write_script(tmp_path, """
        import sys
        print("about to fail")
        sys.exit(3)
    """)
    events, rc = _run_collect([script])
    assert rc == 3
    assert any(e.message == "about to fail" for e in events)


def test_cancel_terminates(tmp_path):
    script = _write_script(tmp_path, """
        import time
        print("PROGRESS:1")
        time.sleep(60)
    """)
    exits = []
    r = ScriptRunner([script], on_event=lambda e: None, on_exit=exits.append)
    r.start()
    time.sleep(1.0)
    r.cancel()
    deadline = time.time() + 10
    while not exits and time.time() < deadline:
        time.sleep(0.05)
    assert exits and exits[0] != 0
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_runner.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'webcompanion.runner'`

- [ ] **Step 3: Implement**

`webcompanion/runner.py`:

```python
"""Qt-free subprocess driver speaking the wizard stdout protocol."""
import os
import subprocess
import sys
import threading
from typing import Callable, List, Optional

from .protocol import ParsedLine, parse_line

_CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


class ScriptRunner:
    """Run ``python -u <argv...>``; stream parsed stdout lines to a callback.

    on_event(ParsedLine) fires for every stdout line; on_exit(returncode)
    fires exactly once after EOF. Both run on the reader thread — callers
    must do their own locking.
    """

    def __init__(
        self,
        argv: List[str],
        on_event: Callable[[ParsedLine], None],
        on_exit: Callable[[int], None],
        cwd: Optional[str] = None,
    ):
        self._argv = list(argv)
        self._on_event = on_event
        self._on_exit = on_exit
        self._cwd = cwd
        self._proc: Optional[subprocess.Popen] = None

    def start(self) -> None:
        self._proc = subprocess.Popen(
            [sys.executable, "-u"] + self._argv,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=self._cwd,
            creationflags=_CREATE_NO_WINDOW,
        )
        threading.Thread(target=self._read_loop, daemon=True).start()

    def _read_loop(self) -> None:
        proc = self._proc
        for line in proc.stdout:
            self._on_event(parse_line(line.rstrip("\r\n")))
        rc = proc.wait()
        self._on_exit(rc)

    def cancel(self) -> None:
        """Terminate, then hard-kill after 2 s if still alive."""
        proc = self._proc
        if proc is None or proc.poll() is not None:
            return
        proc.terminate()

        def _kill():
            if proc.poll() is None:
                proc.kill()

        threading.Timer(2.0, _kill).start()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_runner.py -v`
Expected: 3 passed (allow ~10 s for the cancel test)

- [ ] **Step 5: Commit**

```bash
git add webcompanion/runner.py tests/test_webcompanion/test_runner.py
git commit -m "feat(webcompanion): subprocess script runner"
```

---

### Task 5: Task definitions + session bridging

**Files:**
- Create: `webcompanion/task_defs.py`
- Test: `tests/test_webcompanion/test_task_defs.py`

- [ ] **Step 1: Write the failing tests**

```python
"""Tests for webcompanion.task_defs."""
import json
from pathlib import Path

from webcompanion import task_defs as T


def test_seven_tasks_registered():
    assert set(T.TASKS) == {
        "summarize_documents", "summarize_discovery", "summarize_depositions",
        "depo_prep", "medical_records", "med_chron_analysis", "separate",
    }


def test_phase1_argv_plain():
    task = T.TASKS["summarize_documents"]
    argv = T.build_phase1_argv(task, r"E:\case\doc.pdf")
    assert argv[0].endswith("summarize.py") and argv[-1] == r"E:\case\doc.pdf"
    assert len(argv) == 2


def test_phase1_argv_med_chron_has_prep_flag():
    task = T.TASKS["med_chron_analysis"]
    argv = T.build_phase1_argv(task, r"E:\case\chron.docx")
    assert argv[1] == "--phase=prep"


def test_phase2_argv():
    task = T.TASKS["summarize_depositions"]
    argv = T.build_phase2_argv(task, r"C:\logs\s.json")
    assert argv[1] == "--phase=summary" and argv[2] == r"C:\logs\s.json"
    assert T.build_phase2_argv(T.TASKS["depo_prep"], "x")[1] == "--phase=generate"


def test_write_depo_prep_config_roundtrip():
    cfg = {"deponent_name": "Smith", "style": "lockdown"}
    path = T.write_depo_prep_config(cfg)
    loaded = json.loads(Path(path).read_text(encoding="utf-8"))
    assert loaded["deponent_name"] == "Smith" and Path(path).name == "config.json"


def test_depo_prep_topics_roundtrip(tmp_path):
    session = tmp_path / "session.json"
    session.write_text("{}", encoding="utf-8")
    topics = [{"id": "t01", "title": "Background", "strategic_note": "",
               "relevant_digest_refs": [], "default_checked": True,
               "lawyer_added": False}]
    T.write_depo_prep_topics(str(session), topics)
    assert T.read_depo_prep_topics(str(session)) == topics


def test_script_path_points_into_scripts_dir():
    p = Path(T.script_path("summarize.py"))
    assert p.parent.name == "Scripts" and p.exists()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_task_defs.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement**

`webcompanion/task_defs.py`:

```python
"""Web-companion task catalogue + script/session bridging.

Mirrors the script-based subset of icharlotte_core/ui/wizard/registry.py.
Session edits reuse the SAME session_manager modules the desktop forms use,
so semantics stay identical.
"""
import json
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List

REPO_ROOT = Path(__file__).resolve().parent.parent

PDF = (".pdf",)
DOCS = (".pdf", ".docx", ".doc", ".txt")


@dataclass(frozen=True)
class TaskDef:
    task_id: str
    title: str
    glyph: str
    script_name: str
    description: str
    default_folders: tuple = ()
    file_exts: tuple = PDF
    two_phase: bool = False
    phase1_args: tuple = ()
    phase2_flag: str = "--phase=summary"
    awaiting_kind: str = ""    # '' | 'deposition' | 'med_chron' | 'depo_prep'
    pre_settings: str = ""     # '' | 'depo_prep'


TASKS = {
    "summarize_documents": TaskDef(
        task_id="summarize_documents", title="Summarize Documents", glyph="\U0001F4C4",
        script_name="summarize.py",
        description="Concise summary of one or more case documents.",
        file_exts=DOCS),
    "summarize_discovery": TaskDef(
        task_id="summarize_discovery", title="Summarize Discovery", glyph="\U0001F4CB",
        script_name="summarize_discovery.py",
        description="Summarize discovery responses with structure and citations.",
        default_folders=("DISCOVERY/RESPONSES", "DISCOVERY"), file_exts=DOCS),
    "summarize_depositions": TaskDef(
        task_id="summarize_depositions", title="Summarize Depositions", glyph="\U0001F399",
        script_name="summarize_deposition.py",
        description="Structured deposition summary (you pick topics mid-run).",
        default_folders=("DISCOVERY/TRANSCRIPTS", "DISCOVERY"), file_exts=DOCS,
        two_phase=True, awaiting_kind="deposition"),
    "depo_prep": TaskDef(
        task_id="depo_prep", title="Depo Prep", glyph="❔",
        script_name="depo_prep.py",
        description="Deposition outline with questions grounded in case sources.",
        default_folders=("DISCOVERY", "PLEADINGS", "RECORDS"), file_exts=DOCS,
        two_phase=True, phase1_args=("--phase=analyze",),
        phase2_flag="--phase=generate", awaiting_kind="depo_prep",
        pre_settings="depo_prep"),
    "medical_records": TaskDef(
        task_id="medical_records", title="Medical Records Review", glyph="\U0001F3E5",
        script_name="med_record.py",
        description="Extract and summarize medical records into a chronology.",
        default_folders=("RECORDS",)),
    "med_chron_analysis": TaskDef(
        task_id="med_chron_analysis", title="Med Chron Analysis", glyph="\U0001FA7A",
        script_name="med_chron.py",
        description="Selectable analyses on a medical chronology (you pick mid-run).",
        default_folders=("RECORDS",), file_exts=DOCS,
        two_phase=True, phase1_args=("--phase=prep",),
        phase2_flag="--phase=run", awaiting_kind="med_chron"),
    "separate": TaskDef(
        task_id="separate", title="Separate Documents", glyph="\U0001F4D1",
        script_name="separate.py",
        description="Split a combined PDF into individually-named documents."),
}


def script_path(script_name: str) -> str:
    return str(REPO_ROOT / "Scripts" / script_name)


def build_phase1_argv(task: TaskDef, input_path: str) -> List[str]:
    return [script_path(task.script_name), *task.phase1_args, input_path]


def build_phase2_argv(task: TaskDef, session_path: str) -> List[str]:
    return [script_path(task.script_name), task.phase2_flag, session_path]


def read_session_json(session_path: str) -> dict:
    return json.loads(Path(session_path).read_text(encoding="utf-8"))


# ---- Session bridging (same modules the desktop forms use) ----

def apply_deposition_user_config(session_path: str, cfg: dict) -> None:
    from icharlotte_core.deposition import session_manager
    session_manager.update_user_config(Path(session_path), cfg)


def apply_med_chron_user_config(session_path: str, cfg: dict) -> None:
    from icharlotte_core.med_chron import session_manager
    session_manager.update_user_config(Path(session_path), cfg)


def read_depo_prep_topics(session_path: str) -> list:
    topics_path = Path(session_path).parent / "topics.json"
    return json.loads(topics_path.read_text(encoding="utf-8")).get("topics", [])


def write_depo_prep_topics(session_path: str, topics: list) -> None:
    topics_path = Path(session_path).parent / "topics.json"
    topics_path.write_text(json.dumps({"topics": topics}, indent=2), encoding="utf-8")


def write_depo_prep_config(cfg: dict) -> str:
    """Persist a depo-prep config.json to a temp dir; return its path.

    Mirrors DepoPrepSettingsPage._on_analyze_clicked() — the config path is
    the positional argv for ``depo_prep.py --phase=analyze``.
    """
    tmpdir = tempfile.mkdtemp(prefix="depo_prep_config_")
    cfg_path = Path(tmpdir) / "config.json"
    cfg_path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    return str(cfg_path)


# ---- Form option lists (mirror the desktop combos exactly) ----

DEPO_PREP_STYLES = [
    ("discovery", "Discovery / Fact-gathering"),
    ("lockdown", "Lock-down (leading admissions)"),
    ("expert", "Expert challenge (Daubert-style)"),
    ("friendly", "Friendly (own client prep)"),
]

DEPO_AUDIENCES = [
    ("neutral", "Neutral"),
    ("pro_plaintiff", "Plaintiff's Counsel"),
    ("pro_defense", "Defense Counsel"),
    ("custom", "Custom…"),
]

DEPO_TONES = [
    ("recitation", "Recitation (no editorializing)"),
    ("editorial", "Editorial (allow analysis)"),
    ("custom", "Custom…"),
]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_task_defs.py -v`
Expected: 7 passed

- [ ] **Step 5: Commit**

```bash
git add webcompanion/task_defs.py tests/test_webcompanion/test_task_defs.py
git commit -m "feat(webcompanion): task catalogue and session bridging"
```

---

### Task 6: Case access + safe path browsing

**Files:**
- Create: `webcompanion/cases.py`
- Test: `tests/test_webcompanion/test_cases.py`

- [ ] **Step 1: Write the failing tests**

```python
"""Tests for webcompanion.cases — path safety and browsing."""
import pytest

from webcompanion import cases


def test_safe_resolve_inside(tmp_path):
    (tmp_path / "DISCOVERY").mkdir()
    p = cases.safe_resolve(str(tmp_path), "DISCOVERY")
    assert p == (tmp_path / "DISCOVERY").resolve()


def test_safe_resolve_root_when_empty(tmp_path):
    assert cases.safe_resolve(str(tmp_path), "") == tmp_path.resolve()


def test_safe_resolve_rejects_traversal(tmp_path):
    with pytest.raises(ValueError):
        cases.safe_resolve(str(tmp_path), "..\\..\\Windows")
    with pytest.raises(ValueError):
        cases.safe_resolve(str(tmp_path), "../etc")


def test_browse_filters_extensions(tmp_path):
    (tmp_path / "sub").mkdir()
    (tmp_path / "a.pdf").write_bytes(b"x")
    (tmp_path / "b.docx").write_bytes(b"x")
    (tmp_path / "c.exe").write_bytes(b"x")
    dirs, files = cases.browse(str(tmp_path), "", (".pdf",))
    assert dirs == ["sub"] and files == ["a.pdf"]
    _, files2 = cases.browse(str(tmp_path), "", (".pdf", ".docx"))
    assert files2 == ["a.pdf", "b.docx"]


def test_browse_missing_dir_returns_empty(tmp_path):
    dirs, files = cases.browse(str(tmp_path), "NOPE", (".pdf",))
    assert dirs == [] and files == []


def test_resolve_start_folder_first_existing(tmp_path):
    (tmp_path / "DISCOVERY").mkdir()
    rel = cases.resolve_start_folder(
        str(tmp_path), ("DISCOVERY/RESPONSES", "DISCOVERY"))
    assert rel == "DISCOVERY"


def test_resolve_start_folder_falls_back_to_root(tmp_path):
    assert cases.resolve_start_folder(str(tmp_path), ("NOPE",)) == ""
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_cases.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement**

`webcompanion/cases.py`:

```python
"""Case listing (master DB) + traversal-safe browsing inside a case folder."""
import os
from pathlib import Path
from typing import List, Optional, Tuple


def list_cases(query: str = "") -> List[dict]:
    from icharlotte_core.master_db import MasterCaseDatabase
    rows = MasterCaseDatabase().get_all_cases()
    q = (query or "").strip().lower()
    if not q:
        return rows
    return [
        r for r in rows
        if q in str(r.get("file_number", "")).lower()
        or q in str(r.get("plaintiff_last_name", "")).lower()
    ]


def get_case(file_number: str) -> Optional[dict]:
    from icharlotte_core.master_db import MasterCaseDatabase
    return MasterCaseDatabase().get_case(file_number)


def safe_resolve(case_root: str, rel: str) -> Path:
    """Resolve rel under case_root; raise ValueError if it escapes the root."""
    root = Path(case_root).resolve()
    target = (root / rel).resolve() if rel else root
    if target != root and root not in target.parents:
        raise ValueError("Path escapes the case folder.")
    return target


def browse(case_root: str, rel: str, exts: tuple) -> Tuple[List[str], List[str]]:
    """Return (subdir_names, file_names) under case_root/rel, ext-filtered."""
    target = safe_resolve(case_root, rel)
    dirs: List[str] = []
    files: List[str] = []
    if not target.is_dir():
        return dirs, files
    for entry in sorted(os.listdir(target), key=str.lower):
        p = target / entry
        if p.is_dir():
            dirs.append(entry)
        elif p.suffix.lower() in exts:
            files.append(entry)
    return dirs, files


def resolve_start_folder(case_root: str, default_folders: tuple) -> str:
    """First default folder that exists under the case root, else '' (root)."""
    for rel in default_folders:
        try:
            if safe_resolve(case_root, rel).is_dir():
                return rel
        except ValueError:
            continue
    return ""
```

NOTE for executor: if `MasterCaseDatabase.get_case()` returns a sqlite Row
or tuple instead of a dict, read `icharlotte_core/master_db.py:142-150` and
adapt `get_case` here to return a plain dict (`dict(row)`), keeping the
public contract "dict or None".

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_cases.py -v`
Expected: 7 passed

- [ ] **Step 5: Commit**

```bash
git add webcompanion/cases.py tests/test_webcompanion/test_cases.py
git commit -m "feat(webcompanion): case access and safe folder browsing"
```

---

### Task 7: JobManager (queue, lifecycle, two-phase)

**Files:**
- Create: `webcompanion/job_manager.py`
- Test: `tests/test_webcompanion/test_job_manager.py`

- [ ] **Step 1: Write the failing tests**

The tests register a fake task (monkeypatched into `TASKS`) whose "script"
is a temp .py file, so no real agents run.

```python
"""Tests for webcompanion.job_manager."""
import textwrap
import time

import pytest

from webcompanion import jobs as J
from webcompanion import task_defs as T
from webcompanion.job_manager import JobManager
from webcompanion.jobs import JobStore, new_job
from webcompanion.task_defs import TaskDef


def _wait(cond, timeout=20.0):
    deadline = time.time() + timeout
    while not cond() and time.time() < deadline:
        time.sleep(0.05)
    assert cond(), "condition not met in time"


@pytest.fixture
def fake_task(tmp_path, monkeypatch):
    """Register task 'fake' whose script is a stub we control via env-free args.

    The stub script reads its positional arg (a .txt 'mode file') whose first
    line selects behavior: ok | fail | await.
    """
    script = tmp_path / "fake_script.py"
    script.write_text(textwrap.dedent("""
        import sys
        from pathlib import Path
        arg = sys.argv[-1]
        if sys.argv[1].startswith("--phase=resume"):
            print("PROGRESS:90:phase2")
            print("OUTPUT:" + str(Path(arg).with_name("phase2.docx")))
            sys.exit(0)
        mode = Path(arg).read_text(encoding="utf-8").strip()
        print("PROGRESS:10:working")
        if mode == "fail":
            print("boom")
            sys.exit(2)
        if mode == "await":
            session = Path(arg).with_name("session.json")
            session.write_text("{}", encoding="utf-8")
            print("AWAITING_INPUT:" + str(session))
            sys.exit(0)
        print("OUTPUT:" + str(Path(arg).with_name("result.docx")))
        sys.exit(0)
    """), encoding="utf-8")

    spec = TaskDef(task_id="fake", title="Fake", glyph="F",
                   script_name="UNUSED", description="test task",
                   two_phase=True, phase2_flag="--phase=resume")
    monkeypatch.setitem(T.TASKS, "fake", spec)
    monkeypatch.setattr(T, "script_path", lambda name: str(script))
    # job_manager imported build_* from task_defs; patch there too
    import webcompanion.job_manager as jm
    monkeypatch.setattr(
        jm, "build_phase1_argv",
        lambda task, p: [str(script), *task.phase1_args, p])
    monkeypatch.setattr(
        jm, "build_phase2_argv",
        lambda task, s: [str(script), task.phase2_flag, s])
    return script


def _submit(manager, tmp_path, mode, name="mode.txt"):
    mode_file = tmp_path / name
    mode_file.write_text(mode, encoding="utf-8")
    job = new_job("fake", str(tmp_path), "9999", [str(mode_file)])
    return manager.submit(job)


def test_success_with_explicit_output(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    job = _submit(manager, tmp_path, "ok")
    _wait(lambda: manager.store.get(job.id).state == J.DONE)
    final = manager.store.get(job.id)
    assert final.output_path.endswith("result.docx")
    assert final.progress == 100


def test_failure_marks_failed(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    job = _submit(manager, tmp_path, "fail")
    _wait(lambda: manager.store.get(job.id).state == J.FAILED)
    final = manager.store.get(job.id)
    assert "code 2" in final.error
    assert any("boom" in ln for ln in final.log)


def test_awaiting_then_resume(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    job = _submit(manager, tmp_path, "await")
    _wait(lambda: manager.store.get(job.id).state == J.AWAITING_INPUT)
    mid = manager.store.get(job.id)
    assert mid.session_path.endswith("session.json")
    manager.resume(job.id)
    _wait(lambda: manager.store.get(job.id).state == J.DONE)
    assert manager.store.get(job.id).output_path.endswith("phase2.docx")


def test_per_case_serialization(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    a = _submit(manager, tmp_path, "ok", "a.txt")
    b = _submit(manager, tmp_path, "ok", "b.txt")  # same case_path → queued
    # b must not run while a runs; both eventually done
    _wait(lambda: manager.store.get(a.id).state == J.DONE)
    _wait(lambda: manager.store.get(b.id).state == J.DONE)


def test_two_phase_rejects_multiple_files(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    f1 = tmp_path / "x.txt"; f1.write_text("ok", encoding="utf-8")
    f2 = tmp_path / "y.txt"; f2.write_text("ok", encoding="utf-8")
    with pytest.raises(ValueError):
        manager.submit(new_job("fake", str(tmp_path), "9999",
                               [str(f1), str(f2)]))


def test_cancel_queued_job(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"), max_concurrent=1)
    a = _submit(manager, tmp_path, "await", "a.txt")   # occupies the slot
    _wait(lambda: manager.store.get(a.id).state == J.AWAITING_INPUT)
    b = _submit(manager, tmp_path, "ok", "b.txt")
    # a is awaiting (not RUNNING) so b may start; cancel a instead
    manager.cancel(a.id)
    assert manager.store.get(a.id).state == J.CANCELLED
    _wait(lambda: manager.store.get(b.id).state == J.DONE)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_job_manager.py -v`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Implement**

`webcompanion/job_manager.py`:

```python
"""Job lifecycle: queue, launch, protocol handling, two-phase pause/resume.

Mirrors the semantics of icharlotte_core/ui/wizard/runners/subprocess_worker.py:
multi-file sequential runs with scaled progress, AWAITING_INPUT honored only
on clean exit, OUTPUT: line authoritative with an mtime-diff scan of
NOTES/AI Output as fallback, terminate-then-kill cancellation.
"""
import os
import threading
from pathlib import Path
from typing import Dict, Optional, Set

from . import jobs as J
from .jobs import Job, JobStore
from .protocol import ParsedLine
from .runner import ScriptRunner
from .task_defs import REPO_ROOT, TASKS, build_phase1_argv, build_phase2_argv


class JobManager:
    def __init__(self, store: JobStore, max_concurrent: int = 2):
        self.store = store
        self._max = max_concurrent
        self._lock = threading.RLock()
        self._runners: Dict[str, ScriptRunner] = {}
        self._file_idx: Dict[str, int] = {}
        self._snapshots: Dict[str, dict] = {}
        self._awaiting: Dict[str, str] = {}
        self._explicit_output: Dict[str, str] = {}
        self._cancel_requested: Set[str] = set()

    # ---- public API ----

    def submit(self, job: Job) -> Job:
        if job.task_id not in TASKS:
            raise KeyError(f"Unknown task: {job.task_id}")
        if not job.files:
            raise ValueError("Job has no input files.")
        task = TASKS[job.task_id]
        if task.two_phase and len(job.files) != 1:
            raise ValueError("Two-phase tasks accept exactly one input file.")
        with self._lock:
            self.store.add(job)
            self._schedule()
        return job

    def resume(self, job_id: str) -> None:
        """Launch phase 2 for an awaiting job (caller already edited the session)."""
        with self._lock:
            job = self.store.get(job_id)
            if job is None or job.state != J.AWAITING_INPUT or not job.session_path:
                raise ValueError("Job is not awaiting input.")
            task = TASKS[job.task_id]
            job.state = J.RUNNING
            job.add_log("Resuming Phase 2…")
            # Re-snapshot so phase 2's .docx is detected by the diff fallback.
            self._snapshots[job.id] = self._scan_outputs(job.case_path)
            self._explicit_output.pop(job.id, None)
            self.store.save()
            self._launch(job, build_phase2_argv(task, job.session_path))

    def cancel(self, job_id: str) -> None:
        with self._lock:
            job = self.store.get(job_id)
            if job is None:
                return
            if job.state in (J.QUEUED, J.AWAITING_INPUT):
                job.state = J.CANCELLED
                job.add_log("Cancelled.")
                self.store.save()
                self._schedule()
            elif job.state == J.RUNNING:
                self._cancel_requested.add(job_id)
                runner = self._runners.get(job_id)
                if runner is not None:
                    runner.cancel()

    # ---- scheduling ----

    def _running_case_paths(self) -> Set[str]:
        return {j.case_path for j in self.store.all() if j.state == J.RUNNING}

    def _running_count(self) -> int:
        return sum(1 for j in self.store.all() if j.state == J.RUNNING)

    def _schedule(self) -> None:
        # store.all() is newest-first; iterate oldest-first for FIFO.
        for job in reversed(self.store.all()):
            if job.state != J.QUEUED:
                continue
            if self._running_count() >= self._max:
                return
            if job.case_path in self._running_case_paths():
                continue
            self._start_current_file(job)

    def _start_current_file(self, job: Job) -> None:
        task = TASKS[job.task_id]
        idx = self._file_idx.get(job.id, 0)
        job.state = J.RUNNING
        job.add_log(f"Starting {os.path.basename(job.files[idx])}…")
        self._snapshots[job.id] = self._scan_outputs(job.case_path)
        self._awaiting.pop(job.id, None)
        self._explicit_output.pop(job.id, None)
        self.store.save()
        self._launch(job, build_phase1_argv(task, job.files[idx]))

    def _launch(self, job: Job, argv) -> None:
        runner = ScriptRunner(
            argv,
            on_event=lambda ev, jid=job.id: self._on_event(jid, ev),
            on_exit=lambda rc, jid=job.id: self._on_exit(jid, rc),
            cwd=str(REPO_ROOT),
        )
        self._runners[job.id] = runner
        runner.start()

    # ---- events (called on runner reader threads) ----

    def _scaled_pct(self, job: Job, raw: int) -> int:
        total = max(1, len(job.files))
        idx = self._file_idx.get(job.id, 0)
        return max(0, min(100, int((idx * 100 + raw) / total)))

    def _on_event(self, job_id: str, ev: ParsedLine) -> None:
        with self._lock:
            job = self.store.get(job_id)
            if job is None:
                return
            if ev.kind == "progress":
                job.progress = self._scaled_pct(job, ev.pct or 0)
                if ev.message:
                    job.add_log(ev.message)
                else:
                    job.touch()
            elif ev.kind == "awaiting":
                self._awaiting[job_id] = ev.path
            elif ev.kind == "output":
                self._explicit_output[job_id] = ev.path
            elif ev.message.strip():
                job.add_log(ev.message)
            self.store.save()

    def _on_exit(self, job_id: str, returncode: int) -> None:
        with self._lock:
            self._runners.pop(job_id, None)
            job = self.store.get(job_id)
            if job is None:
                return
            if job_id in self._cancel_requested:
                self._cancel_requested.discard(job_id)
                job.state = J.CANCELLED
                job.add_log("Cancelled.")
            elif returncode != 0:
                job.state = J.FAILED
                job.error = f"Script exited with code {returncode}."
                job.add_log(job.error)
            elif self._awaiting.get(job_id):
                job.state = J.AWAITING_INPUT
                job.session_path = self._awaiting.pop(job_id)
                job.add_log("Awaiting your input.")
            else:
                output = self._explicit_output.pop(job_id, None) or \
                    self._diff_outputs(job.case_path,
                                       self._snapshots.pop(job_id, {}))
                if output:
                    job.output_path = output
                next_idx = self._file_idx.get(job_id, 0) + 1
                if next_idx < len(job.files):
                    self._file_idx[job_id] = next_idx
                    self.store.save()
                    self._start_current_file(job)
                    return
                job.state = J.DONE
                job.progress = 100
                job.add_log(
                    "Done." if job.output_path
                    else "Done (no .docx detected — check NOTES/AI Output).")
            self.store.save()
            self._schedule()

    # ---- output detection ----

    def _scan_outputs(self, case_path: str) -> dict:
        out_dir = Path(case_path) / "NOTES" / "AI Output"
        if not out_dir.is_dir():
            return {}
        result = {}
        for p in out_dir.rglob("*.docx"):
            try:
                result[str(p)] = p.stat().st_mtime
            except OSError:
                continue
        return result

    def _diff_outputs(self, case_path: str, before: dict) -> Optional[str]:
        after = self._scan_outputs(case_path)
        changed = [
            (mtime, path) for path, mtime in after.items()
            if path not in before or mtime > before[path]
        ]
        if not changed:
            return None
        return max(changed)[1]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_job_manager.py -v`
Expected: 6 passed (allow up to ~60 s; subprocess startup on Windows is slow)

- [ ] **Step 5: Run the full webcompanion suite for regressions**

Run: `python -m pytest tests/test_webcompanion/ -v`
Expected: all pass

- [ ] **Step 6: Commit**

```bash
git add webcompanion/job_manager.py tests/test_webcompanion/test_job_manager.py
git commit -m "feat(webcompanion): job manager with queue and two-phase resume"
```

---

### Task 8: Server core — app factory, home/case/picker/start, templates

**Files:**
- Create: `webcompanion/server.py`, `webcompanion/templates/base.html`, `home.html`, `case.html`, `picker.html`
- Test: `tests/test_webcompanion/test_server.py`

- [ ] **Step 1: Write the failing tests**

```python
"""Endpoint tests for the web companion server."""
import textwrap

import pytest
from fastapi.testclient import TestClient

from webcompanion import cases as cases_mod
from webcompanion import jobs as J
from webcompanion import task_defs as T
from webcompanion.job_manager import JobManager
from webcompanion.jobs import JobStore
from webcompanion.server import create_app


@pytest.fixture
def case_dir(tmp_path):
    root = tmp_path / "case_9999"
    (root / "DISCOVERY").mkdir(parents=True)
    (root / "DISCOVERY" / "resp.pdf").write_bytes(b"x")
    (root / "doc.pdf").write_bytes(b"x")
    return root


@pytest.fixture
def client(tmp_path, case_dir, monkeypatch):
    fake_case = {"file_number": "9999", "plaintiff_last_name": "Smith",
                 "case_path": str(case_dir)}
    monkeypatch.setattr(cases_mod, "list_cases", lambda q="": [fake_case])
    monkeypatch.setattr(cases_mod, "get_case",
                        lambda fn: fake_case if fn == "9999" else None)
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    app = create_app(manager)
    c = TestClient(app)
    c.manager = manager
    return c


def test_home_lists_cases(client):
    r = client.get("/")
    assert r.status_code == 200 and "9999" in r.text and "Smith" in r.text


def test_case_page_shows_task_cards(client):
    r = client.get("/case/9999")
    assert r.status_code == 200
    assert "Summarize Documents" in r.text and "Depo Prep" in r.text


def test_case_page_404(client):
    assert client.get("/case/0000").status_code == 404


def test_picker_lists_dirs_and_files(client):
    r = client.get("/case/9999/task/summarize_documents")
    assert r.status_code == 200
    assert "DISCOVERY" in r.text and "doc.pdf" in r.text


def test_picker_rejects_traversal(client):
    r = client.get("/case/9999/task/summarize_documents",
                   params={"path": "../.."})
    assert r.status_code == 400


def test_start_requires_files(client):
    r = client.post("/case/9999/task/summarize_documents/start",
                    data={}, follow_redirects=False)
    assert r.status_code == 400


def test_start_submits_job_and_redirects(client, monkeypatch):
    # Don't actually launch a subprocess.
    submitted = {}
    monkeypatch.setattr(client.manager, "submit",
                        lambda job: submitted.setdefault("job", job) or job)
    r = client.post("/case/9999/task/summarize_documents/start",
                    data={"files": ["doc.pdf"]}, follow_redirects=False)
    assert r.status_code == 303
    job = submitted["job"]
    assert job.task_id == "summarize_documents"
    assert job.files[0].endswith("doc.pdf")
    assert r.headers["location"] == f"/job/{job.id}"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'webcompanion.server'`

- [ ] **Step 3: Create the templates**

`webcompanion/templates/base.html`:

```html
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>iCharlotte Companion</title>
<style>
 body{font-family:-apple-system,system-ui,sans-serif;margin:0;padding:12px;background:#f5f6f8;color:#1c1e21}
 a{color:#1565c0;text-decoration:none}
 .card{background:#fff;border-radius:10px;padding:12px;margin:8px 0;box-shadow:0 1px 2px rgba(0,0,0,.08)}
 .btn{display:inline-block;background:#1565c0;color:#fff;border:none;border-radius:8px;padding:10px 16px;font-size:16px;margin-top:8px}
 .btn.secondary{background:#666}
 input[type=text],input[type=number],textarea,select{width:100%;box-sizing:border-box;padding:8px;font-size:16px;border:1px solid #ccc;border-radius:6px;margin:4px 0}
 label{display:block;margin-top:8px}
 .row{display:flex;align-items:center;gap:8px}
 progress{width:100%;height:14px}
 pre{background:#23272e;color:#d7dae0;padding:8px;border-radius:8px;font-size:12px;overflow-x:auto;white-space:pre-wrap}
 h1{font-size:20px}h2{font-size:17px}
 .state{font-weight:600}
 .crumb{margin:4px 0;font-size:14px}
</style>
</head>
<body>
<div class="crumb"><a href="/">&#127968; Home</a></div>
{% block content %}{% endblock %}
</body>
</html>
```

`webcompanion/templates/home.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>iCharlotte Companion</h1>
{% if jobs %}
<h2>Jobs</h2>
{% for job in jobs %}
<div class="card"><a href="/job/{{ job.id }}">
 <b>{{ tasks[job.task_id].title if job.task_id in tasks else job.task_id }}</b>
 &mdash; {{ job.file_number }}<br>
 <span class="state">{{ job.state }}</span> {{ job.progress }}%
</a></div>
{% endfor %}
{% endif %}
<h2>Cases</h2>
<form method="get" action="/">
 <input type="text" name="q" value="{{ q }}" placeholder="Search file number or plaintiff" autocomplete="off">
</form>
{% for case in cases %}
<div class="card"><a href="/case/{{ case.file_number }}">
 <b>{{ case.file_number }}</b> &mdash; {{ case.plaintiff_last_name }}
</a></div>
{% endfor %}
{% endblock %}
```

`webcompanion/templates/case.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>{{ case.file_number }} &mdash; {{ case.plaintiff_last_name }}</h1>
{% for task in tasks %}
<div class="card"><a href="/case/{{ case.file_number }}/task/{{ task.task_id }}">
 {{ task.glyph }} <b>{{ task.title }}</b><br>
 <small>{{ task.description }}</small>
</a></div>
{% endfor %}
{% endblock %}
```

`webcompanion/templates/picker.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>{{ task.glyph }} {{ task.title }}</h1>
<div class="crumb">
 <a href="/case/{{ case.file_number }}/task/{{ task.task_id }}?path=">{{ case.file_number }}</a>
 {% for label, crumb_rel in crumbs %} / <a href="/case/{{ case.file_number }}/task/{{ task.task_id }}?path={{ crumb_rel | urlencode }}">{{ label }}</a>{% endfor %}
</div>
{% for d in dirs %}
<div class="card"><a href="/case/{{ case.file_number }}/task/{{ task.task_id }}?path={{ ((path ~ '/' ~ d) if path else d) | urlencode }}">&#128193; {{ d }}</a></div>
{% endfor %}
<form method="post" action="/case/{{ case.file_number }}/task/{{ task.task_id }}/start">
{% for f in files %}
<div class="card"><label class="row">
 <input type="checkbox" name="files" value="{{ (path ~ '/' ~ f) if path else f }}"> {{ f }}
</label></div>
{% endfor %}
{% if files %}
<button class="btn" type="submit">{% if task.pre_settings %}Continue&hellip;{% else %}Run {{ task.title }}{% endif %}</button>
{% else %}
<p>No matching files in this folder.</p>
{% endif %}
</form>
{% endblock %}
```

- [ ] **Step 4: Implement the server core**

`webcompanion/server.py`:

```python
"""FastAPI app for the wizard web companion.

Entry point: ``python -m webcompanion.server`` (binds the Tailscale IP by
default; ``--lan`` binds 0.0.0.0 for same-network development).
"""
import argparse
import subprocess
import sys
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import (FileResponse, HTMLResponse, JSONResponse,
                               RedirectResponse)
from fastapi.templating import Jinja2Templates

from . import cases
from . import jobs as J
from . import task_defs as T
from .job_manager import JobManager
from .jobs import JobStore, new_job

_TEMPLATES_DIR = Path(__file__).parent / "templates"
DEFAULT_JOBS_PATH = T.REPO_ROOT / "logs" / "webcompanion" / "jobs.json"
_DOCX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


def _crumbs(path: str):
    """[('DISCOVERY', 'DISCOVERY'), ('RESPONSES', 'DISCOVERY/RESPONSES')]"""
    out, acc = [], []
    for part in (path or "").replace("\\", "/").split("/"):
        if not part:
            continue
        acc.append(part)
        out.append((part, "/".join(acc)))
    return out


def create_app(manager: JobManager) -> FastAPI:
    app = FastAPI(title="iCharlotte Web Companion")
    templates = Jinja2Templates(directory=str(_TEMPLATES_DIR))

    def _render(name, request, **ctx):
        return templates.TemplateResponse(name, {"request": request, **ctx})

    # ---- home / case / picker ----

    @app.get("/", response_class=HTMLResponse)
    def home(request: Request, q: str = ""):
        return _render("home.html", request, jobs=manager.store.all()[:20],
                       cases=cases.list_cases(q), q=q, tasks=T.TASKS)

    @app.get("/case/{file_number}", response_class=HTMLResponse)
    def case_page(request: Request, file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        return _render("case.html", request, case=case,
                       tasks=list(T.TASKS.values()))

    @app.get("/case/{file_number}/task/{task_id}", response_class=HTMLResponse)
    def picker(request: Request, file_number: str, task_id: str,
               path: str = None):
        case = cases.get_case(file_number)
        if case is None or task_id not in T.TASKS:
            return HTMLResponse("Not found", status_code=404)
        task = T.TASKS[task_id]
        if path is None:
            path = cases.resolve_start_folder(case["case_path"],
                                              task.default_folders)
        try:
            dirs, files = cases.browse(case["case_path"], path, task.file_exts)
        except ValueError:
            return HTMLResponse("Invalid path", status_code=400)
        return _render("picker.html", request, case=case, task=task,
                       path=path, dirs=dirs, files=files,
                       crumbs=_crumbs(path))

    @app.post("/case/{file_number}/task/{task_id}/start")
    async def start(request: Request, file_number: str, task_id: str):
        case = cases.get_case(file_number)
        if case is None or task_id not in T.TASKS:
            return HTMLResponse("Not found", status_code=404)
        task = T.TASKS[task_id]
        form = await request.form()
        rel_files = [f for f in form.getlist("files") if f]
        if not rel_files:
            return HTMLResponse("Pick at least one file.", status_code=400)
        try:
            abs_files = [str(cases.safe_resolve(case["case_path"], rf))
                         for rf in rel_files]
        except ValueError:
            return HTMLResponse("Invalid path", status_code=400)
        if task.two_phase and len(abs_files) > 1:
            return HTMLResponse(
                "This task accepts exactly one input file.", status_code=400)
        if task.pre_settings == "depo_prep":
            return _render("depo_prep_settings.html", request, case=case,
                           task=task, rel_files=rel_files,
                           styles=T.DEPO_PREP_STYLES)
        job = new_job(task_id, case["case_path"], file_number, abs_files)
        manager.submit(job)
        return RedirectResponse(f"/job/{job.id}", status_code=303)

    _register_job_routes(app, manager, templates)
    _register_depo_prep_routes(app, manager, templates)
    _register_awaiting_routes(app, manager, templates)
    return app


# Filled in by Tasks 9-11. Keep these stubs so Task 8 imports cleanly.
def _register_job_routes(app, manager, templates):
    pass


def _register_depo_prep_routes(app, manager, templates):
    pass


def _register_awaiting_routes(app, manager, templates):
    pass


# ---- entry point ----

def detect_tailscale_ip() -> str | None:
    try:
        out = subprocess.run(["tailscale", "ip", "-4"],
                             capture_output=True, text=True, timeout=10)
        if out.returncode == 0:
            lines = [ln.strip() for ln in out.stdout.splitlines() if ln.strip()]
            if lines:
                return lines[0]
    except (OSError, subprocess.TimeoutExpired):
        pass
    return None


def main() -> None:
    parser = argparse.ArgumentParser(description="iCharlotte web companion")
    parser.add_argument("--lan", action="store_true",
                        help="Bind 0.0.0.0 instead of the Tailscale IP")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    if args.lan:
        host = "0.0.0.0"
    else:
        host = detect_tailscale_ip()
        if not host:
            print("ERROR: Could not detect a Tailscale IPv4 address. "
                  "Is Tailscale installed and running? "
                  "Use --lan to bind the local network instead.")
            sys.exit(1)

    app = create_app(JobManager(JobStore(DEFAULT_JOBS_PATH)))
    import uvicorn
    print(f"iCharlotte web companion: http://{host}:{args.port}")
    uvicorn.run(app, host=host, port=args.port)


if __name__ == "__main__":
    main()
```

Also create an empty-for-now `webcompanion/templates/depo_prep_settings.html` placeholder is NOT allowed — instead, Task 8 tests don't touch depo_prep settings, and the template is created in Task 10. To keep `start()` safe until then, the depo_prep branch renders a template that doesn't exist yet — acceptable because Task 10 immediately follows; do not ship between tasks.

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v`
Expected: 7 passed

- [ ] **Step 6: Commit**

```bash
git add webcompanion/server.py webcompanion/templates/base.html webcompanion/templates/home.html webcompanion/templates/case.html webcompanion/templates/picker.html tests/test_webcompanion/test_server.py
git commit -m "feat(webcompanion): server core with case browsing and job start"
```

---

### Task 9: Job pages — status API, run page, cancel, output download

**Files:**
- Modify: `webcompanion/server.py` (replace the `_register_job_routes` stub)
- Create: `webcompanion/templates/job.html`
- Test: append to `tests/test_webcompanion/test_server.py`

- [ ] **Step 1: Write the failing tests** (append to test_server.py)

```python
def _make_job(client, state=J.RUNNING, **kw):
    from webcompanion.jobs import new_job
    job = new_job("summarize_documents", "E:/case", "9999", ["E:/case/d.pdf"])
    job.state = state
    for k, v in kw.items():
        setattr(job, k, v)
    client.manager.store.add(job)
    return job


def test_job_page_renders(client):
    job = _make_job(client, progress=40)
    r = client.get(f"/job/{job.id}")
    assert r.status_code == 200 and "Summarize Documents" in r.text


def test_job_page_404(client):
    assert client.get("/job/nope").status_code == 404


def test_job_state_api(client):
    job = _make_job(client, progress=55)
    job.add_log("hello")
    r = client.get(f"/api/job/{job.id}")
    body = r.json()
    assert body["state"] == "running" and body["progress"] == 55
    assert body["log"][-1] == "hello" and body["has_output"] is False


def test_cancel_route(client):
    job = _make_job(client, state=J.QUEUED)
    r = client.post(f"/job/{job.id}/cancel", follow_redirects=False)
    assert r.status_code == 303
    assert client.manager.store.get(job.id).state == J.CANCELLED


def test_output_download(client, tmp_path):
    docx = tmp_path / "result.docx"
    docx.write_bytes(b"PK fake docx")
    job = _make_job(client, state=J.DONE, output_path=str(docx))
    r = client.get(f"/job/{job.id}/output")
    assert r.status_code == 200
    assert r.headers["content-type"].startswith(
        "application/vnd.openxmlformats")


def test_output_missing_file_404(client):
    job = _make_job(client, state=J.DONE, output_path="E:/nope/gone.docx")
    assert client.get(f"/job/{job.id}/output").status_code == 404
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v -k "job or cancel or output"`
Expected: new tests FAIL with 404s (routes don't exist)

- [ ] **Step 3: Create the run-page template**

`webcompanion/templates/job.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>{{ task.title }} &mdash; {{ job.file_number }}</h1>
<div class="card">
 <div>Status: <span class="state" id="state">{{ job.state }}</span></div>
 <progress id="bar" max="100" value="{{ job.progress }}"></progress>
 <div id="actions">
  {% if job.state == 'awaiting_input' %}
   <a class="btn" href="/job/{{ job.id }}/awaiting">Provide input</a>
  {% elif job.state == 'done' and job.output_path %}
   <a class="btn" href="/job/{{ job.id }}/output">View result</a>
  {% endif %}
 </div>
 <pre id="log">{{ job.log[-50:] | join('\n') }}</pre>
 {% if job.state in ('queued', 'running', 'awaiting_input') %}
 <form method="post" action="/job/{{ job.id }}/cancel">
  <button class="btn secondary" type="submit">Cancel</button>
 </form>
 {% endif %}
</div>
<script>
async function tick(){
  let r;
  try { r = await fetch('/api/job/{{ job.id }}'); } catch(e) { return; }
  if(!r.ok) return;
  const s = await r.json();
  document.getElementById('state').textContent = s.state;
  document.getElementById('bar').value = s.progress;
  document.getElementById('log').textContent = s.log.join('\n');
  const a = document.getElementById('actions');
  if(s.state === 'awaiting_input'){
    a.innerHTML = '<a class="btn" href="/job/{{ job.id }}/awaiting">Provide input</a>';
  } else if(s.state === 'done' && s.has_output){
    a.innerHTML = '<a class="btn" href="/job/{{ job.id }}/output">View result</a>';
  }
  if(['done','failed','cancelled','interrupted'].includes(s.state)){
    clearInterval(timer);
  }
}
const timer = setInterval(tick, 3000);
tick();
</script>
{% endblock %}
```

- [ ] **Step 4: Implement the routes** — replace the `_register_job_routes` stub in `webcompanion/server.py`:

```python
def _register_job_routes(app, manager, templates):
    import os

    def _render(name, request, **ctx):
        return templates.TemplateResponse(name, {"request": request, **ctx})

    @app.get("/job/{job_id}", response_class=HTMLResponse)
    def job_page(request: Request, job_id: str):
        job = manager.store.get(job_id)
        if job is None:
            return HTMLResponse("Job not found", status_code=404)
        task = T.TASKS.get(job.task_id)
        return _render("job.html", request, job=job, task=task)

    @app.get("/api/job/{job_id}")
    def job_state(job_id: str):
        job = manager.store.get(job_id)
        if job is None:
            return JSONResponse({"error": "not found"}, status_code=404)
        return JSONResponse({
            "state": job.state,
            "progress": job.progress,
            "log": job.log[-50:],
            "has_output": bool(job.output_path),
            "error": job.error,
        })

    @app.post("/job/{job_id}/cancel")
    def cancel_job(job_id: str):
        manager.cancel(job_id)
        return RedirectResponse(f"/job/{job_id}", status_code=303)

    @app.get("/job/{job_id}/output")
    def job_output(job_id: str):
        job = manager.store.get(job_id)
        if job is None or not job.output_path \
                or not os.path.isfile(job.output_path):
            return HTMLResponse("Output not found", status_code=404)
        return FileResponse(
            job.output_path,
            media_type=_DOCX_MEDIA_TYPE,
            filename=os.path.basename(job.output_path),
        )
```

- [ ] **Step 5: Run all server tests**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v`
Expected: 13 passed

- [ ] **Step 6: Commit**

```bash
git add webcompanion/server.py webcompanion/templates/job.html tests/test_webcompanion/test_server.py
git commit -m "feat(webcompanion): job run page, status API, cancel and output download"
```

---

### Task 10: Depo Prep pre-run settings flow

**Files:**
- Modify: `webcompanion/server.py` (replace the `_register_depo_prep_routes` stub)
- Create: `webcompanion/templates/depo_prep_settings.html`
- Test: append to `tests/test_webcompanion/test_server.py`

- [ ] **Step 1: Write the failing tests** (append to test_server.py)

```python
def test_depo_prep_start_shows_settings(client):
    r = client.post("/case/9999/task/depo_prep/start",
                    data={"files": ["doc.pdf"]})
    assert r.status_code == 200
    assert "deponent_name" in r.text and "Lock-down" in r.text


def test_depo_prep_submit_builds_config(client, monkeypatch):
    import json
    from pathlib import Path
    submitted = {}
    monkeypatch.setattr(client.manager, "submit",
                        lambda job: submitted.setdefault("job", job) or job)
    r = client.post("/case/9999/task/depo_prep/submit", data={
        "files": "doc.pdf",
        "deponent_name": "Dr. Jones",
        "deponent_role": "Treating physician",
        "style": "expert",
        "free_text_notes": "Focus on causation.",
        "flag_strategic": "on",
    }, follow_redirects=False)
    assert r.status_code == 303
    job = submitted["job"]
    assert job.task_id == "depo_prep" and len(job.files) == 1
    cfg = json.loads(Path(job.files[0]).read_text(encoding="utf-8"))
    assert cfg["deponent_name"] == "Dr. Jones"
    assert cfg["style"] == "expert"
    assert cfg["per_topic_flags"]["strategic_note"] is True
    assert cfg["per_topic_flags"]["source_facts"] is False
    assert cfg["deponent_sources"][0].endswith("doc.pdf")
    assert cfg["context_sources"] == []
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v -k depo_prep`
Expected: FAIL (template missing / route missing)

- [ ] **Step 3: Create the settings template**

`webcompanion/templates/depo_prep_settings.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>{{ task.glyph }} Depo Prep &mdash; Settings</h1>
<form method="post" action="/case/{{ case.file_number }}/task/depo_prep/submit" class="card">
{% for rf in rel_files %}
 <input type="hidden" name="files" value="{{ rf }}">
{% endfor %}
<p><b>Sources:</b> {{ rel_files | join(', ') }}</p>
<label>Deponent name
 <input type="text" name="deponent_name" required></label>
<label>Deponent role
 <input type="text" name="deponent_role" placeholder="e.g. Treating physician"></label>
<label>Style
 <select name="style">
 {% for key, label in styles %}
  <option value="{{ key }}">{{ label }}</option>
 {% endfor %}
 </select></label>
<label>Strategy notes
 <textarea name="free_text_notes" rows="4"
  placeholder="Case theory, topics to emphasize, key admissions to extract..."></textarea></label>
<label class="row"><input type="checkbox" name="flag_strategic" checked> Strategic notes per topic</label>
<label class="row"><input type="checkbox" name="flag_source_facts" checked> Source facts per topic</label>
<label class="row"><input type="checkbox" name="flag_impeachment"> Impeachment hooks</label>
<label class="row"><input type="checkbox" name="flag_objection"> Objection alternatives</label>
<button class="btn" type="submit">Analyze sources</button>
</form>
{% endblock %}
```

- [ ] **Step 4: Implement the submit route** — replace the `_register_depo_prep_routes` stub in `webcompanion/server.py`:

```python
def _register_depo_prep_routes(app, manager, templates):
    @app.post("/case/{file_number}/task/depo_prep/submit")
    async def depo_prep_submit(request: Request, file_number: str):
        case = cases.get_case(file_number)
        if case is None:
            return HTMLResponse("Case not found", status_code=404)
        form = await request.form()
        rel_files = [f for f in form.getlist("files") if f]
        if not rel_files:
            return HTMLResponse("No source files.", status_code=400)
        try:
            abs_files = [str(cases.safe_resolve(case["case_path"], rf))
                         for rf in rel_files]
        except ValueError:
            return HTMLResponse("Invalid path", status_code=400)
        cfg = {
            "deponent_name": (form.get("deponent_name") or "").strip(),
            "deponent_role": (form.get("deponent_role") or "").strip(),
            "deponent_sources": abs_files,
            "context_sources": [],
            "style": form.get("style") or "discovery",
            "free_text_notes": (form.get("free_text_notes") or "").strip(),
            "per_topic_flags": {
                "strategic_note": form.get("flag_strategic") == "on",
                "source_facts": form.get("flag_source_facts") == "on",
                "impeachment_hook": form.get("flag_impeachment") == "on",
                "objection_alts": form.get("flag_objection") == "on",
            },
            "case_root": case["case_path"],
        }
        cfg_path = T.write_depo_prep_config(cfg)
        job = new_job("depo_prep", case["case_path"], file_number, [cfg_path])
        manager.submit(job)
        return RedirectResponse(f"/job/{job.id}", status_code=303)
```

- [ ] **Step 5: Run all server tests**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v`
Expected: 15 passed

- [ ] **Step 6: Commit**

```bash
git add webcompanion/server.py webcompanion/templates/depo_prep_settings.html tests/test_webcompanion/test_server.py
git commit -m "feat(webcompanion): depo prep pre-run settings form"
```

---

### Task 11: Awaiting-input forms + resume (deposition, med chron, depo prep)

**Files:**
- Modify: `webcompanion/server.py` (replace the `_register_awaiting_routes` stub)
- Create: `webcompanion/templates/awaiting_deposition.html`, `awaiting_med_chron.html`, `awaiting_depo_prep.html`
- Test: append to `tests/test_webcompanion/test_server.py`

- [ ] **Step 1: Write the failing tests** (append to test_server.py)

```python
import json as _json


def _awaiting_job(client, tmp_path, task_id, session_data, topics=None):
    from webcompanion.jobs import new_job
    session = tmp_path / "session.json"
    session.write_text(_json.dumps(session_data), encoding="utf-8")
    if topics is not None:
        (tmp_path / "topics.json").write_text(
            _json.dumps({"topics": topics}), encoding="utf-8")
    job = new_job(task_id, "E:/case", "9999", ["E:/case/d.pdf"])
    job.state = J.AWAITING_INPUT
    job.session_path = str(session)
    client.manager.store.add(job)
    return job


def test_awaiting_deposition_form(client, tmp_path):
    job = _awaiting_job(client, tmp_path, "summarize_depositions", {
        "deponent_name": "Dr. Jones", "deponent_type": "expert",
        "deposition_date": "2026-01-15",
        "topics": [{"title": "Background"}, {"title": "Treatment"}],
    })
    r = client.get(f"/job/{job.id}/awaiting")
    assert r.status_code == 200
    assert "Dr. Jones" in r.text and "Background" in r.text


def test_awaiting_deposition_resume(client, tmp_path, monkeypatch):
    job = _awaiting_job(client, tmp_path, "summarize_depositions", {
        "topics": [{"title": "Background"}],
    })
    applied, resumed = {}, []
    monkeypatch.setattr(
        "webcompanion.server.T.apply_deposition_user_config",
        lambda sp, cfg: applied.update(cfg))
    monkeypatch.setattr(client.manager, "resume",
                        lambda jid: resumed.append(jid))
    r = client.post(f"/job/{job.id}/resume", data={
        "topic": ["Background"], "added_topics": "Damages\nPrognosis",
        "bullets": "7", "deponent_label": "Dr. Jones",
        "audience": "pro_defense", "tone": "recitation",
        "cross_check": "on",
    }, follow_redirects=False)
    assert r.status_code == 303 and resumed == [job.id]
    assert applied["selected_topics"] == ["Background"]
    assert applied["added_topics"] == ["Damages", "Prognosis"]
    assert applied["bullets_per_topic"] == 7
    assert applied["cross_check_enabled"] is True
    assert applied["audience"] == "pro_defense"


def test_awaiting_deposition_requires_topic(client, tmp_path):
    job = _awaiting_job(client, tmp_path, "summarize_depositions",
                        {"topics": []})
    r = client.post(f"/job/{job.id}/resume", data={"added_topics": ""})
    assert r.status_code == 400


def test_awaiting_med_chron_form_and_resume(client, tmp_path, monkeypatch):
    job = _awaiting_job(client, tmp_path, "med_chron_analysis", {
        "provider_name": "Kaiser",
        "catalog": [{"id": "gaps", "label": "Treatment gaps"},
                    {"id": "billing", "label": "Billing analysis"}],
    })
    r = client.get(f"/job/{job.id}/awaiting")
    assert "Kaiser" in r.text and "Treatment gaps" in r.text

    applied, resumed = {}, []
    monkeypatch.setattr(
        "webcompanion.server.T.apply_med_chron_user_config",
        lambda sp, cfg: applied.update(cfg))
    monkeypatch.setattr(client.manager, "resume",
                        lambda jid: resumed.append(jid))
    r = client.post(f"/job/{job.id}/resume", data={
        "analysis": ["gaps"],
        "custom_label_1": "IME prep", "custom_instruction_1": "Flag inconsistencies",
    }, follow_redirects=False)
    assert r.status_code == 303 and resumed == [job.id]
    assert applied["selected_catalog_ids"] == ["gaps"]
    assert applied["custom_analyses"] == [
        {"label": "IME prep", "instruction": "Flag inconsistencies",
         "context_files": []}]


def test_awaiting_depo_prep_form_and_resume(client, tmp_path, monkeypatch):
    topics = [{"id": "t01", "title": "Background", "strategic_note": "note",
               "relevant_digest_refs": ["d1"], "default_checked": True,
               "lawyer_added": False}]
    job = _awaiting_job(client, tmp_path, "depo_prep", {}, topics=topics)
    r = client.get(f"/job/{job.id}/awaiting")
    assert "Background" in r.text

    resumed = []
    monkeypatch.setattr(client.manager, "resume",
                        lambda jid: resumed.append(jid))
    r = client.post(f"/job/{job.id}/resume", data={
        "keep_0": "on", "title_0": "Background (edited)", "note_0": "note",
        "new_title_1": "Damages", "new_note_1": "",
    }, follow_redirects=False)
    assert r.status_code == 303 and resumed == [job.id]
    written = _json.loads(
        (tmp_path / "topics.json").read_text(encoding="utf-8"))["topics"]
    assert written[0]["title"] == "Background (edited)"
    assert written[0]["relevant_digest_refs"] == ["d1"]
    assert written[1]["title"] == "Damages" and written[1]["lawyer_added"] is True


def test_awaiting_on_non_awaiting_job_404(client):
    job = _make_job(client, state=J.RUNNING)
    assert client.get(f"/job/{job.id}/awaiting").status_code == 404
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_webcompanion/test_server.py -v -k awaiting`
Expected: FAIL with 404s

- [ ] **Step 3: Create the three templates**

`webcompanion/templates/awaiting_deposition.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>Pick topics &mdash; {{ session.get('deponent_name', 'Deponent') }}</h1>
<p>{{ session.get('deponent_type', '') }} {{ session.get('deposition_date', '') }}</p>
<form method="post" action="/job/{{ job.id }}/resume" class="card">
<h2>Topics</h2>
{% for t in session.get('topics', []) %}
<label class="row"><input type="checkbox" name="topic" value="{{ t.get('title', '') }}" checked> {{ t.get('title', '') }}</label>
{% endfor %}
<label>Add topics (one per line)
 <textarea name="added_topics" rows="3"></textarea></label>
<label>Bullets per topic
 <input type="number" name="bullets" value="5" min="1" max="15"></label>
<label>Deponent label
 <input type="text" name="deponent_label" value="{{ session.get('deponent_name', 'Deponent') }}"></label>
<label>Audience
 <select name="audience">
 {% for key, label in audiences %}<option value="{{ key }}">{{ label }}</option>{% endfor %}
 </select></label>
<label>Custom audience (if Custom)
 <input type="text" name="audience_custom"></label>
<label>Tone
 <select name="tone">
 {% for key, label in tones %}<option value="{{ key }}">{{ label }}</option>{% endfor %}
 </select></label>
<label>Custom tone (if Custom)
 <input type="text" name="tone_custom"></label>
<label>Custom rules
 <textarea name="custom_rules" rows="2"></textarea></label>
<label class="row"><input type="checkbox" name="cross_check"> Cross-check citations</label>
<button class="btn" type="submit">Generate summary</button>
</form>
{% endblock %}
```

`webcompanion/templates/awaiting_med_chron.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>Analyses &mdash; {{ session.get('provider_name', '') }}</h1>
<form method="post" action="/job/{{ job.id }}/resume" class="card">
<h2>Catalog</h2>
{% for entry in session.get('catalog', []) %}
<label class="row"><input type="checkbox" name="analysis" value="{{ entry.id }}"> {{ entry.label }}</label>
{% endfor %}
<h2>Custom analyses</h2>
{% for i in (1, 2, 3) %}
<label>Label <input type="text" name="custom_label_{{ i }}"></label>
<label>Instruction <input type="text" name="custom_instruction_{{ i }}"></label>
{% endfor %}
<button class="btn" type="submit">Run analyses</button>
</form>
{% endblock %}
```

`webcompanion/templates/awaiting_depo_prep.html`:

```html
{% extends "base.html" %}
{% block content %}
<h1>Review topics</h1>
<form method="post" action="/job/{{ job.id }}/resume" class="card">
{% for t in topics %}
<div class="card">
 <label class="row"><input type="checkbox" name="keep_{{ loop.index0 }}"
  {% if t.get('default_checked', True) %}checked{% endif %}> Include</label>
 <input type="text" name="title_{{ loop.index0 }}" value="{{ t.get('title', '') }}">
 <input type="text" name="note_{{ loop.index0 }}" value="{{ t.get('strategic_note', '') }}" placeholder="Strategic note">
</div>
{% endfor %}
<h2>Add topics</h2>
{% for i in (1, 2, 3) %}
<input type="text" name="new_title_{{ i }}" placeholder="New topic title">
<input type="text" name="new_note_{{ i }}" placeholder="Strategic note (optional)">
{% endfor %}
<button class="btn" type="submit">Generate outline</button>
</form>
{% endblock %}
```

- [ ] **Step 4: Implement the routes** — replace the `_register_awaiting_routes` stub in `webcompanion/server.py`:

```python
def _register_awaiting_routes(app, manager, templates):
    def _render(name, request, **ctx):
        return templates.TemplateResponse(name, {"request": request, **ctx})

    def _get_awaiting(job_id):
        job = manager.store.get(job_id)
        if job is None or job.state != J.AWAITING_INPUT \
                or not job.session_path:
            return None
        return job

    @app.get("/job/{job_id}/awaiting", response_class=HTMLResponse)
    def awaiting_form(request: Request, job_id: str):
        job = _get_awaiting(job_id)
        if job is None:
            return HTMLResponse("Not awaiting input", status_code=404)
        kind = T.TASKS[job.task_id].awaiting_kind
        if kind == "deposition":
            session = T.read_session_json(job.session_path)
            return _render("awaiting_deposition.html", request, job=job,
                           session=session, audiences=T.DEPO_AUDIENCES,
                           tones=T.DEPO_TONES)
        if kind == "med_chron":
            session = T.read_session_json(job.session_path)
            return _render("awaiting_med_chron.html", request, job=job,
                           session=session)
        if kind == "depo_prep":
            topics = T.read_depo_prep_topics(job.session_path)
            return _render("awaiting_depo_prep.html", request, job=job,
                           topics=topics)
        return HTMLResponse("Unknown task kind", status_code=400)

    @app.post("/job/{job_id}/resume")
    async def resume_job(request: Request, job_id: str):
        job = _get_awaiting(job_id)
        if job is None:
            return HTMLResponse("Not awaiting input", status_code=404)
        form = await request.form()
        kind = T.TASKS[job.task_id].awaiting_kind

        if kind == "deposition":
            cfg = {
                "selected_topics": [t for t in form.getlist("topic")
                                    if t.strip()],
                "added_topics": [
                    ln.strip()
                    for ln in (form.get("added_topics") or "").splitlines()
                    if ln.strip()],
                "bullets_per_topic": int(form.get("bullets") or 5),
                "deponent_label": (form.get("deponent_label") or "").strip()
                                  or "Deponent",
                "custom_rules": (form.get("custom_rules") or "").strip(),
                "cross_check_enabled": form.get("cross_check") == "on",
                "context_doc_paths": [],
                "audience": form.get("audience") or "neutral",
                "audience_custom": (
                    (form.get("audience_custom") or "").strip()
                    if form.get("audience") == "custom" else ""),
                "tone": form.get("tone") or "recitation",
                "tone_custom": (
                    (form.get("tone_custom") or "").strip()
                    if form.get("tone") == "custom" else ""),
            }
            if not cfg["selected_topics"] and not cfg["added_topics"]:
                return HTMLResponse("Select or add at least one topic.",
                                    status_code=400)
            T.apply_deposition_user_config(job.session_path, cfg)

        elif kind == "med_chron":
            selected = [a for a in form.getlist("analysis") if a]
            custom = []
            for i in (1, 2, 3):
                lbl = (form.get(f"custom_label_{i}") or "").strip()
                instr = (form.get(f"custom_instruction_{i}") or "").strip()
                if lbl and instr:
                    custom.append({"label": lbl, "instruction": instr,
                                   "context_files": []})
            if not selected and not custom:
                return HTMLResponse(
                    "Select or add at least one analysis.", status_code=400)
            T.apply_med_chron_user_config(job.session_path, {
                "selected_catalog_ids": selected,
                "custom_analyses": custom,
            })

        elif kind == "depo_prep":
            existing = T.read_depo_prep_topics(job.session_path)
            topics = []
            for i, t in enumerate(existing):
                topics.append({
                    "id": t.get("id") or f"t{i + 1:02d}",
                    "title": (form.get(f"title_{i}")
                              or t.get("title", "")).strip(),
                    "strategic_note": (form.get(f"note_{i}") or "").strip(),
                    "relevant_digest_refs": t.get("relevant_digest_refs", []),
                    "default_checked": form.get(f"keep_{i}") == "on",
                    "lawyer_added": bool(t.get("lawyer_added", False)),
                })
            for i in (1, 2, 3):
                title = (form.get(f"new_title_{i}") or "").strip()
                if title:
                    topics.append({
                        "id": f"t{len(topics) + 1:02d}", "title": title,
                        "strategic_note": (form.get(f"new_note_{i}")
                                           or "").strip(),
                        "relevant_digest_refs": [],
                        "default_checked": True, "lawyer_added": True,
                    })
            if not any(t["default_checked"] for t in topics):
                return HTMLResponse("Keep or add at least one topic.",
                                    status_code=400)
            T.write_depo_prep_topics(job.session_path, topics)

        else:
            return HTMLResponse("Unknown task kind", status_code=400)

        manager.resume(job_id)
        return RedirectResponse(f"/job/{job_id}", status_code=303)
```

- [ ] **Step 5: Run the full test suite**

Run: `python -m pytest tests/test_webcompanion/ -v`
Expected: all pass (~45 tests)

- [ ] **Step 6: Commit**

```bash
git add webcompanion/server.py webcompanion/templates/awaiting_deposition.html webcompanion/templates/awaiting_med_chron.html webcompanion/templates/awaiting_depo_prep.html tests/test_webcompanion/test_server.py
git commit -m "feat(webcompanion): awaiting-input forms and two-phase resume"
```

---

### Task 12: Launcher + docs

**Files:**
- Create: `run_webcompanion.bat`
- Modify: `CLAUDE.md` (add a short section under "Recent Features")

- [ ] **Step 1: Create the launcher**

`run_webcompanion.bat`:

```bat
@echo off
cd /d "%~dp0"
python -m webcompanion.server %*
pause
```

- [ ] **Step 2: Smoke-test the server boots in LAN mode**

Run: `Start-Process python -ArgumentList "-m","webcompanion.server","--lan","--port","8766" -PassThru` then after ~5 s: `Invoke-WebRequest http://127.0.0.1:8766/ -UseBasicParsing | Select-Object -ExpandProperty StatusCode`; stop the process afterwards.
Expected: `200`

- [ ] **Step 3: Document in CLAUDE.md** — append under "Recent Features":

```markdown
### Wizard Web Companion (2026-06-12)
- Standalone FastAPI server (`python -m webcompanion.server`, port 8765) for
  running the seven script-based wizard tasks from a phone over Tailscale
- Reuses the wizard stdout protocol and session managers; desktop app untouched
- `--lan` flag binds 0.0.0.0 for local development; default binds Tailscale IP
- Spec: `docs/superpowers/specs/2026-06-12-wizard-web-companion-design.md`
```

- [ ] **Step 4: Commit**

```bash
git add run_webcompanion.bat CLAUDE.md
git commit -m "feat(webcompanion): launcher script and docs"
```

---

### Task 13: Manual end-to-end verification (MANDATORY before declaring done)

No code. Verify against the real machine per the global "always test after developing" rule.

- [ ] **Step 1:** Start the server: `run_webcompanion.bat` (Tailscale must be up; note the printed URL).
- [ ] **Step 2:** From iPhone Safari over Tailscale, open `http://<tailscale-ip>:8765`. Verify the case list loads and search works.
- [ ] **Step 3:** Single-phase run: pick a small test case → Summarize Documents → pick one small PDF → Run. Verify progress updates, completion, and that "View result" previews the .docx in Safari.
- [ ] **Step 4:** Two-phase run: same case → Summarize Depositions on a transcript → wait for "Provide input" → pick topics on the phone → resume → verify the final .docx.
- [ ] **Step 5:** Restart resilience: kill the server mid-run, restart, verify the job shows `interrupted` and the UI stays usable.
- [ ] **Step 6:** Verify the desktop iCharlotte app still launches and its wizard still runs a task (no regression — the desktop code was never touched, this is a sanity check).
- [ ] **Step 7:** Fix anything found, re-run `python -m pytest tests/test_webcompanion/ -v`, commit fixes as `fix(webcompanion): …`.
