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
