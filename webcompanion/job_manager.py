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
        running_count = self._running_count()
        running_paths = self._running_case_paths()
        for job in reversed(self.store.all()):
            if job.state != J.QUEUED:
                continue
            if running_count >= self._max:
                return
            if job.case_path in running_paths:
                continue
            self._start_current_file(job)
            running_count += 1
            running_paths.add(job.case_path)

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
            elif (ev.message or "").strip():
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
                self._file_idx.pop(job_id, None)
                job.state = J.CANCELLED
                job.add_log("Cancelled.")
            elif returncode != 0:
                self._file_idx.pop(job_id, None)
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
                self._file_idx.pop(job_id, None)
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
