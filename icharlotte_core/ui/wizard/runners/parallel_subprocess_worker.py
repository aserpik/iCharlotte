"""ParallelSubprocessWorker — runs an existing Scripts/*.py agent across many
input files concurrently, with a bounded pool of QProcess instances.

Per-input output isolation
--------------------------
Each input file gets its own unique output .docx, passed via the agent's
`--output_path <path>` flag. Filenames are derived from the input file's
basename:  ``AI_OUTPUT - <input_stem>.docx`` under ``NOTES/AI Output/``.
This avoids races on agents that default to a single shared filename
(summarize.py → AI_OUTPUT.docx, summarize_deposition.py →
Deposition_summaries.docx).

Concurrency
-----------
A bounded pool starts up to ``max_concurrent`` processes at once. As each
finishes, the next queued input is launched. Default cap is 3 to respect
LLM RPM/TPM limits and keep memory bounded (~200-500 MB per agent).

Two-phase agents
----------------
``summarize_deposition.py``'s AWAITING_INPUT (Phase 1 → user picks topics
→ Phase 2) flow does **not** map cleanly onto this worker. Callers must
fall back to the sequential ``SubprocessWorker`` for that script. We
guard against it at instantiation time.

Cancellation
------------
``cancel()`` terminates all running processes and drains the queue, then
emits ``cancelled``.

Output signal
-------------
On success, ``finished`` emits the newest (most recently modified) output
path. All produced output paths are available via ``output_paths`` after
completion (for the OutputPage to list/merge if needed).
"""
import os
import re
import sys
from typing import Dict, List, Optional, Set

from PySide6.QtCore import QProcess, QTimer, Signal

from .base import BaseWorker


_AI_OUTPUT_SUBPATH = os.path.join("NOTES", "AI Output")
_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*(?::(.*))?$")
_AWAITING_RE = re.compile(r"^AWAITING_INPUT:(.+)$")

# Filesystem-safe replacement for chars that are illegal in Windows paths.
_UNSAFE_FILENAME = re.compile(r'[\\/*?:"<>|]')

# Scripts whose two-phase AWAITING_INPUT flow is incompatible with this worker.
_TWO_PHASE_SCRIPTS = {"summarize_deposition.py", "med_chron.py"}


class ParallelSubprocessWorker(BaseWorker):
    """Run a Scripts/*.py agent across multiple files concurrently.

    Compared to the sequential :class:`SubprocessWorker`:
    - launches up to ``max_concurrent`` QProcesses at once
    - assigns each input file a unique output path via ``--output_path``
    - aggregates progress as the mean across in-flight + completed runs
    - emits ``file_completed(output_path)`` each time a child run produces
      a docx, so the UI can stream outputs into the Output page as they
      arrive instead of waiting for the whole batch
    - reports failure if *any* child process exits non-zero (but still
      waits for the rest to finish before emitting ``failed``)
    """

    # Emitted each time a child process completes successfully and its
    # output .docx is on disk. Carries the absolute output path.
    file_completed = Signal(str)

    def __init__(
        self,
        script_name: str,
        case_path: str,
        file_number: str,
        files: List[str],
        settings: dict,
        max_concurrent: int = 3,
        parent=None,
    ):
        if script_name in _TWO_PHASE_SCRIPTS:
            raise ValueError(
                f"{script_name} is a two-phase agent; use SubprocessWorker (sequential) "
                "for its AWAITING_INPUT/Phase 2 flow."
            )
        if not files:
            raise ValueError("ParallelSubprocessWorker requires at least one input file.")
        if max_concurrent < 1:
            raise ValueError("max_concurrent must be >= 1.")

        super().__init__(case_path, file_number, files, settings, parent)
        self._script_name = script_name
        self._max_concurrent = max_concurrent

        # File-level state
        self._queued: List[str] = list(files)
        self._in_flight: Dict[QProcess, str] = {}          # process -> input file
        self._stdout_buf: Dict[QProcess, bytearray] = {}
        self._proc_output_path: Dict[QProcess, str] = {}   # process -> target output path
        self._raw_pct: Dict[str, int] = {f: 0 for f in files}   # 0-100 per file
        self._completed_files: Set[str] = set()
        self._failed_files: Set[str] = set()
        self._output_paths: List[str] = []                 # populated as each child completes
        self._first_failure_msg: Optional[str] = None

    # ---- Public API ----

    @property
    def output_paths(self) -> List[str]:
        return list(self._output_paths)

    # ---- Lifecycle ----

    def start(self) -> None:
        # Ensure output directory exists (it normally does).
        out_dir = os.path.join(self.case_path, _AI_OUTPUT_SUBPATH)
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as e:
            self.failed.emit(f"Could not create output directory {out_dir}: {e}")
            return

        self.status.emit(
            f"Starting {self._script_name} on {len(self._queued)} file(s) "
            f"(up to {self._max_concurrent} concurrent)…"
        )
        # Kick off the initial wave.
        for _ in range(min(self._max_concurrent, len(self._queued))):
            self._spawn_next()

    def cancel(self) -> None:
        super().cancel()
        # Drop remaining queue so no new work starts.
        self._queued.clear()
        if not self._in_flight:
            self.cancelled.emit()
            return
        self.status.emit(f"Cancellation requested — terminating {len(self._in_flight)} running agent(s)…")
        for proc in list(self._in_flight.keys()):
            try:
                proc.terminate()
            except RuntimeError:
                pass
        # Force-kill any survivors after 2s.
        QTimer.singleShot(2000, self._hard_kill_survivors)

    def _hard_kill_survivors(self) -> None:
        for proc in list(self._in_flight.keys()):
            try:
                if proc.state() != QProcess.ProcessState.NotRunning:
                    proc.kill()
            except RuntimeError:
                pass

    # ---- Spawning ----

    def _spawn_next(self) -> None:
        if not self._queued or self.is_cancel_requested:
            return
        file_path = self._queued.pop(0)
        output_path = self._unique_output_path_for(file_path)

        proc = QProcess(self)
        proc.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        proc.readyReadStandardOutput.connect(lambda p=proc: self._drain_stdout(p))
        proc.finished.connect(
            lambda code, status, p=proc: self._on_process_finished(p, code, status)
        )
        proc.errorOccurred.connect(lambda err, p=proc: self._on_process_error(p, err))

        argv = [
            "-u",
            self._script_path(),
            file_path,
            "--output_path",
            output_path,
        ]
        self.status.emit(
            f"[{os.path.basename(file_path)}] running → {os.path.basename(output_path)}"
        )

        self._in_flight[proc] = file_path
        self._stdout_buf[proc] = bytearray()
        self._proc_output_path[proc] = output_path

        proc.start(sys.executable, argv)

    def _unique_output_path_for(self, input_file: str) -> str:
        """Build a per-input-file output path under NOTES/AI Output/.

        Convention: ``AI_OUTPUT - <input_stem>.docx``. If that path is
        already in use by an earlier scheduled file in this batch, append
        ``(2)``, ``(3)``, ... to disambiguate.
        """
        stem = os.path.splitext(os.path.basename(input_file))[0]
        safe = _UNSAFE_FILENAME.sub("", stem).strip()
        if not safe:
            safe = "input"
        base = f"AI_OUTPUT - {safe}"
        out_dir = os.path.join(self.case_path, _AI_OUTPUT_SUBPATH)
        candidate = os.path.join(out_dir, f"{base}.docx")
        # Disambiguate against other targets already chosen in this batch.
        n = 2
        chosen = set(self._output_paths) | set(self._proc_output_path.values())
        while candidate in chosen:
            candidate = os.path.join(out_dir, f"{base} ({n}).docx")
            n += 1
        return candidate

    def _script_path(self) -> str:
        # __file__ lives 5 levels deep:
        #   <root>/icharlotte_core/ui/wizard/runners/parallel_subprocess_worker.py
        here = os.path.abspath(__file__)
        repo_root = here
        for _ in range(5):
            repo_root = os.path.dirname(repo_root)
        return os.path.join(repo_root, "Scripts", self._script_name)

    # ---- Stdout / events ----

    def _drain_stdout(self, proc: QProcess) -> None:
        if proc not in self._stdout_buf:
            return
        data = bytes(proc.readAllStandardOutput())
        if not data:
            return
        buf = self._stdout_buf[proc]
        buf.extend(data)
        while b"\n" in buf:
            line_bytes, _, rest = buf.partition(b"\n")
            self._stdout_buf[proc] = bytearray(rest)
            buf = self._stdout_buf[proc]
            line = line_bytes.decode("utf-8", errors="replace").rstrip("\r")
            self._handle_line(proc, line)

    def _handle_line(self, proc: QProcess, line: str) -> None:
        input_file = self._in_flight.get(proc)
        if input_file is None:
            return

        m = _PROGRESS_RE.match(line)
        if m:
            try:
                raw = int(m.group(1))
                self._raw_pct[input_file] = raw
                self._emit_aggregate_progress()
                msg = (m.group(2) or "").strip()
                if msg:
                    self.status.emit(f"[{os.path.basename(input_file)}] {msg}")
            except ValueError:
                pass
            return

        # AWAITING_INPUT inside parallel is a configuration error: we already
        # refused to start two-phase scripts. Surface it loudly.
        if _AWAITING_RE.match(line):
            self.status.emit(
                f"[{os.path.basename(input_file)}] unexpected AWAITING_INPUT in parallel mode; ignoring"
            )
            return

        if line:
            self.status.emit(f"[{os.path.basename(input_file)}] {line}")

    def _emit_aggregate_progress(self) -> None:
        total = len(self.files)
        if total == 0:
            return
        # Completed files contribute 100; failed contribute 100 (run is done either way);
        # in-flight contribute their last reported raw_pct.
        pct_sum = 0
        for f in self.files:
            if f in self._completed_files or f in self._failed_files:
                pct_sum += 100
            else:
                pct_sum += max(0, min(100, self._raw_pct.get(f, 0)))
        self.progress.emit(max(0, min(100, pct_sum // total)))

    def _on_process_finished(
        self, proc: QProcess, exit_code: int, exit_status: QProcess.ExitStatus
    ) -> None:
        # Drain any tail bytes.
        self._drain_stdout(proc)
        tail = bytes(self._stdout_buf.get(proc, b""))
        if tail:
            self._handle_line(proc, tail.decode("utf-8", errors="replace"))
            self._stdout_buf[proc] = bytearray()

        input_file = self._in_flight.pop(proc, None)
        self._stdout_buf.pop(proc, None)
        output_path = self._proc_output_path.pop(proc, None)

        if input_file is None:
            # Already accounted for (e.g. duplicate finished signal).
            return

        if self.is_cancel_requested:
            # Don't process — wait for all to wind down then emit cancelled once.
            self._maybe_emit_cancelled()
            return

        if exit_code != 0:
            self._failed_files.add(input_file)
            msg = f"[{os.path.basename(input_file)}] agent exited with code {exit_code}"
            if self._first_failure_msg is None:
                self._first_failure_msg = msg
            self.status.emit(msg)
        else:
            # Success: claim the targeted output_path if it exists.
            if output_path and os.path.isfile(output_path):
                self._output_paths.append(output_path)
                # Notify subscribers (TaskTab) immediately so the user can
                # view this file's output without waiting for the rest.
                self.file_completed.emit(output_path)
            else:
                # Agent reported success but the expected file isn't there —
                # treat as failure so the user notices.
                self._failed_files.add(input_file)
                msg = (
                    f"[{os.path.basename(input_file)}] reported success but "
                    f"{os.path.basename(output_path) if output_path else 'output file'} is missing"
                )
                if self._first_failure_msg is None:
                    self._first_failure_msg = msg
                self.status.emit(msg)
            self._completed_files.add(input_file)

        self._emit_aggregate_progress()

        # Refill the pool.
        if self._queued and not self.is_cancel_requested:
            self._spawn_next()

        # All done?
        if not self._in_flight and not self._queued:
            self._finalize()

    def _on_process_error(self, proc: QProcess, err: QProcess.ProcessError) -> None:
        if self.is_cancel_requested:
            return
        input_file = self._in_flight.get(proc)
        msg = f"Process error ({err})"
        if input_file:
            msg = f"[{os.path.basename(input_file)}] {msg}"
            self._failed_files.add(input_file)
        if self._first_failure_msg is None:
            self._first_failure_msg = msg
        self.status.emit(msg)

    def _finalize(self) -> None:
        if self.is_cancel_requested:
            self.cancelled.emit()
            return
        if self._failed_files:
            failed_names = ", ".join(sorted(os.path.basename(f) for f in self._failed_files))
            self.failed.emit(
                f"{len(self._failed_files)} of {len(self.files)} agent runs failed: "
                f"{failed_names}. First error: {self._first_failure_msg}"
            )
            return
        if not self._output_paths:
            self.failed.emit("All agent runs completed but no output files were produced.")
            return
        # Surface the newest output path for the OutputPage's initial view.
        newest = max(self._output_paths, key=lambda p: os.path.getmtime(p))
        self.finished.emit(newest)

    def _maybe_emit_cancelled(self) -> None:
        if not self._in_flight and not self._queued:
            self.cancelled.emit()
