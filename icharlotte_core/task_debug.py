"""Task execution debug recorder.

This module is intentionally best-effort: recorder failures must not interrupt
the task being observed.
"""

from __future__ import annotations

from collections import deque
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
import logging
import math
from pathlib import Path
import re
import threading
import time
from typing import Any
from uuid import uuid4

try:
    from PySide6.QtCore import QObject, Signal
except Exception:  # pragma: no cover - only used when Qt is unavailable.
    QObject = object  # type: ignore[assignment]

    class _FallbackSignal:
        def __init__(self, *args: object) -> None:
            self._subscribers = []

        def connect(self, callback: object) -> None:
            self._subscribers.append(callback)

        def emit(self, value: object) -> None:
            for callback in list(self._subscribers):
                try:
                    callback(value)
                except Exception:
                    pass

    Signal = _FallbackSignal  # type: ignore[assignment]


logger = logging.getLogger(__name__)

_DEFAULT_MAX_EVENTS = 1000
_SENSITIVE_KEY_RE = re.compile(r"(token|api_key|password|secret|credential)", re.I)


@dataclass(frozen=True)
class TaskDebugEvent:
    timestamp: str
    run_id: str
    task_id: str
    task_title: str
    phase: str
    level: str
    message: str
    elapsed_ms: int
    source: str
    details: dict[str, Any]
    trace_path: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class TaskDebugBridge(QObject):
    event_emitted = Signal(object)


class _TaskDebugRecorder:
    def __init__(
        self,
        *,
        trace_dir: str | Path | None = None,
        max_events: int | None = None,
    ) -> None:
        self._lock = threading.RLock()
        self._trace_dir = Path(trace_dir) if trace_dir is not None else _default_trace_dir()
        self._max_events = max_events if max_events is not None else _DEFAULT_MAX_EVENTS
        self._events: deque[TaskDebugEvent] = deque(maxlen=self._max_events)
        self._active_runs: dict[str, dict[str, Any]] = {}
        self._bridge = TaskDebugBridge()

    def reset_for_tests(
        self,
        *,
        trace_dir: str | Path | None = None,
        max_events: int | None = None,
    ) -> None:
        with self._lock:
            self._trace_dir = Path(trace_dir) if trace_dir is not None else _default_trace_dir()
            self._max_events = max_events if max_events is not None else _DEFAULT_MAX_EVENTS
            self._events = deque(maxlen=self._max_events)
            self._active_runs = {}
            self._bridge = TaskDebugBridge()

    def start_run(
        self,
        task_id: str,
        task_title: str,
        source: str = "",
        details: dict[str, Any] | None = None,
    ) -> str:
        run_id = uuid4().hex
        started_at = time.perf_counter()
        trace_path = self._trace_dir / _trace_filename(run_id, task_id)
        with self._lock:
            self._active_runs[run_id] = {
                "task_id": task_id,
                "task_title": task_title,
                "source": source,
                "started_at": started_at,
                "trace_path": str(trace_path),
            }
        self._record(
            run_id=run_id,
            task_id=task_id,
            task_title=task_title,
            phase="start",
            level="info",
            message="Task started",
            source=source,
            details=details,
            elapsed_ms=0,
            trace_path=str(trace_path),
        )
        return run_id

    def emit_event(
        self,
        run_id: str,
        task_id: str,
        task_title: str,
        phase: str,
        message: str,
        level: str = "info",
        source: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        with self._lock:
            meta = dict(self._active_runs.get(run_id, {}))
        started_at = meta.get("started_at")
        elapsed_ms = _elapsed_ms(started_at)
        trace_path = meta.get("trace_path", "")
        self._record(
            run_id=run_id,
            task_id=task_id,
            task_title=task_title,
            phase=phase,
            level=level,
            message=message,
            source=source,
            details=details,
            elapsed_ms=elapsed_ms,
            trace_path=trace_path,
        )

    def finish_run(
        self,
        run_id: str,
        status: str = "success",
        message: str = "Task complete",
        details: dict[str, Any] | None = None,
    ) -> None:
        with self._lock:
            meta = dict(self._active_runs.get(run_id, {}))
        started_at = meta.get("started_at")
        finish_details = dict(details or {})
        finish_details["status"] = status
        level = "info" if status == "success" else "error"
        self._record(
            run_id=run_id,
            task_id=meta.get("task_id", ""),
            task_title=meta.get("task_title", ""),
            phase="finish",
            level=level,
            message=message,
            source=meta.get("source", ""),
            details=finish_details,
            elapsed_ms=_elapsed_ms(started_at),
            trace_path=meta.get("trace_path", ""),
        )
        with self._lock:
            self._active_runs.pop(run_id, None)

    def get_events(self) -> list[TaskDebugEvent]:
        with self._lock:
            return list(self._events)

    def clear_events(self) -> None:
        with self._lock:
            self._events.clear()

    def get_trace_dir(self) -> Path:
        with self._lock:
            return self._trace_dir

    def get_bridge(self) -> TaskDebugBridge:
        with self._lock:
            return self._bridge

    def _record(
        self,
        *,
        run_id: str,
        task_id: str,
        task_title: str,
        phase: str,
        level: str,
        message: str,
        source: str,
        details: dict[str, Any] | None,
        elapsed_ms: int,
        trace_path: str,
    ) -> None:
        try:
            event = TaskDebugEvent(
                timestamp=datetime.now(timezone.utc).isoformat(),
                run_id=str(run_id),
                task_id=str(task_id),
                task_title=str(task_title),
                phase=str(phase),
                level=str(level),
                message=str(message),
                elapsed_ms=max(0, int(elapsed_ms)),
                source=str(source),
                details=_safe_details(details),
                trace_path=str(trace_path),
            )
        except Exception:
            _safe_log("Failed to build task debug event")
            return

        bridge: TaskDebugBridge | None = None
        with self._lock:
            try:
                self._events.append(event)
            except Exception:
                _safe_log("Failed to buffer task debug event")
            self._write_jsonl(event)
            try:
                bridge = self._bridge
            except Exception:
                bridge = None

        if bridge is not None:
            try:
                bridge.event_emitted.emit(event)
            except Exception:
                _safe_log("Failed to emit task debug event")

    def _write_jsonl(self, event: TaskDebugEvent) -> None:
        if not event.trace_path:
            return
        try:
            path = Path(event.trace_path)
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(
                    json.dumps(event.to_dict(), ensure_ascii=False, allow_nan=False)
                    + "\n"
                )
        except Exception:
            _safe_log("Failed to write task debug trace")


def start_run(
    task_id: str,
    task_title: str,
    source: str = "",
    details: dict[str, Any] | None = None,
) -> str:
    try:
        return _RECORDER.start_run(task_id, task_title, source=source, details=details)
    except Exception:
        _safe_log("Task debug start_run failed")
        return uuid4().hex


def emit_event(
    run_id: str,
    task_id: str,
    task_title: str,
    phase: str,
    message: str,
    level: str = "info",
    source: str = "",
    details: dict[str, Any] | None = None,
) -> None:
    try:
        _RECORDER.emit_event(
            run_id=run_id,
            task_id=task_id,
            task_title=task_title,
            phase=phase,
            message=message,
            level=level,
            source=source,
            details=details,
        )
    except Exception:
        _safe_log("Task debug emit_event failed")


def finish_run(
    run_id: str,
    status: str = "success",
    message: str = "Task complete",
    details: dict[str, Any] | None = None,
) -> None:
    try:
        _RECORDER.finish_run(run_id, status=status, message=message, details=details)
    except Exception:
        _safe_log("Task debug finish_run failed")


def get_events() -> list[TaskDebugEvent]:
    try:
        return _RECORDER.get_events()
    except Exception:
        _safe_log("Task debug get_events failed")
        return []


def clear_events() -> None:
    try:
        _RECORDER.clear_events()
    except Exception:
        _safe_log("Task debug clear_events failed")


def get_trace_dir() -> Path:
    try:
        return _RECORDER.get_trace_dir()
    except Exception:
        _safe_log("Task debug get_trace_dir failed")
        return _default_trace_dir()


def get_bridge() -> TaskDebugBridge:
    try:
        return _RECORDER.get_bridge()
    except Exception:
        _safe_log("Task debug get_bridge failed")
        return TaskDebugBridge()


def reset_for_tests(
    trace_dir: str | Path | None = None,
    max_events: int | None = None,
) -> None:
    try:
        _RECORDER.reset_for_tests(trace_dir=trace_dir, max_events=max_events)
    except Exception:
        _safe_log("Task debug reset_for_tests failed")


def _default_trace_dir() -> Path:
    return Path(__file__).resolve().parent.parent / "logs" / "task_debug"


def _trace_filename(run_id: str, task_id: str) -> str:
    safe_task_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(task_id)).strip("._")
    if not safe_task_id:
        safe_task_id = "task"
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    return f"{stamp}_{safe_task_id}_{run_id[:12]}.jsonl"


def _elapsed_ms(started_at: Any) -> int:
    if not isinstance(started_at, (int, float)):
        return 0
    return max(0, int((time.perf_counter() - started_at) * 1000))


def _safe_details(details: dict[str, Any] | None) -> dict[str, Any]:
    if not details:
        return {}
    if not isinstance(details, dict):
        return {"value": _make_json_safe(details, set())}
    return _make_json_safe(details, set())


def _redact_or_clean(key: str, value: Any, seen: set[int]) -> Any:
    if _SENSITIVE_KEY_RE.search(key):
        return "[REDACTED]"
    return _make_json_safe(value, seen)


def _make_json_safe(value: Any, seen: set[int]) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        if isinstance(value, float) and not math.isfinite(value):
            return str(value)
        return value
    if isinstance(value, dict):
        value_id = id(value)
        if value_id in seen:
            return "[CYCLE]"
        seen.add(value_id)
        try:
            return {
                str(key): _redact_or_clean(str(key), item, seen)
                for key, item in value.items()
            }
        finally:
            seen.remove(value_id)
    if isinstance(value, (list, tuple)):
        value_id = id(value)
        if value_id in seen:
            return "[CYCLE]"
        seen.add(value_id)
        try:
            return [_make_json_safe(item, seen) for item in value]
        finally:
            seen.remove(value_id)
    try:
        json.dumps(value, allow_nan=False)
        return value
    except Exception:
        return str(value)


def _safe_log(message: str) -> None:
    try:
        logger.warning(message, exc_info=True)
    except Exception:
        pass


_RECORDER = _TaskDebugRecorder()
