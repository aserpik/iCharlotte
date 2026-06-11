"""Small helpers for wizard task debug instrumentation."""

from __future__ import annotations

from typing import Any

from icharlotte_core import task_debug


def start_debug_run(
    owner: Any,
    *,
    source: str,
    details: dict[str, Any] | None = None,
) -> None:
    if getattr(owner, "_debug_run_id", None):
        return
    spec = getattr(owner, "_spec", None)
    owner._debug_run_id = task_debug.start_run(
        task_id=str(getattr(spec, "task_id", "")),
        task_title=str(getattr(spec, "title", "")),
        source=source,
        details=dict(details or {}),
    )


def emit_debug(
    owner: Any,
    *,
    phase: str,
    message: str,
    level: str = "info",
    source: str,
    details: dict[str, Any] | None = None,
) -> None:
    run_id = getattr(owner, "_debug_run_id", None)
    if not run_id:
        return
    spec = getattr(owner, "_spec", None)
    task_debug.emit_event(
        run_id,
        str(getattr(spec, "task_id", "")),
        str(getattr(spec, "title", "")),
        phase=phase,
        message=str(message),
        level=level,
        source=source,
        details=dict(details or {}),
    )


def finish_debug_run(
    owner: Any,
    *,
    status: str,
    message: str,
    details: dict[str, Any] | None = None,
) -> None:
    run_id = getattr(owner, "_debug_run_id", None)
    if not run_id:
        return
    task_debug.finish_run(run_id, status=status, message=message, details=details)
    owner._debug_run_id = None


def record_status(
    owner: Any,
    message: str,
    *,
    source: str,
    phase: str = "status",
    details: dict[str, Any] | None = None,
) -> None:
    text = str(message)
    upper = text.upper()
    level = "warning" if upper.startswith("WARNING") else "info"
    if upper.startswith("FAILED") or "ERROR" in upper:
        level = "error"
    emit_debug(
        owner,
        phase=phase,
        message=text,
        level=level,
        source=source,
        details=details,
    )


def record_progress(
    owner: Any,
    value: int,
    *,
    source: str,
) -> None:
    try:
        pct = int(value)
    except (TypeError, ValueError):
        pct = 0
    emit_debug(
        owner,
        phase="progress",
        message=f"Progress {pct}%",
        source=source,
        details={"progress": pct},
    )
