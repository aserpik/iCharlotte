"""Shared file drag/drop support for wizard settings pages."""

from __future__ import annotations

import os
from collections.abc import Callable

from PySide6.QtCore import QEvent, QObject
from PySide6.QtWidgets import QWidget


def local_file_paths_from_mime_data(mime_data) -> list[str]:
    """Return existing local files from a Qt mime payload, preserving order."""
    if mime_data is None or not mime_data.hasUrls():
        return []

    paths: list[str] = []
    seen: set[str] = set()
    for url in mime_data.urls():
        if not url.isLocalFile():
            continue
        path = os.path.normpath(url.toLocalFile())
        key = os.path.normcase(path)
        if not path or key in seen or not os.path.isfile(path):
            continue
        paths.append(path)
        seen.add(key)
    return paths


class FileDropTarget(QObject):
    """Event filter that turns local-file drops into a callback."""

    def __init__(
        self,
        on_files_dropped: Callable[[list[str]], None],
        parent: QObject | None = None,
        *,
        path_filter: Callable[[str], bool] | None = None,
    ):
        super().__init__(parent)
        self._on_files_dropped = on_files_dropped
        self._path_filter = path_filter

    def eventFilter(self, watched, event) -> bool:  # noqa: N802 - Qt override
        event_type = event.type()
        if event_type not in (
            QEvent.Type.DragEnter,
            QEvent.Type.DragMove,
            QEvent.Type.Drop,
        ):
            return False

        paths = local_file_paths_from_mime_data(event.mimeData())
        if self._path_filter is not None:
            paths = [path for path in paths if self._path_filter(path)]
        if not paths:
            event.ignore()
            return True

        event.acceptProposedAction()
        if event_type == QEvent.Type.Drop:
            self._on_files_dropped(paths)
        return True


def enable_file_drop(
    widget: QWidget,
    on_files_dropped: Callable[[list[str]], None],
    *,
    path_filter: Callable[[str], bool] | None = None,
) -> FileDropTarget:
    """Install a local-file drop handler on ``widget`` and keep it alive."""
    target = FileDropTarget(on_files_dropped, widget, path_filter=path_filter)

    def install(receiver: QWidget) -> None:
        receiver.setAcceptDrops(True)
        receiver.installEventFilter(target)
        handlers = list(getattr(receiver, "_icharlotte_file_drop_handlers", []))
        handlers.append(target)
        setattr(receiver, "_icharlotte_file_drop_handlers", handlers)

    install(widget)
    viewport = widget.viewport() if hasattr(widget, "viewport") else None
    if viewport is not None and viewport is not widget:
        install(viewport)
    return target
