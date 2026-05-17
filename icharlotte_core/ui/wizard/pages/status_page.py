"""StatusPage — progress bar + log + Cancel button while a task is running."""
import time
from collections import deque

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class StatusPage(QWidget):
    """Shows progress + log lines. Emits cancel_requested when Cancel is clicked."""

    cancel_requested = Signal()
    configure_requested = Signal()   # emitted when "Configure Topics & Continue" is clicked

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        # 1. Status label
        self.status_label = QLabel("Starting…")
        self.status_label.setStyleSheet("font-size: 13px; font-weight: 600;")
        outer.addWidget(self.status_label)

        # 2. Progress row: bar + percent + ETA
        progress_row = QHBoxLayout()
        progress_row.setSpacing(8)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_row.addWidget(self.progress_bar, 1)

        self.percent_label = QLabel("")
        self.percent_label.setStyleSheet("font-weight: 600; min-width: 40px;")
        progress_row.addWidget(self.percent_label)

        self.eta_label = QLabel("")
        self.eta_label.setStyleSheet("color: #666; font-size: 11px;")
        progress_row.addWidget(self.eta_label)

        outer.addLayout(progress_row)

        # 3. Log toggle button
        self.log_toggle_btn = QPushButton("Show details ▾")
        self.log_toggle_btn.setCheckable(True)
        self.log_toggle_btn.setChecked(False)
        self.log_toggle_btn.setStyleSheet(
            "QPushButton { background: #f5f5f5; border: 1px solid #ccc; padding: 4px 10px;"
            " border-radius: 3px; text-align: left; }"
            "QPushButton:hover { background: #e8e8e8; }"
        )
        self.log_toggle_btn.setMaximumWidth(140)
        self.log_toggle_btn.toggled.connect(self._on_log_toggle_clicked)
        outer.addWidget(self.log_toggle_btn)

        # 4. Log view (hidden by default)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setStyleSheet("font-family: Consolas, 'Courier New', monospace; font-size: 11px;")
        self.log_view.setVisible(False)
        outer.addWidget(self.log_view, 1)

        # --- Awaiting-input widget (hidden until Phase 1 of depo flow completes) ---
        self.awaiting_input_widget = QWidget()
        ai_layout = QHBoxLayout(self.awaiting_input_widget)
        ai_layout.setContentsMargins(0, 4, 0, 4)
        ai_label = QLabel("Topics extracted — pick which to summarize.")
        ai_label.setStyleSheet("font-weight: bold; font-size: 13px;")
        ai_layout.addWidget(ai_label)
        ai_layout.addStretch()
        self.configure_btn = QPushButton("Configure Topics & Continue")
        self.configure_btn.setStyleSheet(
            "background-color: #1565c0; color: white; font-weight: 600;"
            " padding: 8px 20px; border-radius: 4px;"
        )
        self.configure_btn.clicked.connect(self.configure_requested)
        ai_layout.addWidget(self.configure_btn)
        self.awaiting_input_widget.setVisible(False)
        outer.addWidget(self.awaiting_input_widget)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(
            "background-color: #d32f2f; color: white; font-weight: 600; padding: 8px 20px; border-radius: 4px;"
        )
        self.cancel_btn.clicked.connect(self._on_cancel)
        btn_row.addWidget(self.cancel_btn)
        outer.addLayout(btn_row)

        # ETA tracking state
        self._run_start_monotonic: float | None = None
        self._last_pct: int = 0
        self._paused_elapsed: float = 0.0
        self._pause_start: float | None = None
        self._rate_window: deque = deque()  # (monotonic, pct) pairs over last 30s

    # ---- Toggle ----

    def _on_log_toggle_clicked(self, checked: bool) -> None:
        self.log_view.setVisible(checked)
        self.log_toggle_btn.setText("Hide details ▴" if checked else "Show details ▾")

    # ---- Reset ----

    def reset(self) -> None:
        self.status_label.setText("Starting…")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.percent_label.setText("")
        self.eta_label.setText("")
        self.log_view.clear()
        self.log_view.setVisible(False)
        self.log_toggle_btn.setChecked(False)
        self.log_toggle_btn.setText("Show details ▾")
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setText("Cancel")
        self.awaiting_input_widget.setVisible(False)
        self._run_start_monotonic = None
        self._last_pct = 0
        self._paused_elapsed = 0.0
        self._pause_start = None
        self._rate_window.clear()

    # ---- ETA pause/resume ----

    def pause_eta(self) -> None:
        """Pause ETA accumulation (e.g. while user is in a config dialog)."""
        if self._pause_start is None and self._run_start_monotonic is not None:
            self._pause_start = time.monotonic()

    def resume_eta(self) -> None:
        """Resume ETA accumulation after a pause."""
        if self._pause_start is not None:
            self._paused_elapsed += time.monotonic() - self._pause_start
            self._pause_start = None

    # ---- Awaiting-input mode ----

    def show_awaiting_input(self, session_path: str) -> None:
        """Switch to awaiting-input mode after Phase 1 of a deposition run."""
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(1)
        self.status_label.setText("Phase 1 complete — awaiting topic selection")
        self.awaiting_input_widget.setVisible(True)

    # ---- Slots / public API for the worker connection ----

    def on_status(self, line: str) -> None:
        self.status_label.setText(line)
        self.log_view.appendPlainText(line)
        self.log_view.verticalScrollBar().setValue(self.log_view.verticalScrollBar().maximum())

    def on_progress(self, pct: int) -> None:
        if pct < 0:
            pct = 0
        if pct > 100:
            pct = 100
        # Start the ETA timer on the first real progress event.
        if self._run_start_monotonic is None and pct > 0:
            self._run_start_monotonic = time.monotonic()
        # Record in rolling window and trim entries older than 30s.
        now = time.monotonic()
        self._rate_window.append((now, pct))
        cutoff = now - 30.0
        while self._rate_window and self._rate_window[0][0] < cutoff:
            self._rate_window.popleft()
        if self.progress_bar.maximum() == 0:
            self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(pct)
        self.percent_label.setText(f"{pct}%")
        self.eta_label.setText(self._format_eta(pct))
        self._last_pct = pct

    # ---- ETA helpers ----

    def _format_eta(self, pct: int) -> str:
        # Skip ETA for the noisy early phase (< 5%) and when finished.
        if self._run_start_monotonic is None or pct < 5 or pct >= 100:
            return ""
        now = time.monotonic()
        elapsed = now - self._run_start_monotonic - self._paused_elapsed
        if self._pause_start is not None:
            elapsed -= (now - self._pause_start)
        if elapsed < 1.0:
            return ""
        # Prefer rolling-window rate (last 30s) when we have enough data.
        rate: float = 0.0
        if len(self._rate_window) >= 2:
            oldest_t, oldest_pct = self._rate_window[0]
            newest_t, newest_pct = self._rate_window[-1]
            window_span = newest_t - oldest_t
            pct_delta = newest_pct - oldest_pct
            if window_span >= 5.0 and pct_delta > 0:
                rate = pct_delta / window_span
        # Fall back to lifetime rate when window is too sparse.
        if rate <= 0 and elapsed > 0:
            rate = pct / elapsed
        if rate <= 0:
            return ""
        remaining = max(0.0, (100 - pct) / rate)
        return f"~{self._format_duration(remaining)} remaining"

    @staticmethod
    def _format_duration(seconds: float) -> str:
        seconds = int(round(seconds))
        if seconds < 60:
            return f"{seconds}s"
        m, s = divmod(seconds, 60)
        if m < 60:
            return f"{m}m {s}s"
        h, m = divmod(m, 60)
        return f"{h}h {m}m"

    def _on_cancel(self) -> None:
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setText("Cancelling…")
        self.cancel_requested.emit()
