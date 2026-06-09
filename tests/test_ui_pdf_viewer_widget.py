"""Tests for shared PDF viewer page navigation behavior."""
from __future__ import annotations

from icharlotte_core.ui.pdf_viewer_widget import PdfViewerWidget


class _FakePage:
    def __init__(self) -> None:
        self.javascript_calls: list[str] = []

    def runJavaScript(self, js: str, callback=None) -> None:
        self.javascript_calls.append(js)
        if callback is not None:
            callback(None)


class _FakeWebView:
    def __init__(self) -> None:
        self.fake_page = _FakePage()

    def page(self) -> _FakePage:
        return self.fake_page


class _FakeLabel:
    def __init__(self) -> None:
        self.text = ""

    def setText(self, text: str) -> None:
        self.text = text


class _FakeSpinBox:
    def __init__(self) -> None:
        self.maximum = None
        self.value = None
        self.signal_blocks: list[bool] = []

    def setMaximum(self, maximum: int) -> None:
        self.maximum = maximum

    def blockSignals(self, blocked: bool) -> None:
        self.signal_blocks.append(blocked)

    def setValue(self, value: int) -> None:
        self.value = value


def _viewer_double() -> PdfViewerWidget:
    viewer = PdfViewerWidget.__new__(PdfViewerWidget)
    viewer.current_page = 1
    viewer.total_pages = 0
    viewer._viewer_ready = True
    viewer._pending_pdf = None
    viewer._pending_page_number = None
    viewer._nav_cooldown_until = 0
    viewer._load_poll_count = 0
    viewer.web_view = _FakeWebView()
    viewer.page_label = _FakeLabel()
    viewer.page_spin = _FakeSpinBox()
    return viewer


def test_page_request_before_pdf_page_count_is_known_runs_after_load() -> None:
    viewer = _viewer_double()

    viewer.go_to_page(599)

    assert viewer._pending_page_number == 599
    assert viewer.web_view.fake_page.javascript_calls == []
    assert viewer.current_page == 1

    viewer._on_total_pages_result(2530)

    assert viewer.web_view.fake_page.javascript_calls == [
        "window.pdfViewer.goToPage(599)"
    ]
    assert viewer.current_page == 599
    assert viewer.page_label.text == "Page: 599 / 2530"
    assert viewer.page_spin.maximum == 2530
    assert viewer.page_spin.value == 599
    assert viewer._pending_page_number is None
