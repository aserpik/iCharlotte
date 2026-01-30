"""PDF Viewer widget with page tracking using pdf.js"""
import os
import tempfile
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox
from PySide6.QtCore import QUrl, Signal, QTimer
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineSettings


class PdfViewerWidget(QWidget):
    """PDF viewer with automatic page tracking using embedded pdf.js."""

    pageChanged = Signal(int)  # Emitted when current page changes

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_page = 1
        self.total_pages = 0
        self.current_pdf_path = None
        self._viewer_ready = False
        self._pending_pdf = None
        self._setup_ui()
        self._start_page_polling()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Page info bar
        info_layout = QHBoxLayout()
        self.page_label = QLabel("Page: - / -")
        self.page_label.setStyleSheet("font-weight: bold; padding: 2px 5px;")
        info_layout.addWidget(self.page_label)

        # Go to page control
        info_layout.addWidget(QLabel("Go to:"))
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setMaximum(1)
        self.page_spin.valueChanged.connect(self._on_spin_changed)
        info_layout.addWidget(self.page_spin)

        info_layout.addStretch()
        layout.addLayout(info_layout)

        # PDF viewer (pdf.js embedded)
        self.web_view = QWebEngineView()
        settings = self.web_view.settings()
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, True)

        # Connect load finished signal
        self.web_view.loadFinished.connect(self._on_viewer_loaded)

        # Load the pdf.js viewer HTML
        self._load_viewer()
        layout.addWidget(self.web_view)

    def _load_viewer(self):
        """Load the pdf.js viewer HTML."""
        # Use a file:// base URL to allow pdf.js to access local files
        # This is necessary because about:blank origin cannot access file:// URLs
        base_url = QUrl.fromLocalFile(tempfile.gettempdir() + "/")
        self.web_view.setHtml(self._get_viewer_html(), base_url)

    def _on_viewer_loaded(self, success):
        """Called when the viewer HTML is loaded."""
        if success:
            self._viewer_ready = True
            # If there's a pending PDF to load, load it now
            if self._pending_pdf:
                self._do_load_pdf(self._pending_pdf)
                self._pending_pdf = None

    def _get_viewer_html(self):
        """Return inline pdf.js viewer HTML."""
        return '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { width: 100%; height: 100%; overflow: hidden; background: #525659; }
        #viewerContainer {
            width: 100%;
            height: 100%;
            overflow: auto;
            background: #525659;
        }
        .page {
            margin: 10px auto;
            background: white;
            box-shadow: 0 2px 5px rgba(0,0,0,0.3);
            position: relative;
        }
        canvas { display: block; }

        /* Text layer for selection */
        .textLayer {
            position: absolute;
            left: 0;
            top: 0;
            right: 0;
            bottom: 0;
            overflow: hidden;
            opacity: 0.2;
            line-height: 1.0;
            user-select: text;
            -webkit-user-select: text;
            -moz-user-select: text;
            -ms-user-select: text;
        }

        .textLayer > span {
            color: transparent;
            position: absolute;
            white-space: pre;
            cursor: text;
            transform-origin: 0% 0%;
        }

        .textLayer ::selection {
            background: rgba(0, 0, 255, 0.3);
        }

        .textLayer ::-moz-selection {
            background: rgba(0, 0, 255, 0.3);
        }

        #loading {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            color: white;
            font-family: Arial, sans-serif;
            font-size: 16px;
        }
        #error {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            color: #ff6b6b;
            font-family: Arial, sans-serif;
            font-size: 14px;
            text-align: center;
            max-width: 80%;
            display: none;
        }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
</head>
<body>
    <div id="loading">Loading PDF viewer...</div>
    <div id="error"></div>
    <div id="viewerContainer"></div>
    <script>
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

        let pdfDoc = null;
        let currentPage = 1;
        let totalPages = 0;
        let scale = 1.5;
        let renderedPages = new Set();
        let isLoading = false;

        document.getElementById('loading').textContent = 'Ready to load PDF';

        async function loadPdf(url) {
            if (isLoading) {
                console.log('Already loading a PDF, please wait...');
                return { success: false, error: 'Already loading' };
            }

            isLoading = true;
            document.getElementById('loading').style.display = 'block';
            document.getElementById('loading').textContent = 'Loading PDF...';
            document.getElementById('error').style.display = 'none';

            try {
                console.log('Loading PDF from:', url);
                const loadingTask = pdfjsLib.getDocument(url);
                pdfDoc = await loadingTask.promise;
                totalPages = pdfDoc.numPages;
                currentPage = 1;
                renderedPages.clear();

                const container = document.getElementById('viewerContainer');
                container.innerHTML = '';

                document.getElementById('loading').textContent = 'Rendering pages...';

                // Render all pages
                for (let i = 1; i <= totalPages; i++) {
                    await renderPage(i);
                    document.getElementById('loading').textContent = 'Rendering page ' + i + ' of ' + totalPages + '...';
                }

                // Setup scroll listener for page tracking
                setupScrollListener();

                document.getElementById('loading').style.display = 'none';
                isLoading = false;

                console.log('PDF loaded successfully:', totalPages, 'pages');
                return { success: true, totalPages: totalPages };
            } catch (err) {
                console.error('Error loading PDF:', err);
                document.getElementById('loading').style.display = 'none';
                document.getElementById('error').textContent = 'Error loading PDF: ' + err.message;
                document.getElementById('error').style.display = 'block';
                isLoading = false;
                return { success: false, error: err.message };
            }
        }

        async function renderPage(pageNum) {
            const page = await pdfDoc.getPage(pageNum);
            const viewport = page.getViewport({ scale: scale });

            const pageDiv = document.createElement('div');
            pageDiv.className = 'page';
            pageDiv.id = 'page-' + pageNum;
            pageDiv.dataset.pageNumber = pageNum;
            pageDiv.style.width = viewport.width + 'px';
            pageDiv.style.height = viewport.height + 'px';

            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            canvas.width = viewport.width;
            canvas.height = viewport.height;

            pageDiv.appendChild(canvas);
            document.getElementById('viewerContainer').appendChild(pageDiv);

            // Render canvas
            await page.render({ canvasContext: ctx, viewport: viewport }).promise;

            // Add text layer for selection
            const textContent = await page.getTextContent();
            const textLayerDiv = document.createElement('div');
            textLayerDiv.className = 'textLayer';
            textLayerDiv.style.width = viewport.width + 'px';
            textLayerDiv.style.height = viewport.height + 'px';
            pageDiv.appendChild(textLayerDiv);

            // Render text layer
            pdfjsLib.renderTextLayer({
                textContentSource: textContent,
                container: textLayerDiv,
                viewport: viewport,
                textDivs: []
            });

            renderedPages.add(pageNum);
        }

        function setupScrollListener() {
            const container = document.getElementById('viewerContainer');
            container.addEventListener('scroll', () => {
                const pages = document.querySelectorAll('.page');
                const containerRect = container.getBoundingClientRect();
                const containerCenter = containerRect.top + containerRect.height / 2;

                let closestPage = 1;
                let closestDist = Infinity;

                pages.forEach(page => {
                    const rect = page.getBoundingClientRect();
                    const pageCenter = rect.top + rect.height / 2;
                    const dist = Math.abs(pageCenter - containerCenter);
                    if (dist < closestDist) {
                        closestDist = dist;
                        closestPage = parseInt(page.dataset.pageNumber);
                    }
                });

                currentPage = closestPage;
            });
        }

        function goToPage(pageNum) {
            if (pageNum < 1 || pageNum > totalPages) return false;
            const pageEl = document.getElementById('page-' + pageNum);
            if (pageEl) {
                pageEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                currentPage = pageNum;
                return true;
            }
            return false;
        }

        function getCurrentPage() {
            return currentPage;
        }

        function getTotalPages() {
            return totalPages;
        }

        function setZoom(newScale) {
            scale = newScale;
            // Re-render would be needed for zoom changes
        }

        function isReady() {
            return true;
        }

        // Expose API
        window.pdfViewer = {
            loadPdf: loadPdf,
            goToPage: goToPage,
            getCurrentPage: getCurrentPage,
            getTotalPages: getTotalPages,
            setZoom: setZoom,
            isReady: isReady
        };

        console.log('PDF viewer initialized');
    </script>
</body>
</html>'''

    def _start_page_polling(self):
        """Poll for current page changes."""
        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self._poll_current_page)
        self.poll_timer.start(500)  # Poll every 500ms

    def _poll_current_page(self):
        """Query current page from pdf.js."""
        if not self._viewer_ready:
            return
        self.web_view.page().runJavaScript(
            "window.pdfViewer ? window.pdfViewer.getCurrentPage() : 0",
            self._on_page_result
        )

    def _on_page_result(self, page):
        """Handle page query result."""
        if page and page != self.current_page and page > 0:
            self.current_page = page
            self.page_label.setText(f"Page: {page} / {self.total_pages}")
            self.page_spin.blockSignals(True)
            self.page_spin.setValue(page)
            self.page_spin.blockSignals(False)
            self.pageChanged.emit(page)

    def _on_spin_changed(self, value):
        """Navigate to page when spin box changes."""
        self.go_to_page(value)

    def load_pdf(self, path):
        """Load a PDF file."""
        self.current_pdf_path = path

        if not self._viewer_ready:
            # Store for later loading once viewer is ready
            self._pending_pdf = path
            return

        self._do_load_pdf(path)

    def _do_load_pdf(self, path):
        """Actually load the PDF into the viewer."""
        # Convert path to file:// URL with proper escaping
        url = QUrl.fromLocalFile(path).toString()
        # Escape single quotes in the URL for JavaScript
        url_escaped = url.replace("'", "\\'")
        js = f"window.pdfViewer.loadPdf('{url_escaped}')"
        self.web_view.page().runJavaScript(js, self._on_pdf_loaded)

    def _on_pdf_loaded(self, result):
        """Handle PDF load result."""
        if result and isinstance(result, dict) and result.get('success'):
            self.total_pages = result.get('totalPages', 0)
            self.page_spin.setMaximum(max(1, self.total_pages))
            self.page_label.setText(f"Page: 1 / {self.total_pages}")
            self.current_page = 1

    def go_to_page(self, page_num):
        """Navigate to a specific page."""
        if not self._viewer_ready:
            return
        js = f"window.pdfViewer.goToPage({page_num})"
        self.web_view.page().runJavaScript(js)
        self.current_page = page_num

    def get_current_page(self):
        """Get the current page number (cached value)."""
        return self.current_page

    def get_total_pages(self):
        """Get the total number of pages."""
        return self.total_pages
