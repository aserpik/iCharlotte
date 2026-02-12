"""PDF Viewer widget with page tracking using pdf.js"""
import os
import tempfile
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox
from PySide6.QtCore import QUrl, Signal, QTimer
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineSettings


class PdfViewerWidget(QWidget):
    """PDF viewer with automatic page tracking using embedded pdf.js."""

    pageChanged = Signal(int)     # Emitted when current page changes
    annotationDeleteRequested = Signal(str)  # Emitted with exchange ID when Delete pressed on highlight

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
        """Return inline pdf.js viewer HTML with lazy page rendering."""
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
            background: #e0e0e0;
            box-shadow: 0 2px 5px rgba(0,0,0,0.3);
            position: relative;
        }
        .page.rendered { background: white; }
        .page-num-label {
            position: absolute; top: 50%; left: 50%;
            transform: translate(-50%, -50%);
            color: #999; font: 14px Arial; pointer-events: none;
        }
        .page.rendered .page-num-label { display: none; }
        canvas { display: block; }
        .textLayer {
            position: absolute; left: 0; top: 0; right: 0; bottom: 0;
            overflow: hidden; opacity: 0.2; line-height: 1.0;
            user-select: text; z-index: 2;
        }
        .textLayer > span {
            color: transparent; position: absolute; white-space: pre;
            cursor: text; transform-origin: 0% 0%;
        }
        .textLayer ::selection { background: rgba(0,0,255,0.3); }
        .annot-overlay {
            position: absolute; cursor: pointer; border: 2px solid transparent;
            border-radius: 2px; z-index: 5;
        }
        .annot-overlay:hover { border-color: rgba(0,0,0,0.3); }
        .annot-overlay.selected { border-color: #e00; border-style: dashed; }
        #loading {
            position: absolute; top: 50%; left: 50%;
            transform: translate(-50%,-50%);
            color: white; font: 16px Arial;
        }
        #error {
            position: absolute; top: 50%; left: 50%;
            transform: translate(-50%,-50%);
            color: #ff6b6b; font: 14px Arial; text-align: center;
            max-width: 80%; display: none;
        }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
</head>
<body>
    <div id="loading">Ready to load PDF</div>
    <div id="error"></div>
    <div id="viewerContainer"></div>
    <script>
        pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

        let pdfDoc = null, currentPage = 1, totalPages = 0, scale = 1.5;
        let renderedPages = new Set(), renderingPages = new Set();
        let isLoading = false, pageHeights = [];
        const BUFFER = 2; // render this many pages above/below viewport

        // Annotation selection state
        let selectedAnnotExId = null;
        let pendingDeleteExId = null;  // Set by Delete key, polled by Python

        async function loadPdf(url) {
            if (isLoading) return { success: false, error: 'Already loading' };
            isLoading = true;

            const loadEl = document.getElementById('loading');
            const errEl = document.getElementById('error');
            loadEl.style.display = 'block';
            loadEl.textContent = 'Loading PDF...';
            errEl.style.display = 'none';

            try {
                // Clean up previous document
                if (pdfDoc) { pdfDoc.destroy(); pdfDoc = null; }
                renderedPages.clear();
                renderingPages.clear();
                pageHeights = [];

                pdfDoc = await pdfjsLib.getDocument(url).promise;
                totalPages = pdfDoc.numPages;
                currentPage = 1;

                const container = document.getElementById('viewerContainer');
                container.innerHTML = '';

                // Create placeholder divs for all pages (measure first page for size)
                const firstPage = await pdfDoc.getPage(1);
                const vp = firstPage.getViewport({ scale });
                const pw = vp.width, ph = vp.height;

                for (let i = 1; i <= totalPages; i++) {
                    const div = document.createElement('div');
                    div.className = 'page';
                    div.id = 'page-' + i;
                    div.dataset.pageNumber = i;
                    div.style.width = pw + 'px';
                    div.style.height = ph + 'px';
                    const lbl = document.createElement('div');
                    lbl.className = 'page-num-label';
                    lbl.textContent = 'Page ' + i;
                    div.appendChild(lbl);
                    container.appendChild(div);
                    pageHeights.push(ph);
                }

                container.addEventListener('scroll', onScroll);
                loadEl.style.display = 'none';
                isLoading = false;

                // Render initial visible pages
                await renderVisible();
                return { success: true, totalPages };
            } catch (err) {
                loadEl.style.display = 'none';
                errEl.textContent = 'Error: ' + err.message;
                errEl.style.display = 'block';
                isLoading = false;
                return { success: false, error: err.message };
            }
        }

        let scrollTimer = null;
        function onScroll() {
            // Track current page
            updateCurrentPage();
            // Debounced lazy render
            if (scrollTimer) clearTimeout(scrollTimer);
            scrollTimer = setTimeout(renderVisible, 100);
        }

        function updateCurrentPage() {
            const container = document.getElementById('viewerContainer');
            const ct = container.scrollTop + container.clientHeight / 2;
            const pages = container.children;
            for (let i = 0; i < pages.length; i++) {
                const top = pages[i].offsetTop;
                const bot = top + pages[i].offsetHeight;
                if (ct >= top && ct < bot) {
                    currentPage = i + 1;
                    return;
                }
            }
        }

        async function renderVisible() {
            if (!pdfDoc) return;
            const container = document.getElementById('viewerContainer');
            const scrollTop = container.scrollTop;
            const viewH = container.clientHeight;

            // Find visible page range
            let firstVisible = 1, lastVisible = totalPages;
            const pages = container.children;
            for (let i = 0; i < pages.length; i++) {
                const top = pages[i].offsetTop;
                const bot = top + pages[i].offsetHeight;
                if (bot >= scrollTop) { firstVisible = i + 1; break; }
            }
            for (let i = pages.length - 1; i >= 0; i--) {
                const top = pages[i].offsetTop;
                if (top <= scrollTop + viewH) { lastVisible = i + 1; break; }
            }

            const lo = Math.max(1, firstVisible - BUFFER);
            const hi = Math.min(totalPages, lastVisible + BUFFER);

            // Unload pages far from viewport to free memory
            for (const pn of renderedPages) {
                if (pn < lo - BUFFER || pn > hi + BUFFER) {
                    unloadPage(pn);
                }
            }

            // Render pages in range
            for (let i = lo; i <= hi; i++) {
                if (!renderedPages.has(i) && !renderingPages.has(i)) {
                    await renderPage(i);
                }
            }
        }

        function unloadPage(pageNum) {
            const div = document.getElementById('page-' + pageNum);
            if (!div) return;
            // Keep the placeholder div but remove rendered content
            const w = div.style.width, h = div.style.height;
            div.innerHTML = '';
            div.className = 'page';
            div.style.width = w;
            div.style.height = h;
            const lbl = document.createElement('div');
            lbl.className = 'page-num-label';
            lbl.textContent = 'Page ' + pageNum;
            div.appendChild(lbl);
            renderedPages.delete(pageNum);
        }

        async function renderPage(pageNum) {
            if (renderedPages.has(pageNum) || renderingPages.has(pageNum)) return;
            renderingPages.add(pageNum);

            try {
                const page = await pdfDoc.getPage(pageNum);
                const viewport = page.getViewport({ scale });
                const pageDiv = document.getElementById('page-' + pageNum);
                if (!pageDiv) { renderingPages.delete(pageNum); return; }

                pageDiv.style.width = viewport.width + 'px';
                pageDiv.style.height = viewport.height + 'px';
                pageDiv.innerHTML = '';
                pageDiv.classList.add('rendered');

                // Canvas
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                pageDiv.appendChild(canvas);
                await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

                // Highlight annotations (drawn on canvas + clickable overlays)
                try {
                    const annots = await page.getAnnotations();
                    const highlights = annots.filter(a => a.subtype === 'Highlight');
                    if (highlights.length > 0) {
                        const ctx = canvas.getContext('2d');
                        for (const a of highlights) {
                            const c = a.color || {0:255, 1:255, 2:0};
                            ctx.fillStyle = 'rgba(' + (c[0]||255) + ',' + (c[1]||255) + ',' + (c[2]||0) + ',0.3)';

                            // Extract exchange ID from annotation content (format: "ex:123")
                            const exId = (a.contentsObj && a.contentsObj.str || '').replace('ex:', '');

                            if (a.quadPoints && a.quadPoints.length) {
                                for (let q = 0; q < a.quadPoints.length; q += 8) {
                                    const pts = a.quadPoints.slice(q, q + 8);
                                    const xs = [pts[0], pts[2], pts[4], pts[6]];
                                    const ys = [pts[1], pts[3], pts[5], pts[7]];
                                    const minX = Math.min(...xs), maxX = Math.max(...xs);
                                    const minY = Math.min(...ys), maxY = Math.max(...ys);
                                    const [vx1, vy1] = viewport.convertToViewportPoint(minX, maxY);
                                    const [vx2, vy2] = viewport.convertToViewportPoint(maxX, minY);
                                    ctx.fillRect(vx1, vy1, vx2 - vx1, vy2 - vy1);
                                    if (exId) createAnnotOverlay(pageDiv, vx1, vy1, vx2-vx1, vy2-vy1, exId);
                                }
                            } else {
                                const r = a.rect;
                                const [x1, y1] = viewport.convertToViewportPoint(r[0], r[3]);
                                const [x2, y2] = viewport.convertToViewportPoint(r[2], r[1]);
                                ctx.fillRect(x1, y1, x2 - x1, y2 - y1);
                                if (exId) createAnnotOverlay(pageDiv, x1, y1, x2-x1, y2-y1, exId);
                            }
                        }
                    }
                } catch(e) {
                    console.log('Highlight error page ' + pageNum + ': ' + e.message);
                }

                // Text layer
                const tc = await page.getTextContent();
                const tDiv = document.createElement('div');
                tDiv.className = 'textLayer';
                tDiv.style.width = viewport.width + 'px';
                tDiv.style.height = viewport.height + 'px';
                tDiv.style.setProperty('--scale-factor', scale);
                pageDiv.appendChild(tDiv);
                pdfjsLib.renderTextLayer({
                    textContentSource: tc, container: tDiv,
                    viewport: viewport, textDivs: []
                });

                renderedPages.add(pageNum);
            } finally {
                renderingPages.delete(pageNum);
            }
        }

        function createAnnotOverlay(pageDiv, x, y, w, h, exId) {
            const div = document.createElement('div');
            div.className = 'annot-overlay';
            div.dataset.exId = exId;
            div.style.left = x + 'px';
            div.style.top = y + 'px';
            div.style.width = w + 'px';
            div.style.height = h + 'px';
            div.addEventListener('click', function(e) {
                e.stopPropagation();
                // Deselect all, select this exchange's overlays
                document.querySelectorAll('.annot-overlay.selected').forEach(el => el.classList.remove('selected'));
                document.querySelectorAll('.annot-overlay[data-ex-id="' + exId + '"]').forEach(el => el.classList.add('selected'));
                selectedAnnotExId = exId;
            });
            pageDiv.appendChild(div);
        }

        // Click on empty area deselects
        document.getElementById('viewerContainer').addEventListener('click', function(e) {
            if (!e.target.classList.contains('annot-overlay')) {
                document.querySelectorAll('.annot-overlay.selected').forEach(el => el.classList.remove('selected'));
                selectedAnnotExId = null;
            }
        });

        // Delete key triggers annotation deletion
        document.addEventListener('keydown', function(e) {
            if ((e.key === 'Delete' || e.key === 'Backspace') && selectedAnnotExId) {
                pendingDeleteExId = selectedAnnotExId;
                // Remove overlays visually immediately
                document.querySelectorAll('.annot-overlay[data-ex-id="' + selectedAnnotExId + '"]').forEach(el => el.remove());
                selectedAnnotExId = null;
            }
        });

        function getSelectedAnnotation() { return selectedAnnotExId; }
        function getPendingDelete() {
            const id = pendingDeleteExId;
            pendingDeleteExId = null;
            return id;
        }

        function goToPage(pageNum) {
            if (pageNum < 1 || pageNum > totalPages) return false;
            const el = document.getElementById('page-' + pageNum);
            if (el) {
                el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                currentPage = pageNum;
                // Immediately render that area
                setTimeout(renderVisible, 50);
                return true;
            }
            return false;
        }

        function getCurrentPage() { return currentPage; }
        function getTotalPages() { return totalPages; }
        function isReady() { return true; }

        window.pdfViewer = {
            loadPdf, goToPage, getCurrentPage, getTotalPages, isReady,
            getSelectedAnnotation, getPendingDelete
        };
    </script>
</body>
</html>'''

    def _start_page_polling(self):
        """Poll for current page changes."""
        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self._poll_current_page)
        self.poll_timer.start(500)  # Poll every 500ms

    def _poll_current_page(self):
        """Query current page and pending annotation deletes from pdf.js."""
        if not self._viewer_ready:
            return
        self.web_view.page().runJavaScript(
            "window.pdfViewer ? window.pdfViewer.getCurrentPage() : 0",
            self._on_page_result
        )
        # Check for pending annotation deletion
        self.web_view.page().runJavaScript(
            "window.pdfViewer ? window.pdfViewer.getPendingDelete() : null",
            self._on_pending_delete
        )

    def _on_pending_delete(self, ex_id):
        """Handle pending annotation deletion from JS."""
        if ex_id:
            self.annotationDeleteRequested.emit(str(ex_id))

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

    def clear(self):
        """Reset the viewer to a blank state."""
        self.current_pdf_path = None
        self.current_page = 1
        self.total_pages = 0
        self._pending_pdf = None
        self.page_label.setText("Page: - / -")
        self.page_spin.setMaximum(1)
        self.page_spin.setValue(1)
        if self._viewer_ready:
            self.web_view.page().runJavaScript("window.pdfViewer && window.pdfViewer.close && window.pdfViewer.close()")

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
