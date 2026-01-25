import os
import json
from PyQt6.QtWidgets import QWidget, QVBoxLayout
from PyQt6.QtCore import pyqtSignal, QUrl, QObject
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEnginePage, QWebEngineScript
from PyQt6.QtWebChannel import QWebChannel

from ..config import NOTETAKER_DIR
from ..utils import log_event
from ..bridge import NoteTakerBridge, LocalFileSchemeHandler
from .logs_tab import LogManager

class NoteTakerPage(QWebEnginePage):
    consoleMessageSignal = pyqtSignal(int, str, int, str)

    def javaScriptConsoleMessage(self, level, message, line, sourceID):
        try:
            # level is an enum in PyQt6 (JavaScriptConsoleMessageLevel)
            lvl = level.value if hasattr(level, 'value') else int(level)
            self.consoleMessageSignal.emit(lvl, message, line, sourceID)
        except:
            self.consoleMessageSignal.emit(0, message, line, sourceID)

class NoteTakerBrowser(QWebEngineView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPage(NoteTakerPage(self))
        self.setAcceptDrops(True)

        # Setup Bridge attached to the Browser instance
        self.channel = QWebChannel(self.page())
        self.bridge = NoteTakerBridge(self)
        self.channel.registerObject("python_bridge", self.bridge)
        self.page().setWebChannel(self.channel)

        self.inject_shim()

    def inject_shim(self):
        js_shim = """
        "use strict";

        // Define loadPdfExternal globally immediately
        window.loadPdfExternal = function(p) {
            console.log("loadPdfExternal called with:", p);
            if (window.setPdfExternal) {
                window.setPdfExternal(p);
            } else {
                console.warn("setPdfExternal not found yet, retrying in 500ms...");
                setTimeout(() => {
                    if (window.setPdfExternal) window.setPdfExternal(p);
                    else console.error("setPdfExternal still not found after retry.");
                }, 500);
            }
        };

        (function() {
            // Check if QWebChannel already exists
            if (window.QWebChannel) return;

            var script = document.createElement('script');
            script.src = 'qrc:///qtwebchannel/qwebchannel.js';
            script.onload = function() {
                initBridge();
            };
            script.onerror = function() {
                console.error("NoteTaker Shim: Failed to load qwebchannel.js from QRC");
            };
            // Append to head or doc
            (document.head || document.documentElement).appendChild(script);

            function initBridge() {
                console.log("NoteTaker Shim: QWebChannel.js loaded. Initializing...");
                
                // QWebChannel is now available globally
                if (typeof QWebChannel === 'undefined') {
                    console.error("NoteTaker Shim: QWebChannel is undefined despite script load!");
                    return;
                }

                new QWebChannel(qt.webChannelTransport, function(channel) {
                    console.log("NoteTaker Shim: QWebChannel handshake COMPLETE");
                    window.bridge = channel.objects.python_bridge;
                    
                    if (!window.bridge) {
                        console.error("NoteTaker Shim: python_bridge object not found in channel!");
                        return;
                    }
                    
                    // Helper
                    function base64ToArrayBuffer(base64) {
                        if (!base64 || typeof base64 !== 'string') {
                            console.warn("base64ToArrayBuffer: received invalid input:", typeof base64);
                            return null;
                        }
                        try {
                            var binary_string = window.atob(base64);
                            if (!binary_string) {
                                console.warn("base64ToArrayBuffer: atob returned empty/null");
                                return null;
                            }
                            var len = binary_string.length;
                            var bytes = new Uint8Array(len);
                            for (var i = 0; i < len; i++) bytes[i] = binary_string.charCodeAt(i);
                            return bytes.buffer;
                        } catch (e) {
                            console.error("base64ToArrayBuffer error:", e);
                            return null;
                        }
                    }

                    window.bridge.ocrFinished.connect(function(path, success, message) {
                        console.log("NoteTaker Shim: ocrFinished received:", path, success, message);
                        window.dispatchEvent(new CustomEvent('ocr-finished', { detail: { path, success, message } }));
                    });

                    window.api = {
                        saveMarkdown: (c, n) => new Promise(r => {
                            window.bridge.saveMarkdown(c, n, r);
                        }),
                        selectDirectory: () => new Promise(r => {
                            window.bridge.selectDirectory(r);
                        }),
                        listDirectory: (p) => new Promise(r => {
                            window.bridge.listDirectory(p, r);
                        }),
                        readFile: (p) => new Promise(r => {
                            window.bridge.readFile(p, r);
                        }),
                        readBuffer: (p) => new Promise(r => {
                            console.log("api.readBuffer called for:", p);
                            window.bridge.readBuffer(p, res => {
                                try {
                                    console.log("api.readBuffer callback triggered");
                                    if (res === undefined) {
                                        console.error("api.readBuffer: received undefined response");
                                        r(null);
                                    } else if (res === null) {
                                        console.error("api.readBuffer: received null response");
                                        r(null);
                                    } else {
                                        // console.log("api.readBuffer: received response of type", typeof res);
                                        const buffer = base64ToArrayBuffer(res);
                                        r(buffer);
                                    }
                                } catch (err) {
                                    console.error("api.readBuffer exception in callback:", err);
                                    r(null);
                                }
                            });
                        }),
                        writeFile: (p, c) => new Promise(r => {
                            window.bridge.writeFile(p, c, r);
                        }),
                        moveFile: (o, n) => new Promise(r => {
                            window.bridge.moveFile(o, n, r);
                        }),
                        saveAppState: (s) => new Promise(r => {
                            window.bridge.saveAppState(JSON.stringify(s), r);
                        }),
                        loadAppState: () => new Promise(r => {
                            window.bridge.loadAppState(res => {
                                try {
                                    r(res ? JSON.parse(res) : null);
                                } catch (e) {
                                    console.error("Failed to parse app state", e);
                                    r(null);
                                }
                            });
                        }),
                        checkNeedsOCR: (p) => new Promise(r => {
                            window.bridge.checkNeedsOCR(p, r);
                        }),
                        runOCR: (p) => new Promise(r => {
                            window.bridge.runOCR(p, r);
                        }),
                        summarizeText: (t) => new Promise(r => {
                            window.bridge.summarizeText(t, r);
                        })
                    };
                    
                    window.electron = {
                        ipcRenderer: {
                            invoke: (ch, ...args) => {
                                if (window.api[ch]) return window.api[ch](...args);
                                let m = ch.replace(/-([a-z])/g, g => g[1].toUpperCase());
                                if (window.api[m]) return window.api[m](...args);
                            }
                        }
                    };
                    
                    console.log("NoteTaker Shim: Bridge initialized successfully");
                    window.bridge_ready = true;
                    window.dispatchEvent(new CustomEvent('bridge-ready'));
                    if (window.api) window.api.isReady = true;
                });
            }
        })();
        """

        # Inject shim using QWebEngineScript for persistent injection at DocumentReady
        self.script = QWebEngineScript()
        self.script.setSourceCode(js_shim)
        self.script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentReady)
        self.script.setWorldId(QWebEngineScript.ScriptWorldId.MainWorld)
        self.script.setRunsOnSubFrames(True)
        self.page().scripts().insert(self.script)


    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith('.pdf'):
                    event.accept()
                    return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path:
                msg = f"Drop detected: Raw path from URL: {path}"
                log_event(msg)
                LogManager().add_log("Note Taker", msg)

                if path.lower().endswith('.pdf'):
                    # Ensure path is clean for JS
                    clean_path = path.replace('\\', '/')
                    msg = f"Triggering loadPdfExternal with: {clean_path}"
                    log_event(msg)
                    LogManager().add_log("Note Taker", msg)
                    
                    # Use json.dumps to handle quotes and escaping safely
                    json_path = json.dumps(clean_path)
                    js = f"if (window.loadPdfExternal) {{ window.loadPdfExternal({json_path}); }} else {{ console.error('loadPdfExternal not found'); }}"
                    
                    self.page().runJavaScript(js)
                    event.acceptProposedAction()
                    return
            else:
                log_event("NoteTaker Drop: URL has no local file path.", "warning")
                LogManager().add_log("Note Taker", "Drop: URL has no local file path.")
        super().dropEvent(event)

class NoteTakerTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.browser = NoteTakerBrowser()
        # Enable some features
        s = self.browser.settings()
        s.setAttribute(s.WebAttribute.JavascriptEnabled, True)
        s.setAttribute(s.WebAttribute.JavascriptCanAccessClipboard, True)
        s.setAttribute(s.WebAttribute.LocalStorageEnabled, True)
        s.setAttribute(s.WebAttribute.XSSAuditingEnabled, False)
        s.setAttribute(s.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        s.setAttribute(s.WebAttribute.PluginsEnabled, True)
        s.setAttribute(s.WebAttribute.PdfViewerEnabled, True)
        
        # Log console messages
        self.browser.page().consoleMessageSignal.connect(self._on_console_message)

        layout.addWidget(self.browser)

        # Install Scheme Handler
        self.handler = LocalFileSchemeHandler(self)
        self.browser.page().profile().installUrlSchemeHandler(b"local-resource", self.handler)

        # Better qwebchannel.js shim
        dev_url = os.environ.get("NOTE_TAKER_DEV_URL")
        if dev_url:
            log_event(f"NoteTaker: Loading Dev URL: {dev_url}")
            self.browser.setUrl(QUrl(dev_url))
        else:
            index_path = os.path.normpath(os.path.join(NOTETAKER_DIR, "index.html"))
            if os.path.exists(index_path):
                self.browser.setUrl(QUrl.fromLocalFile(index_path))
            else:
                self.browser.setHtml(f"<h1>NoteTaker build files not found</h1><p>Expected path: {index_path}</p>")

    def _on_console_message(self, level, message, line, sourceID):
        msg = f"JS ({level}): {message} [{sourceID}:{line}]"
        log_event(f"NoteTaker {msg}")
        LogManager().add_log("NoteTaker", msg)

    def reset_state(self):
        """Reloads the browser to reset state/context."""
        self.browser.reload()
