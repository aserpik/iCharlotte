
import sys
import os
import json
import base64
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebChannel import QWebChannel
from PyQt6.QtCore import QObject, pyqtSlot, QUrl, QTimer

# The exact bridge from your codebase
class NoteTakerBridge(QObject):
    @pyqtSlot(str, result=str)
    def readBuffer(self, path):
        print(f"DEBUG: Bridge received readBuffer request for {path}")
        return base64.b64encode(b"test data").decode("utf-8")

def run_test():
    app = QApplication(sys.argv)
    win = QMainWindow()
    win.setWindowTitle("Bridge Diagnostic Tool")
    win.resize(800, 600)

    browser = QWebEngineView()
    channel = QWebChannel()
    bridge = NoteTakerBridge()
    channel.registerObject("python_bridge", bridge)
    browser.page().setWebChannel(channel)

    # Use a minimal HTML that tries to connect to the bridge
    test_html = """
    <html>
    <body>
        <div id="status">Waiting for bridge...</div>
        <pre id="log"></pre>
        <script>
            const log = document.getElementById('log');
            const status = document.getElementById('status');
            function logMsg(m) { 
                log.textContent += new Date().toLocaleTimeString() + ": " + m + "\n"; 
                console.log(m);
            }

            logMsg("Diagnostic started");
            
            // Minimal QWebChannel shim
            (function() {
                logMsg("Polling for qt.webChannelTransport...");
                let count = 0;
                function check() {
                    count++;
                    if (typeof qt !== 'undefined' && qt.webChannelTransport) {
                        logMsg("Transport FOUND after " + count + " tries");
                        // We rely on the real qwebchannel.js logic from your app
                        // but here we just test if the transport even responds
                        status.textContent = "Transport Detected!";
                        status.style.color = "green";
                    } else {
                        if (count < 50) setTimeout(check, 100);
                        else {
                            logMsg("Transport NOT FOUND after 5s");
                            status.textContent = "Transport Missing";
                            status.style.color = "red";
                        }
                    }
                }
                check();
            })();
        </script>
    </body>
    </html>
    """
    browser.setHtml(test_html)
    
    layout = QVBoxLayout()
    layout.addWidget(browser)
    container = QWidget()
    container.setLayout(layout)
    win.setCentralWidget(container)
    
    win.show()
    
    # Auto-close after 10 seconds so I can read the shell results if any
    QTimer.singleShot(10000, app.quit)
    sys.exit(app.exec())

if __name__ == "__main__":
    run_test()
