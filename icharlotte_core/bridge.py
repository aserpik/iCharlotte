import os
import urllib.parse
from PyQt6.QtCore import QObject, pyqtSlot, QIODevice, QUrl, QThread, pyqtSignal, QFileInfo
from PyQt6.QtQml import QJSValue
from PyQt6.QtWebEngineCore import QWebEngineUrlSchemeHandler, QWebEngineUrlRequestJob, QWebEngineUrlScheme
from PyQt6.QtWidgets import QApplication, QFileDialog

from .utils import log_event
from .ui.logs_tab import LogManager
from .config import SCRIPTS_DIR

import sys
import subprocess

class OCRWorker(QThread):
    finished = pyqtSignal(bool, str) # success, message

    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path

    def run(self):
        try:
            script_path = os.path.join(SCRIPTS_DIR, "ocr.py")
            log_event(f"OCRWorker: Starting OCR on {self.pdf_path}")
            result = subprocess.run(
                [sys.executable, script_path, self.pdf_path],
                capture_output=True,
                text=True,
                check=False
            )
            if result.returncode == 0:
                log_event(f"OCRWorker: OCR Successful for {self.pdf_path}")
                self.finished.emit(True, "OCR Completed")
            else:
                log_event(f"OCRWorker: OCR Failed for {self.pdf_path}: {result.stderr}", "error")
                self.finished.emit(False, f"OCR Failed: {result.stderr}")
        except Exception as e:
            log_event(f"OCRWorker: Error running OCR: {e}", "error")
            self.finished.emit(False, str(e))

class SummarizeWorker(QThread):
    finished = pyqtSignal(str) # result

    def __init__(self, text):
        super().__init__()
        self.text = text

    def run(self):
        from .llm import LLMHandler
        try:
            # We use Gemini 1.5 Flash for quick summaries
            system_prompt = "You are a helpful legal assistant. Summarize the provided text from a legal or medical document concisely but accurately. Keep the summary focused on key facts and dates."
            result = LLMHandler.generate(
                provider="Gemini",
                model="gemini-1.5-flash",
                system_prompt=system_prompt,
                user_prompt=f"Please summarize this selection:\n\n{self.text}",
                file_contents="",
                settings={"temperature": 1.0}
            )
            self.finished.emit(result)
        except Exception as e:
            log_event(f"SummarizeWorker Error: {e}", "error")
            self.finished.emit(f"Summarization error: {str(e)}")

# --- Custom Scheme Handler ---
            
class LocalFileSchemeHandler(QWebEngineUrlSchemeHandler):
    def requestStarted(self, job: QWebEngineUrlRequestJob):
        url = job.requestUrl()
        # Use url.path() and ensure we don't include query params for file system lookups
        path = url.path()
        path = urllib.parse.unquote(path)
        
        # Strip query parameters if they somehow ended up in the path string
        if '?' in path:
            path = path.split('?')[0]
        
        # Clean up path for Windows
        # If URL is local-resource:///C:/path, url.path() is /C:/path
        # If URL is local-resource://C:/path, url.path() is /path and host is C:
        
        if os.name == 'nt':
            if path.startswith('/') and len(path) > 2 and path[2] == ':':
                path = path[1:] # /C:/path -> C:/path
            elif not (len(path) > 1 and path[1] == ':'):
                # Could be UNC path or host-based path
                host = url.host()
                if host:
                    if len(host) == 2 and host[1] == ':':
                        path = host + path
                    else:
                        # UNC path: //host/path
                        path = "//" + host + path
        
        path = os.path.normpath(path)
        
        msg = f"SchemeHandler: Requesting {path}"
        log_event(msg)
        LogManager().add_log("Note Taker", msg)
        
        if not os.path.exists(path):
            msg = f"SchemeHandler: File not found: {path}"
            log_event(msg, "error")
            LogManager().add_log("Note Taker", msg)
            job.fail(QWebEngineUrlRequestJob.Error.UrlNotFound)
            return
            
        try:
            # Use QFile for better integration with QWebEngine streaming
            from PyQt6.QtCore import QFile
            file = QFile(path)
            if not file.open(QIODevice.OpenModeFlag.ReadOnly):
                msg = f"SchemeHandler: Failed to open file: {path}"
                log_event(msg, "error")
                LogManager().add_log("Note Taker", msg)
                job.fail(QWebEngineUrlRequestJob.Error.RequestFailed)
                return
                
            # Keep file alive by parenting it to the job
            file.setParent(job)
            
            size = file.size()
            msg = f"SchemeHandler: Serving {path} ({size} bytes)"
            log_event(msg)
            LogManager().add_log("Note Taker", msg)
            
            job.reply(b"application/pdf", file)
        except Exception as e:
            msg = f"SchemeHandler Error: {e}"
            log_event(msg, "error")
            LogManager().add_log("Note Taker", msg)
            job.fail(QWebEngineUrlRequestJob.Error.RequestFailed)

class NoteTakerBridge(QObject):
    ocrFinished = pyqtSignal(str, bool, str) # path, success, message

    def __init__(self, parent=None):
        super().__init__(parent)
        self._parent = parent
        self.ocr_workers = {}
        self.summarize_workers = []
        log_event("NoteTakerBridge initialized")

    @pyqtSlot(str, result=bool)
    def checkNeedsOCR(self, path):
        import fitz
        doc = None
        try:
            doc = fitz.open(path)
            has_text = False
            pages_to_check = min(5, len(doc))
            for i in range(pages_to_check):
                page = doc[i]
                if page.get_text().strip():
                    has_text = True
                    break
            return not has_text
        except Exception as e:
            log_event(f"Error in checkNeedsOCR: {e}", "error")
            return False
        finally:
            if doc:
                doc.close()

    @pyqtSlot(str, result=bool)
    def runOCR(self, path):
        if path in self.ocr_workers:
            return True
            
        worker = OCRWorker(path)
        self.ocr_workers[path] = worker
        
        worker.finished.connect(lambda s, m: self.on_ocr_finished(path, s, m))
        worker.start()
        return True

    def on_ocr_finished(self, path, success, message):
        if path in self.ocr_workers:
            del self.ocr_workers[path]
        log_event(f"OCR Finished for {path}: {success} - {message}")
        self.ocrFinished.emit(path, success, message)

    @pyqtSlot(str, 'QJSValue')
    def summarizeText(self, text, callback):
        worker = SummarizeWorker(text)
        self.summarize_workers.append(worker)
        
        def on_done(result):
            try:
                callback.call([result])
            except Exception as e:
                log_event(f"Error calling back summarizeText: {e}", "error")
            if worker in self.summarize_workers:
                self.summarize_workers.remove(worker)
                
        worker.finished.connect(on_done)
        worker.start()

    @pyqtSlot(str, result=bool)
    def saveMarkdown(self, content, defaultName):
        path, _ = QFileDialog.getSaveFileName(self._parent, "Save Markdown", defaultName, "Markdown Files (*.md)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                return True
            except Exception as e:
                log_event(f"Error saving markdown: {e}", "error")
                return False
        return False

    @pyqtSlot(result=str)
    def selectDirectory(self):
        try:
            # Use active window as parent for better dialog behavior
            parent = QApplication.activeWindow() or self._parent
            dir_path = QFileDialog.getExistingDirectory(parent, "Select Directory")
            log_event(f"Selected directory: {dir_path}")
            return dir_path if dir_path else None
        except Exception as e:
            log_event(f"Error in selectDirectory: {e}", "error")
            return None

    @pyqtSlot(str, result=list)
    def listDirectory(self, dir_path):
        try:
            if not dir_path or not os.path.exists(dir_path):
                return []
            files = []
            for item in os.listdir(dir_path):
                full_path = os.path.join(dir_path, item)
                try:
                    stats = os.stat(full_path)
                    files.append({
                        "name": item,
                        "isDirectory": os.path.isdir(full_path),
                        "path": full_path,
                        "size": stats.st_size,
                        "mtime": stats.st_mtime * 1000 # ms
                    })
                except:
                    continue
            return files
        except Exception as e:
            log_event(f"Error listing directory {dir_path}: {e}", "error")
            return []

    @pyqtSlot(str, str, result=bool)
    def moveFile(self, old_path, new_path):
        import shutil
        try:
            shutil.move(old_path, new_path)
            return True
        except Exception as e:
            log_event(f"Error moving file: {e}", "error")
            return False

    @pyqtSlot(str, result=str)
    def readFile(self, path):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception as e:
            log_event(f"Error reading file: {e}", "error")
            return None

    @pyqtSlot(str, result=str)
    def readBuffer(self, path):
        """Returns file content as base64 string because QWebChannel handles strings better than raw bytes."""
        import base64
        import os
        try:
            # Handle potential URL encoding or prefixes if passed from JS
            if path.startswith('file:///'):
                path = path[8:]
            elif path.startswith('file://'):
                path = path[7:]
            
            # Decode URL-encoded characters (like %20 for spaces)
            path = urllib.parse.unquote(path)
            
            # Normalize path for Windows
            norm_path = os.path.normpath(path)
            
            msg = f"Bridge reading buffer from: {norm_path} (Original: {path})"
            log_event(msg)
            LogManager().add_log("Note Taker", msg)
            
            if not os.path.exists(norm_path):
                msg = f"File does not exist: {norm_path}"
                log_event(msg, "error")
                LogManager().add_log("Note Taker", msg)
                return None
                
            with open(norm_path, "rb") as f:
                data = f.read()
                b64_data = base64.b64encode(data).decode("utf-8")
                msg = f"Successfully read {len(data)} bytes from {norm_path}, base64 size: {len(b64_data)}"
                log_event(msg)
                LogManager().add_log("Note Taker", msg)
                return b64_data
        except Exception as e:
            msg = f"Error reading buffer from {path}: {e}"
            log_event(msg, "error")
            LogManager().add_log("Note Taker", msg)
            return None

    @pyqtSlot(str, str, result=bool)
    def writeFile(self, path, content):
        import os
        try:
            norm_path = os.path.normpath(path)
            with open(norm_path, "w", encoding="utf-8") as f:
                f.write(content)
            return True
        except Exception as e:
            log_event(f"Error writing file {path}: {e}", "error")
            return False

    @pyqtSlot(str, result=bool)
    def saveAppState(self, state_json):
        from .config import GEMINI_DATA_DIR
        try:
            config_dir = os.path.join(GEMINI_DATA_DIR, "..", "config")
            if not os.path.exists(config_dir):
                os.makedirs(config_dir)
            state_path = os.path.join(config_dir, "notetaker_state.json")
            with open(state_path, "w", encoding="utf-8") as f:
                f.write(state_json)
            return True
        except Exception as e:
            log_event(f"Error saving app state: {e}", "error")
            return False

    @pyqtSlot(result=str)
    def loadAppState(self):
        from .config import GEMINI_DATA_DIR
        try:
            state_path = os.path.join(GEMINI_DATA_DIR, "..", "config", "notetaker_state.json")
            if not os.path.exists(state_path):
                return None
            with open(state_path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception as e:
            log_event(f"Error loading app state: {e}", "error")
            return None

def register_custom_schemes():
    from PyQt6.QtWebEngineCore import QWebEngineUrlScheme
    scheme = QWebEngineUrlScheme(b"local-resource")
    scheme.setFlags(
        QWebEngineUrlScheme.Flag.CorsEnabled |
        QWebEngineUrlScheme.Flag.LocalAccessAllowed |
        QWebEngineUrlScheme.Flag.SecureScheme
    )
    QWebEngineUrlScheme.registerScheme(scheme)
    log_event("Custom schemes registered")
