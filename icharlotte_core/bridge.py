import os
import urllib.parse
from PySide6.QtCore import QIODevice
from PySide6.QtWebEngineCore import QWebEngineUrlSchemeHandler, QWebEngineUrlRequestJob, QWebEngineUrlScheme

from .utils import log_event
from .ui.logs_tab import LogManager


class LocalFileSchemeHandler(QWebEngineUrlSchemeHandler):
    """Custom URL scheme handler for serving local files (e.g., PDFs) in QWebEngineView."""

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
        LogManager().add_log("PDF Viewer", msg)

        if not os.path.exists(path):
            msg = f"SchemeHandler: File not found: {path}"
            log_event(msg, "error")
            LogManager().add_log("PDF Viewer", msg)
            job.fail(QWebEngineUrlRequestJob.Error.UrlNotFound)
            return

        try:
            # Use QFile for better integration with QWebEngine streaming
            from PySide6.QtCore import QFile
            file = QFile(path)
            if not file.open(QIODevice.OpenModeFlag.ReadOnly):
                msg = f"SchemeHandler: Failed to open file: {path}"
                log_event(msg, "error")
                LogManager().add_log("PDF Viewer", msg)
                job.fail(QWebEngineUrlRequestJob.Error.RequestFailed)
                return

            # Keep file alive by parenting it to the job
            file.setParent(job)

            size = file.size()
            msg = f"SchemeHandler: Serving {path} ({size} bytes)"
            log_event(msg)
            LogManager().add_log("PDF Viewer", msg)

            job.reply(b"application/pdf", file)
        except Exception as e:
            msg = f"SchemeHandler Error: {e}"
            log_event(msg, "error")
            LogManager().add_log("PDF Viewer", msg)
            job.fail(QWebEngineUrlRequestJob.Error.RequestFailed)


def register_custom_schemes():
    """Register the local-resource:// scheme for serving local files in QWebEngineView."""
    scheme = QWebEngineUrlScheme(b"local-resource")
    scheme.setFlags(
        QWebEngineUrlScheme.Flag.CorsEnabled |
        QWebEngineUrlScheme.Flag.LocalAccessAllowed |
        QWebEngineUrlScheme.Flag.SecureScheme
    )
    QWebEngineUrlScheme.registerScheme(scheme)
    log_event("Custom schemes registered")
