
import sys
import os

# Mock PyQt6
try:
    from PyQt6.QtCore import QObject, pyqtSlot, QIODevice, QUrl
    from PyQt6.QtWebEngineCore import QWebEngineUrlSchemeHandler, QWebEngineUrlRequestJob, QWebEngineUrlScheme
    from PyQt6.QtWidgets import QApplication, QFileDialog
except ImportError:
    print("PyQt6 not installed, skipping import test")
    sys.exit(0)

try:
    from icharlotte_core.bridge import NoteTakerBridge, register_custom_schemes
    print("Bridge imported successfully")
except Exception as e:
    print(f"Bridge import failed: {e}")
    sys.exit(1)

try:
    bridge = NoteTakerBridge()
    print("Bridge instantiated successfully")
except Exception as e:
    print(f"Bridge instantiation failed: {e}")
    sys.exit(1)
