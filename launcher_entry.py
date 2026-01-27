"""Minimal launcher entry point for PyInstaller."""
import sys
import os
import ctypes

# Set AppUserModelID first
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('iCharlotte.LegalSuite.1')

# Set up paths
if getattr(sys, 'frozen', False):
    # Running as compiled exe
    app_dir = os.path.dirname(sys.executable)
else:
    app_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(app_dir)
sys.path.insert(0, app_dir)

# Run the main application
exec(compile(open(os.path.join(app_dir, 'iCharlotte.py')).read(), 'iCharlotte.py', 'exec'))
