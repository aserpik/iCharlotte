"""Build iCharlotte executable with PyInstaller."""
import subprocess
import sys

# PyInstaller command
cmd = [
    sys.executable, '-m', 'PyInstaller',
    '--onefile',
    '--windowed',
    '--name=iCharlotte',
    '--icon=icharlotte.ico',
    '--add-data=icharlotte.ico;.',
    '--add-data=icharlotte_core;icharlotte_core',
    '--add-data=Scripts;Scripts',
    '--add-data=config;config',
    '--hidden-import=PySide6.QtWebEngineWidgets',
    '--hidden-import=PySide6.QtWebEngineCore',
    '--hidden-import=PySide6.QtWebChannel',
    '--collect-all=PySide6',
    'iCharlotte.py'
]

print("Building iCharlotte.exe...")
print(" ".join(cmd))
subprocess.run(cmd)
