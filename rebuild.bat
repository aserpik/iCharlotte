@echo off
echo Rebuilding iCharlotte.exe...
cd /d C:\geminiterminal2
python -m PyInstaller --onedir --windowed --name=iCharlotte --icon=icharlotte.ico --noconfirm iCharlotte.py
echo.
echo Build complete! exe is at: dist\iCharlotte\iCharlotte.exe
pause
