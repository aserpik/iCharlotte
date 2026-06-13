@echo off
cd /d "%~dp0"
python -m webcompanion.server %*
pause
