# CLAUDE.md - iCharlotte Project Context

This file provides context for Claude to understand the iCharlotte suite.

## Project Overview
iCharlotte is a legal document management and automation suite. It consists of a Python-based backend/UI (PyQt6) and a specialized Electron-based React application called **NoteTaker**.

**Technologies:**
- **Backend/Main UI:** Python (PyQt6), QWebEngineView
- **NoteTaker App:** Electron, React, TypeScript, TipTap, react-pdf-highlighter
- **Scripts:** Python for docket scraping, OCR, and document analysis
- **Communication:** QWebChannel bridge for JS/Python interoperability

## Development Conventions
- **UI:** Mimic existing PyQt6 and React patterns.
- **PDF Handling:** Uses a custom `local-resource://` scheme to load local files.
- **Async:** Critical operations (like email) must run in background threads (QThread) to prevent UI freezes.
- **Imports:** Use `qrc:///qtwebchannel/qwebchannel.js` for QWebChannel in the frontend.

## Key Files
- `iCharlotte.py`: Main entry point for the Python application.
- `NoteTaker/`: Source for the Electron/React PDF viewing and highlighting tool.
- `icharlotte_core/`: Core logic for database, email, and LLM integration.
- `Scripts/`: Collection of scrapers and document processing tools.

## Build and Run
- **Python App:** `python iCharlotte.py`
- **NoteTaker (Dev):** `cd NoteTaker && npm run dev`
- **NoteTaker (Build):** `cd NoteTaker && npm run build`

## Screenshots
When taking a screenshot of the iCharlotte app, always use Monitor 0:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\geminiterminal2\screenshot_util.ps1" -Monitor 0
```
