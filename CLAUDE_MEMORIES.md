# CLAUDE_MEMORIES.md - User Preferences & History

This file contains persistent memories and specific logic fixes from the Gemini environment.

## Visual & Formatting Preferences
- **Master Case Status:** All text must be formatted in **Size 8 Times New Roman**.
- **Visual Separator:** The user prefers responses to start with: `==============================================================================`

## Application Logic (iCharlotte / NoteTaker)
- **Separate Function:** Now implemented as a checkbox column (index 3) in the Case View tree.
- **Drag & Drop:** The Chat tab supports .pdf, .docx, and .txt across the entire tab.
- **PDF Loading:** 
  - `loadPdfExternal` is aliased to `setPdfExternal` in `App.tsx`.
  - Uses `local-resource://` URL scheme to bypass bridge hangs.
  - Path normalization is handled by the Python backend.
- **Bridge Stability:** Uses transport polling in the JS shim and a polling fallback in `PDFViewer.tsx` to ensure API availability.

## Known Bug Fixes
- **UI Freezes:** Synchronous Outlook email sending moved to `EmailWorker` (QThread).
- **NoteTaker Errors:** Added defensive checks for `normalizedPath` and `filteredSegments` to prevent "Cannot read length of undefined".
- **Zoom Fix:** PDF zoom is controlled via `ref.current.viewer.currentScaleValue` instead of `pdfScaleValue`.

## Case Folder Mapping
- **1280.001:** `Z:\Shared\Current Clients\1000 - PLAINTIFF\1280.001 - Natalia Finkovskaya`
- **1011.011:** `Z:\Shared\Current Clients\1000 - PLAINTIFF\1011 - Brenna Cox Johnson\1011.011 Iwanaga\JANE DOE v. RALPH DILLON`
- **4519.003:** `Z:\Shared\Current Clients\4500 - JBW DEFENSE\4519 - L1 Technologies\003 - Bushnell`
- **4537.001:** `Z:\Shared\Current Clients\4500 - JBW DEFENSE\4537.001 - Nixon`
- **4553.001:** `Z:\Shared\Current Clients\4500 - JBW DEFENSE\4553.001 - Mark Egerman`
- **4558.001:** `Z:\Shared\Current Clients\4500 - JBW DEFENSE\4558.001 - Pinscreen`
- **4573.001:** `Z:\Shared\Current Clients\4500 - JBW DEFENSE\4573.001 - Gutierrez`
