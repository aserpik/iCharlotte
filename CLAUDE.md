# CLAUDE.md - iCharlotte Project Context

This file provides context for Claude to understand the iCharlotte suite.

## Project Overview

iCharlotte is a legal document management and automation suite for law firm case management. It provides intelligent document processing, multi-provider LLM integration, and automated workflows for legal professionals.

**Technologies:**
- **Backend/Main UI:** Python 3.x, PyQt6/PySide6, SQLite
- **NoteTaker App:** Electron, React, TypeScript, TipTap, react-pdf-highlighter
- **Document Processing:** pypdf, python-docx, Tesseract OCR, PyMuPDF
- **LLM Providers:** Google Gemini, Anthropic Claude, OpenAI (with automatic fallback)
- **Communication:** QWebChannel bridge for JS/Python interoperability
- **External Integration:** Outlook (win32com), Google Calendar API

---

## Directory Structure

```
iCharlotte/
├── iCharlotte.py              # Main application entry point
├── icharlotte_core/           # Core modules
│   ├── llm.py                 # LLM handler (streaming, caching, multi-provider)
│   ├── llm_config.py          # Centralized LLM configuration with fallback
│   ├── master_db.py           # Case database (cases, todos, history)
│   ├── email_manager.py       # Email database and Outlook integration
│   ├── document_processor.py  # Text extraction, OCR, document processing
│   ├── exceptions.py          # Unified exception hierarchy with retry decorators
│   ├── memory_monitor.py      # Memory usage tracking and limits
│   ├── output_validator.py    # Output validation for agent results
│   ├── config.py              # Paths, API keys, configuration
│   ├── utils.py               # Shared utilities
│   ├── bridge.py              # Custom URL scheme handler (local-resource://)
│   ├── sent_items_monitor.py  # Outlook Sent Items monitoring for todos
│   ├── chat/                  # Chat system
│   │   ├── models.py          # Conversation, Message, Attachment models
│   │   ├── persistence.py     # JSON-based chat history storage
│   │   └── token_counter.py   # Token usage estimation
│   ├── calendar/              # Calendar integration
│   │   ├── calendar_monitor.py    # Outlook monitoring for calendar events
│   │   ├── deadline_calculator.py # Legal deadline calculations
│   │   ├── attachment_classifier.py
│   │   └── gcal_client.py     # Google Calendar API client
│   └── ui/                    # PyQt6 UI components
│       ├── tabs.py            # ChatTab, IndexTab
│       ├── case_view_enhanced.py  # Advanced case view with logs, tags, filters
│       ├── chat_widgets.py    # Conversation sidebar, message widgets
│       ├── chat_dialogs.py    # Chat settings dialogs
│       ├── pdf_viewer_widget.py   # PDF viewer using pdf.js
│       ├── master_case_tab.py # Case list with hearings and todos
│       ├── email_tab.py       # Email search and management
│       ├── liability_tab.py   # Liability analysis
│       ├── dialogs.py         # Settings dialogs
│       └── widgets.py         # Common widgets (status, file tree, agent runner)
├── Scripts/                   # 44 Python analysis agents
├── NoteTaker/                 # Electron/React PDF viewer app
├── tests/                     # Test suite
└── config/                    # Configuration files
    └── llm_preferences.json   # LLM model preferences
```

---

## Core Modules

### LLM Integration (`llm.py`, `llm_config.py`)

Multi-provider LLM support with automatic fallback:

- **LLMHandler**: Static methods for content generation with streaming
- **LLMConfig** (Singleton): Centralized configuration from `config/llm_preferences.json`
- **LLMCaller**: Unified caller with automatic model fallback through sequences
- **ModelSpec**: Dataclass for model definition (provider, tokens, capabilities)
- **AgentConfig**: Per-agent LLM configuration with model sequences
- **TaskConfig**: Task-type configs (general, extraction, summary, cross_check)

**28 configured agents** including: Summarize, Medical Records, Docket, Discovery, Chat, Email Intelligence, etc.

### Chat System (`icharlotte_core/chat/`)

Persistent conversation-based chat:

- **Conversation**: Thread with messages, system prompt, model tracking
- **Message**: Individual message with metadata, attachments, edit history
- **QuickPrompt**: Template prompts (Summary, Liability, Timeline, Deposition, etc.)
- **ChatPersistence**: JSON storage per case (`{case_number}_chat.json`)
- **TokenCounter**: Estimation for Claude, GPT, Gemini models

### Database Layer

**MasterCaseDatabase** (`master_db.py`):
- Tables: `cases`, `todos`, `history`, `processed_emails`
- Methods: `upsert_case()`, `create_todo()`, `get_todos()`

**EmailDatabase** (`email_manager.py`):
- Full-text search with FTS5
- Outlook sync integration

### Document Processing (`document_processor.py`)

- Native PDF extraction (pypdf)
- Adaptive OCR with Tesseract
- Per-page OCR decision based on character density
- Memory-aware processing with garbage collection

### Exception Handling (`exceptions.py`)

Unified hierarchy with retry decorators:
- `@retry_on_error()`: Generic retry with exponential backoff
- `@retry_on_llm_error()`: LLM-specific with model fallback
- Exception types: LLMError, ExtractionError, ValidationError, MemoryLimitError

---

## UI Components

### Main Tabs
1. **ChatTab**: AI chat with conversations, attachments, streaming
2. **IndexTab**: File browser and agent runner
3. **Master Case Tab**: Case list with hearings, todos, assignments
4. **Email Tab**: Search and preview emails
5. **Liability Tab**: Liability analysis
6. **Logs Tab**: Activity logging

### Key Widgets
- **EnhancedFileTreeWidget**: File tree with processing status, tags, filters
- **ConversationSidebar**: Chat conversation management
- **MessageWidget**: Message display with code highlighting (Pygments)
- **PDFViewerWidget**: Embedded pdf.js viewer
- **AgentRunner**: Background agent execution with progress

---

## Scripts Collection (44 agents)

**Document Analysis:**
- `summarize.py`, `summarize_discovery.py`, `summarize_deposition.py`
- `detect_contradictions.py` - Cross-document contradiction detection
- `med_record.py`, `med_chron.py` - Medical records processing
- `extract_timeline.py`, `ocr.py`

**Case Management:**
- `complaint.py`, `discovery_requests.py`, `docket.py`
- `liability.py`, `exposure.py`, `subpoena_tracker.py`

**Docket Scrapers** (county-specific):
- Orange, Riverside, San Bernardino, San Diego, Sacramento, Kern, Alameda, Mariposa, Santa Clara

---

## Development Conventions

### Code Style
- **UI**: Follow existing PyQt6 patterns
- **Async**: Critical operations in QThread to prevent UI freezes
- **Error Handling**: Use exception hierarchy from `exceptions.py`
- **LLM Calls**: Always use `LLMCaller` for automatic fallback

### PDF Handling
- Custom `local-resource://` scheme for local files
- Use `bridge.py` LocalFileSchemeHandler

### Database
- SQLite with context managers
- Upsert pattern for case data

### Testing
- Use unittest framework
- Tests in `tests/` directory
- Chat tests in `tests/test_chat/`

---

## Build and Run

```bash
# Python App
python iCharlotte.py

# NoteTaker (Development)
cd NoteTaker && npm run dev

# NoteTaker (Build)
cd NoteTaker && npm run build

# Run Tests
python -m pytest tests/
```

---

## Screenshots

When taking a screenshot of the iCharlotte app:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\geminiterminal2\screenshot_util.ps1" -WindowTitle "iCharlotte"
```

---

## Configuration Files

- `config/llm_preferences.json` - LLM model sequences and agent configs
- `.env` - API keys (GEMINI_API_KEY, ANTHROPIC_API_KEY, OPENAI_API_KEY)
- Activity log at configured LOG_FILE path

---

## Recent Features

See `DEVELOPMENT_LOG.md` for detailed changelog of features added during development.
