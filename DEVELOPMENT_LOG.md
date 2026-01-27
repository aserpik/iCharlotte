# iCharlotte Development Log

This file tracks features, changes, and architectural decisions made during development. Update this file when significant features are added or modified.

---

## How to Use This Log

When adding new features or making significant changes:
1. Add a new entry under the current date section
2. Include: what was added, which files were modified, and any important decisions
3. Keep entries concise but informative

---

## 2026-01-27 - Documentation Update

### Added
- Comprehensive `CLAUDE.md` documentation covering all modules
- Created `DEVELOPMENT_LOG.md` for tracking ongoing changes

---

## Pre-2026-01-27 - Existing Features (Baseline)

### Core Infrastructure

#### LLM Integration (`icharlotte_core/llm.py`, `llm_config.py`)
- Multi-provider support: Gemini, Claude, OpenAI
- Automatic fallback through model sequences
- Streaming response support
- Per-agent and per-task configuration
- 28 configured agents with custom model preferences
- Token counting and cost tracking

#### Exception Handling (`icharlotte_core/exceptions.py`)
- Unified exception hierarchy (ICharlotteException base)
- LLM-specific exceptions: LLMRateLimitError, LLMQuotaExceededError, LLMTimeoutError
- Document exceptions: ExtractionError, OCRError, UnsupportedFormatError
- Retry decorators: @retry_on_error(), @retry_on_llm_error(), @retry_on_file_lock()

#### Document Processing (`icharlotte_core/document_processor.py`)
- Native PDF text extraction with pypdf
- Adaptive OCR with Tesseract (per-page decision)
- Character density analysis for OCR thresholds
- Memory-aware processing with garbage collection
- Support for PDF, DOCX formats

#### Memory Monitoring (`icharlotte_core/memory_monitor.py`)
- Real-time memory usage tracking
- Configurable warning and abort thresholds
- Operation statistics tracking
- Context manager support for monitoring operations

### Chat System (`icharlotte_core/chat/`)

#### Models (`chat/models.py`)
- `Conversation`: Thread with messages, system prompt, model tracking
- `Message`: Individual message with role, content, attachments, timestamps
- `Attachment`: File attachments with token estimation
- `QuickPrompt`: Built-in templates (Summary, Liability, Timeline, Deposition, etc.)

#### Persistence (`chat/persistence.py`)
- JSON-based storage per case (`{case_number}_chat.json`)
- Conversation CRUD operations
- Message management (add, edit, delete, pin)
- Search across conversations
- Import/export functionality
- Settings management (theme, default provider/model)

#### Token Counter (`chat/token_counter.py`)
- Estimation for different model families
- Streaming token counting support

### UI Components (`icharlotte_core/ui/`)

#### Enhanced Case View (`ui/case_view_enhanced.py`)
- `ProcessingLogDB`: Persistent log of file processing operations
- `FileTagsDB`: Tag management for case files
- `AgentSettingsDB`: Per-agent configuration storage
- `EnhancedAgentButton`: Agent runner with status indicators
- `EnhancedFileTreeWidget`: File tree with metadata columns (status, tags, pages)
- `FilePreviewWidget`: Quick file preview pane
- `OutputBrowserWidget`: Browse and compare agent outputs
- `AdvancedFilterWidget`: Complex file filtering
- `ProcessingLogWidget`: Processing history display

#### Chat Widgets (`ui/chat_widgets.py`)
- `ConversationSidebar`: Conversation list with search and filter
- `MessageWidget`: Message display with code syntax highlighting (Pygments)
- `ResizableInputArea`: Multi-line input with auto-expand
- `ContextIndicator`: Token count and context status
- `SearchResultsWidget`: Search results display
- Theme support (Light/Dark)

#### Chat Dialogs (`ui/chat_dialogs.py`)
- System prompt editor dialog
- Model selection dialog
- Quick prompt builder

#### PDF Viewer (`ui/pdf_viewer_widget.py`)
- Embedded pdf.js viewer
- Page tracking and navigation
- Go-to-page control

#### Master Case Tab (`ui/master_case_tab.py`)
- Case list with pagination
- Hearing date tracking with calendar widget
- Todo management with color coding
- Assignment tracking
- Docket date monitoring
- Due date sorting

#### Main Tabs (`ui/tabs.py`)
- `ChatTab`: Full chat interface with streaming, attachments, conversations
- `IndexTab`: File browser and agent runner

#### Common Widgets (`ui/widgets.py`)
- `StatusWidget`: Operation status display
- `AgentRunner`: Background agent execution with QThread
- `FileTreeWidget`: Directory tree navigation
- `OutputBrowserWidget`: Output viewing and comparison

#### Dialogs (`ui/dialogs.py`)
- `FileNumberDialog`: Case file number entry
- `VariablesDialog`: Global variables (firm info, addresses)
- `PromptsDialog`: LLM prompt viewing/editing
- `LLMSettingsDialog`: LLM parameter configuration
- `SettingsDialog`: General app settings

### Database Layer

#### Master Case Database (`icharlotte_core/master_db.py`)
Tables:
- `cases`: file_number, plaintiff, hearing dates, case path, attorney, summary
- `todos`: items, status, due dates, assignments, colors
- `history`: interim updates, status reports
- `processed_emails`: tracking for email-to-todo deduplication

#### Email Database (`icharlotte_core/email_manager.py`)
- Full-text search with SQLite FTS5
- Outlook sync integration
- Email metadata storage (subject, sender, recipients, body, attachments)
- Threading and conversation tracking

### Background Workers

#### Sent Items Monitor (`icharlotte_core/sent_items_monitor.py`)
- Polls Outlook Sent Items every 30 seconds
- Detects emails with configured prefixes (AS -, FM -, HP -, CE -)
- Extracts file numbers from subject
- Auto-creates todos in master database
- Duplicate prevention via processed_emails table

#### Calendar Monitor (`icharlotte_core/calendar/calendar_monitor.py`)
- Monitors Outlook for calendar-related emails
- Classifies attachments (correspondence, discovery, motion, etc.)
- Calculates legal deadlines
- Creates Google Calendar events automatically

### Scripts Collection (44 agents in `/Scripts/`)

Document Analysis:
- `summarize.py` - General document summarization
- `summarize_discovery.py` - Discovery response extraction
- `summarize_deposition.py` - Deposition transcript analysis
- `detect_contradictions.py` - Cross-document contradiction detection
- `med_record.py`, `med_chron.py` - Medical records processing
- `extract_timeline.py` - Timeline extraction
- `ocr.py` - OCR processing

Case Management:
- `complaint.py` - Complaint analysis
- `discovery_requests.py` - Discovery generation
- `docket.py` - Docket processing
- `liability.py` - Liability analysis
- `exposure.py` - Exposure calculations
- `subpoena_tracker.py` - Subpoena tracking

Docket Scrapers (county-specific):
- Orange, Riverside, San Bernardino, San Diego, Sacramento, Kern, Alameda, Mariposa, Santa Clara

Utilities:
- `case_data_manager.py` - Case data CRUD
- `tagging_engine.py` - Document categorization
- `organizer.py` - File organization
- `separate.py` - PDF splitting
- Various cleanup scripts

### Testing (`/tests/`)

Test files:
- `test_agent_logger.py`
- `test_agents_integration.py`
- `test_case_loading.py`
- `test_case_view_db.py`
- `test_case_view_enhanced.py`
- `test_document_processor.py`
- `test_output_validator.py`

Chat tests (`tests/test_chat/`):
- `test_persistence.py`
- `test_streaming.py`
- `test_token_counter.py`

---

## Template for New Entries

```markdown
## YYYY-MM-DD - Brief Description

### Added
- Feature name (`file_path.py`)
  - Key functionality
  - Important classes/functions

### Modified
- What changed and why

### Fixed
- Bug fixes

### Technical Decisions
- Any architectural choices made and rationale

### Files Changed
- List of modified files
```
