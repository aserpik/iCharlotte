# Word Redline Mode

## Overview

The Word AI Assistant now supports **Redline Mode**, which inserts AI suggestions as Track Changes instead of replacing text completely. This provides surgical editing capability crucial for legal document review.

## Usage

1. Select text in Word document
2. Press **Win+V** to open AI Assistant
3. Check **✏️ Use Redline Mode (Track Changes)**
4. Enter your prompt (e.g., "Make this more aggressive")
5. Click **Process**
6. AI suggestions appear as Track Changes in Word
7. Accept or reject changes using Word's review tools

## Configuration

Settings are stored in `~/.gemini/redline_settings.json`:

```json
{
  "redline_mode_default": false,
  "auto_enable_track_changes": true,
  "redline_fallback_notify": true,
  "max_redline_text_length": 50000
}
```

- `redline_mode_default`: Checkbox starts checked if true
- `auto_enable_track_changes`: Auto-enable Track Changes if document has it disabled
- `redline_fallback_notify`: Show notification when falling back to replace mode
- `max_redline_text_length`: Maximum characters for redline mode

## Technical Details

Uses [adeu](https://github.com/dealfluence/adeu) RedlineEngine to inject native Word Track Changes XML (`w:ins`, `w:del`) into documents while preserving formatting.

## Limitations

- Only available for Word (not Outlook emails)
- Requires text selection (cannot redline empty document)
- Very large selections (>50,000 chars) may be slow
- Complex document structures may fall back to replace mode
