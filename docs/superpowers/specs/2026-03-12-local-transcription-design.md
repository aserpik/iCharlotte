# Local Audio/Video Transcription via faster-whisper

**Date:** 2026-03-12
**Status:** Approved

## Problem

The chat tab cannot transcribe audio/video files locally. When users attach media and ask for transcription, non-Gemini models receive only a text placeholder and hallucinate fake transcripts. Gemini can receive the file via upload, but that sends data to Google's servers. There is no local, private transcription path.

## Solution

Add a "Transcribe" quick prompt template that runs faster-whisper locally on CPU, producing a clean transcript as a standalone chat message. No LLM call is made.

## User Flow

1. User attaches one or more audio/video files in the chat tab.
2. User clicks the "Transcribe" quick prompt.
3. System detects checked audio/video files. If none, shows a warning.
4. A background `QThread` runs faster-whisper (`medium` model, CPU) on each file.
5. Progress is shown in the chat area (e.g., "Transcribing 930.mp4...").
6. The transcript appears as a message in the chat.
7. The transcript is saved to conversation persistence.
8. No LLM call is made — this is a standalone action.

## Components

### LocalTranscriberWorker (QThread)

- **Location:** `icharlotte_core/llm.py` (alongside LLMWorker)
- **Inputs:** List of audio/video file paths, model name (default `medium`)
- **Signals:**
  - `progress(str)` — status updates ("Transcribing file 1/3...")
  - `finished(str)` — full transcript text
  - `error(str)` — error message
- **Behavior:**
  - Imports `faster_whisper.WhisperModel` on first use (lazy import)
  - Instantiates model with `device="cpu"`, `compute_type="int8"` for best CPU performance
  - Calls `model.transcribe(path)` for each file
  - Concatenates segment text (no timestamps) into plain transcript
  - Prefixes each file's output with `## Transcript: <filename>` when multiple files

### Quick Prompt Template

- **Name:** "Transcribe"
- **Location:** Added to existing quick prompt list in `ChatTab`
- **Behavior:** When selected, calls `_run_local_transcription()` instead of sending to LLM

### ChatTab._run_local_transcription()

- **Trigger:** "Transcribe" quick prompt selected with audio/video files attached
- **Steps:**
  1. Calls `_get_checked_audio_files()` to get media paths
  2. If no audio/video files, shows warning in chat and returns
  3. Saves a user message to persistence (e.g., "[Transcribe: file1.mp4, file2.mp4]")
  4. Shows "Transcribing..." status in chat
  5. Creates `LocalTranscriberWorker` and starts it
  6. On `finished`: displays transcript, saves to persistence
  7. On `error`: displays error in chat

## Model Configuration

- **Default model:** `medium`
- **Device:** `cpu`
- **Compute type:** `int8` (fastest for CPU)
- **Expected performance:** ~2-4 minutes per 5 minutes of audio on CPU
- Model size constant defined in `LocalTranscriberWorker` — easy to change later

## Existing Behavior Preserved

- **Gemini file upload** remains: manually sending audio/video with Gemini provider still uploads via Files API
- **Non-Gemini warning** remains: attaching audio/video with Claude/OpenAI still shows the hallucination warning dialog
- **Quick prompt "Transcribe"** always uses local faster-whisper, regardless of selected provider

## Dependencies

- `faster-whisper` (pip install) — pulls in CTranslate2, huggingface_hub
- `ffmpeg` — already installed on the system
- No GPU required; runs on CPU with int8 quantization

## File Changes

| File | Change |
|------|--------|
| `icharlotte_core/llm.py` | Add `LocalTranscriberWorker` class |
| `icharlotte_core/ui/tabs.py` | Add "Transcribe" quick prompt, `_run_local_transcription()`, wire signals |
| `requirements.txt` (if exists) | Add `faster-whisper` |

## Out of Scope

- Timestamp output (clean text only)
- Model size selector UI (hardcoded `medium` for now)
- Speaker diarization
- Real-time streaming transcription
