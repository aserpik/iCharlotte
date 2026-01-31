"""
Process-safe DOCX Writer for iCharlotte

Provides file locking to prevent concurrent write conflicts when multiple
agents try to write to the same output file (e.g., AI_OUTPUT.docx).

Uses file-based locking that works across separate Python processes.
"""

import os
import time
import random
import datetime
import logging
from typing import Optional, Callable
from contextlib import contextmanager

try:
    from docx import Document
    from docx.shared import Pt, Inches
except ImportError:
    Document = None
    Pt = None
    Inches = None

# Lock timeout and retry settings
LOCK_TIMEOUT = 120  # Maximum seconds to wait for lock
LOCK_RETRY_INTERVAL = 0.5  # Base interval between retries
MAX_WRITE_RETRIES = 10  # Maximum file write attempts


class DocxWriteError(Exception):
    """Raised when docx write fails after all retries."""
    pass


class LockTimeoutError(Exception):
    """Raised when unable to acquire lock within timeout."""
    pass


def get_docx_lock(output_path: str, timeout: float = LOCK_TIMEOUT):
    """
    Get a file lock context manager for a docx file.

    Use this to wrap custom save logic that needs process-level locking.

    Example:
        with get_docx_lock(output_path):
            doc = Document(output_path)
            # ... modify doc ...
            doc.save(output_path)

    Args:
        output_path: Path to the docx file to lock
        timeout: Maximum seconds to wait for lock

    Returns:
        Context manager that acquires/releases the lock
    """
    lock_path = output_path + ".lock"
    return file_lock(lock_path, timeout)


@contextmanager
def file_lock(lock_path: str, timeout: float = LOCK_TIMEOUT):
    """
    Cross-process file lock using a .lock file.

    Uses exclusive file creation to ensure only one process can hold the lock.
    Falls back to checking lock file age if stale (process crashed).

    Args:
        lock_path: Path to the lock file (e.g., "AI_OUTPUT.docx.lock")
        timeout: Maximum seconds to wait for lock

    Yields:
        None when lock is acquired

    Raises:
        LockTimeoutError: If unable to acquire lock within timeout
    """
    start_time = time.time()
    lock_acquired = False
    lock_fd = None

    # Stale lock threshold (if lock file older than this, assume dead process)
    STALE_THRESHOLD = 300  # 5 minutes

    while time.time() - start_time < timeout:
        try:
            # Try to create lock file exclusively
            # os.O_CREAT | os.O_EXCL ensures atomic create-if-not-exists
            lock_fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)

            # Write our PID to the lock file for debugging
            pid_bytes = f"{os.getpid()}\n{time.time()}".encode('utf-8')
            os.write(lock_fd, pid_bytes)

            lock_acquired = True
            break

        except FileExistsError:
            # Lock file exists - check if it's stale
            try:
                lock_age = time.time() - os.path.getmtime(lock_path)
                if lock_age > STALE_THRESHOLD:
                    # Lock is stale, remove it and retry
                    try:
                        os.remove(lock_path)
                        continue
                    except OSError:
                        pass  # Another process might have removed it
            except OSError:
                pass  # File might have been removed between checks

            # Wait with jitter before retrying
            jitter = random.uniform(0, LOCK_RETRY_INTERVAL)
            time.sleep(LOCK_RETRY_INTERVAL + jitter)

        except OSError as e:
            # Other OS errors (permission, etc.)
            logging.warning(f"Lock file error: {e}")
            time.sleep(LOCK_RETRY_INTERVAL)

    if not lock_acquired:
        raise LockTimeoutError(f"Could not acquire lock on {lock_path} within {timeout}s")

    try:
        yield
    finally:
        # Release lock
        if lock_fd is not None:
            try:
                os.close(lock_fd)
            except OSError:
                pass

        # Remove lock file
        try:
            os.remove(lock_path)
        except OSError:
            pass


def safe_append_to_docx(
    output_path: str,
    content: str,
    title: str,
    format_func: Optional[Callable] = None,
    logger: Optional[Callable] = None
) -> str:
    """
    Safely append content to a DOCX file with process-level locking.

    If the file is locked by another process, waits until available.
    If the file can't be opened (e.g., open in Word), tries versioned names.

    Args:
        output_path: Target DOCX file path
        content: Text content to add (can include markdown)
        title: Title/header for this section
        format_func: Optional function(doc, content) to format content
        logger: Optional logging function

    Returns:
        Actual path where content was saved (may differ if versioning used)

    Raises:
        DocxWriteError: If write fails after all retries
    """
    if Document is None:
        raise DocxWriteError("python-docx not installed")

    def log(msg, level="info"):
        if logger:
            if callable(logger):
                logger(msg)
            elif hasattr(logger, level):
                getattr(logger, level)(msg)
        else:
            try:
                print(msg, flush=True)
            except OSError:
                pass  # stdout pipe broken

    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    # Lock file path
    lock_path = output_path + ".lock"

    base_name, ext = os.path.splitext(output_path)
    version = 0
    current_path = output_path

    try:
        # Acquire process-level lock
        with file_lock(lock_path):
            log(f"Acquired write lock for: {os.path.basename(output_path)}")

            # Now try to write, with versioning fallback for open files
            for attempt in range(MAX_WRITE_RETRIES):
                try:
                    # Open or create document
                    if os.path.exists(current_path):
                        try:
                            doc = Document(current_path)
                            doc.add_page_break()
                            log(f"Appending to existing file: {current_path}")
                        except Exception as e:
                            # File might be corrupted or locked by another app
                            raise PermissionError(f"Cannot open file: {e}")
                    else:
                        doc = Document()
                        log(f"Creating new file: {current_path}")

                    # Apply default styles
                    _apply_default_styles(doc)

                    # Add title
                    p = doc.add_paragraph()
                    run = p.add_run(title)
                    run.bold = True
                    run.underline = True
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

                    # Add timestamp
                    doc.add_paragraph(
                        f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                    )

                    # Format and add content
                    if format_func:
                        format_func(doc, content)
                    else:
                        _add_markdown_to_doc(doc, content)

                    # Save document
                    doc.save(current_path)
                    log(f"Saved to: {current_path}")

                    return current_path

                except PermissionError as e:
                    # File locked by another app (e.g., Word) - try versioned name
                    version += 1
                    current_path = f"{base_name} v.{version}{ext}"
                    log(f"File locked, trying: {current_path}")

                    if attempt >= MAX_WRITE_RETRIES - 1:
                        raise DocxWriteError(
                            f"Could not write to {output_path} or any version after "
                            f"{MAX_WRITE_RETRIES} attempts"
                        )

                except Exception as e:
                    # Other errors - retry with backoff
                    if attempt >= MAX_WRITE_RETRIES - 1:
                        raise DocxWriteError(f"Write failed: {e}")

                    wait_time = (2 ** attempt) * 0.5 + random.uniform(0, 0.5)
                    log(f"Write error: {e}. Retrying in {wait_time:.1f}s...")
                    time.sleep(wait_time)

    except LockTimeoutError as e:
        raise DocxWriteError(f"Lock timeout: {e}")

    return current_path


def _apply_default_styles(doc):
    """Apply default Times New Roman 12pt styles to document."""
    if Pt is None:
        return

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.0

    for i in range(1, 10):
        if f'Heading {i}' in doc.styles:
            h_style = doc.styles[f'Heading {i}']
            h_style.font.name = 'Times New Roman'
            h_style.font.size = Pt(12)
            h_style.paragraph_format.line_spacing = 1.0


def _add_markdown_to_doc(doc, content: str):
    """Parse basic markdown and add to document."""
    import re

    if Pt is None or Inches is None:
        # Fallback: just add as plain text
        doc.add_paragraph(content)
        return

    lines = content.split('\n')
    active_paragraph = None

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Headings
        if stripped.startswith('#'):
            text = stripped.lstrip('#').strip()
            if not text.endswith('.'):
                text += "."
            active_paragraph = doc.add_paragraph()
            run = active_paragraph.add_run(text + " ")
            run.bold = True
            continue

        # List items
        if stripped.startswith('* ') or stripped.startswith('- '):
            text = stripped[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.25)

            run = p.add_run("\t")
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

            # Parse bold markers
            parts = re.split(r'(\*\*.*?\*\*)', text)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    r = p.add_run(part[2:-2])
                    r.bold = True
                else:
                    r = p.add_run(part)
                r.font.name = 'Times New Roman'
                r.font.size = Pt(12)

            active_paragraph = None
            continue

        # Normal text
        p = active_paragraph if active_paragraph else doc.add_paragraph()

        parts = re.split(r'(\*\*.*?\*\*)', stripped)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part[2:-2])
                run.bold = True
            else:
                p.add_run(part)

        active_paragraph = None
