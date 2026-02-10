"""
Deposition Testimony Extraction Module for iCharlotte

Three-stage pipeline:
  1. Parse: PDF transcript → structured Q/A index (deterministic)
  2. Select: LLM identifies relevant exchanges by ID
  3. Output: Verbatim text → Word document + optional PDF highlighting
"""

from .models import QAExchange, TranscriptIndex, DeponentInfo
from .transcript_parser import TranscriptParser
from .testimony_selector import TestimonySelector
from .testimony_formatter import TestimonyFormatter
