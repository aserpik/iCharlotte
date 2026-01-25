import os
import sys
import logging
import datetime
import time

# Third-party imports
try:
    from google import genai
except ImportError as e:
    print(f"Critical Error: Missing dependency {e}. Please ensure google-genai is installed.", file=sys.stderr)
    sys.exit(1)

# Default configuration
DEFAULT_MODEL = "gemini-2.0-flash"
FALLBACK_MODELS = ["gemini-1.5-flash", "gemini-1.5-pro"]

def log_event(message, level="info"):
    """
    Standardized logging for Gemini Utils.
    """
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    msg = f"[{timestamp}] [GeminiUtils] {message}"
    try:
        print(msg, file=sys.stderr)
    except UnicodeEncodeError:
        try:
            encoding = sys.stderr.encoding or 'utf-8'
            print(msg.encode(encoding, errors='replace').decode(encoding), file=sys.stderr)
        except Exception:
            print(msg.encode('ascii', errors='replace').decode('ascii'), file=sys.stderr)
    sys.stderr.flush()

def clean_json_string(s: str) -> str:
    """Cleans Markdown code blocks and extra text to get raw JSON."""
    s = s.strip()
    # Remove markdown code blocks if present
    if "```" in s:
        lines = s.split('\n')
        # Filter out lines that start with ```
        lines = [l for l in lines if not l.strip().startswith("```")]
        s = "\n".join(lines)
    
    # Find the first '{' and last '}'
    start = s.find('{')
    end = s.rfind('}')
    
    if start != -1 and end != -1 and end > start:
        return s[start:end+1]
    
    return s.strip()

def call_gemini_api(prompt: str, context_text: str = "", models: list = None) -> str:
    """
    Calls Gemini with fallback models using the new google-genai SDK.
    
    Args:
        prompt: The instruction for the AI.
        context_text: The document or data to analyze (optional).
        models: List of model names to try. Defaults to [DEFAULT_MODEL] + FALLBACK_MODELS.
        
    Returns:
        The text response from the AI, or None if all attempts fail.
    """
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        log_event("GEMINI_API_KEY not set.", level="error")
        return None

    client = genai.Client(api_key=api_key)
    
    if models is None:
        models = [DEFAULT_MODEL] + FALLBACK_MODELS

    # Construct prompt
    if context_text:
        limit = 10000000
        truncated_context = context_text[:limit]
        if len(context_text) > limit:
            truncated_context += "\n...[TRUNCATED]"
            
        full_prompt = f"{prompt}\n\n[CONTEXT DATA]\n{truncated_context}"
    else:
        full_prompt = prompt

    for model_name in models:
        try:
            # log_event(f"Querying model: {model_name}") 
            # Note: client.models.generate_content returns a response object with a .text attribute
            response = client.models.generate_content(
                model=model_name,
                contents=full_prompt
            )
            if response and response.text:
                return response.text
        except Exception as e:
            log_event(f"Model {model_name} failed: {e}", level="warning")
            time.sleep(1) # Brief pause before retry
            continue
    
    log_event("All AI models failed.", level="error")
    return None
