import os

# --- Configuration ---
BASE_PATH_WIN = r"Z:\\Shared\\Current Clients"
SCRIPTS_DIR = os.path.join(os.getcwd(), "Scripts")
LOG_FILE = os.path.join(os.getcwd(), "icharlotte_Activity.log")
GEMINI_DATA_DIR = os.path.join(os.getcwd(), ".gemini", "case_data")
TEMP_DIR = os.path.join(os.getcwd(), ".gemini", "tmp")

# Templates & Resources
TEMPLATES_DIR = os.path.join(os.getcwd(), "Templates")
RESOURCES_DIR = r"C:\geminiterminal2\LLM Resources"
TEMPLATE_EXTENSIONS = ['.docx', '.txt', '.html', '.rtf']
RESOURCE_EXTENSIONS = ['.pdf', '.docx', '.doc', '.txt', '.html', '.png', '.jpg', '.jpeg']

# Normalize Gemini API key environment variables
# The google-genai SDK checks both GOOGLE_API_KEY and GEMINI_API_KEY, causing warnings if both are set
# We prefer GEMINI_API_KEY but support GOOGLE_API_KEY as fallback
_gemini_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
if _gemini_key:
    # Set GEMINI_API_KEY and remove GOOGLE_API_KEY to prevent SDK warning
    os.environ["GEMINI_API_KEY"] = _gemini_key
    if "GOOGLE_API_KEY" in os.environ:
        del os.environ["GOOGLE_API_KEY"]

API_KEYS = {
    "Gemini": os.environ.get("GEMINI_API_KEY"),
    "OpenAI": os.environ.get("OPENAI_API_KEY"),
    "Claude": os.environ.get("ANTHROPIC_API_KEY")
}

