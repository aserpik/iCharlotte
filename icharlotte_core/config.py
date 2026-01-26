import os

# --- Configuration ---
BASE_PATH_WIN = r"Z:\\Shared\\Current Clients"
SCRIPTS_DIR = os.path.join(os.getcwd(), "Scripts")
LOG_FILE = os.path.join(os.getcwd(), "icharlotte_Activity.log")
GEMINI_DATA_DIR = os.path.join(os.getcwd(), ".gemini", "case_data")
TEMP_DIR = os.path.join(os.getcwd(), ".gemini", "tmp")
NOTETAKER_DIR = os.path.join(os.getcwd(), "NoteTaker", "out", "renderer")

# Templates & Resources
TEMPLATES_DIR = os.path.join(os.getcwd(), "Templates")
RESOURCES_DIR = r"C:\geminiterminal2\LLM Resources"
TEMPLATE_EXTENSIONS = ['.docx', '.txt', '.html', '.rtf']
RESOURCE_EXTENSIONS = ['.pdf', '.docx', '.doc', '.txt', '.html', '.png', '.jpg', '.jpeg']

API_KEYS = {
    "Gemini": os.environ.get("GEMINI_API_KEY"),
    "OpenAI": os.environ.get("OPENAI_API_KEY"),
    "Claude": os.environ.get("ANTHROPIC_API_KEY")
}

