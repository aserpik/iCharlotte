
import sys
import os

# Mock necessary paths/environment if needed
sys.path.append(os.getcwd())

try:
    print("Attempting to import iCharlotte...")
    import iCharlotte
    print("Import successful!")
except ImportError as e:
    print(f"ImportError: {e}")
except Exception as e:
    print(f"Exception during import: {e}")
