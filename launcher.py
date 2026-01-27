"""Simple launcher for iCharlotte - creates a proper Windows exe identity."""
import sys
import os
import ctypes

# Set the App User Model ID before anything else
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('iCharlotte.LegalSuite.1')

# Change to the script directory
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Add the directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import and run the main module
if __name__ == "__main__":
    # Execute the main script
    exec(open("iCharlotte.py").read())
