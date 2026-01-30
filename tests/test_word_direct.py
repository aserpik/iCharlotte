"""
Direct test of Word COM without relying on GetActiveObject.
Creates a fresh Word instance to verify COM works properly.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client
import pythoncom
import time

print("=== Direct Word COM Test ===\n")

# Create a completely fresh Word instance
print("1. Creating fresh Word instance...")
try:
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")  # DispatchEx forces new instance
    word.Visible = True
    print(f"   Created new Word instance")
    print(f"   Documents.Count: {word.Documents.Count}")
except Exception as e:
    print(f"   FAILED: {e}")
    sys.exit(1)

# Create a new document
print("\n2. Creating new document...")
try:
    doc = word.Documents.Add()
    print(f"   Created document: {doc.Name}")
    print(f"   Documents.Count now: {word.Documents.Count}")
except Exception as e:
    print(f"   FAILED: {e}")
    word.Quit()
    sys.exit(1)

# Add some text
print("\n3. Adding text to document...")
try:
    word.Selection.TypeText("Hello, this is a test document.\n\nSelect this text and test the hotkey.")
    print(f"   Added text successfully")
except Exception as e:
    print(f"   FAILED: {e}")

# Test getting selection
print("\n4. Testing Selection access...")
try:
    # Select all
    word.Selection.WholeStory()
    selection = word.Selection
    text = selection.Text
    print(f"   Selection.Type: {selection.Type}")
    print(f"   Selection.Text: {text[:50]}...")
except Exception as e:
    print(f"   FAILED: {e}")

# Now try GetActiveObject to see if it finds this instance
print("\n5. Testing GetActiveObject...")
try:
    word2 = win32com.client.GetActiveObject("Word.Application")
    print(f"   GetActiveObject succeeded")
    print(f"   Documents.Count: {word2.Documents.Count}")
    if word2.Documents.Count > 0:
        print(f"   ActiveDocument: {word2.ActiveDocument.Name}")
except Exception as e:
    print(f"   GetActiveObject FAILED: {e}")

print("\n=== Test Complete ===")
print("A new Word window should now be visible with test text.")
print("You can close it manually or press Enter to close automatically.")

try:
    input("Press Enter to close Word and exit...")
except:
    pass

# Clean up
try:
    doc.Close(SaveChanges=False)
    word.Quit()
except:
    pass
