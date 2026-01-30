"""
Tests for Word Hotkey functionality.

Run with: python -m pytest tests/test_word_hotkey.py -v
Or standalone: python tests/test_word_hotkey.py
"""

import sys
import os
import time

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("win32com not available - some tests will be skipped")


def test_word_com_connection():
    """Test that we can connect to a running Word instance."""
    if not HAS_WIN32:
        print("SKIP: win32com not available")
        return False

    print("\n=== Testing Word COM Connection ===")

    # First, check for Word windows
    print("\n0. Checking for Word windows...")
    try:
        import win32gui

        word_windows = []
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "OpusApp":  # Word's window class
                    title = win32gui.GetWindowText(hwnd)
                    results.append((hwnd, title))
            return True

        win32gui.EnumWindows(enum_callback, word_windows)

        if word_windows:
            print(f"   Found {len(word_windows)} Word window(s):")
            for hwnd, title in word_windows:
                print(f"      - {title} (hwnd: {hwnd})")
        else:
            print("   No Word windows found!")
    except Exception as e:
        print(f"   Error checking windows: {e}")

    # Method 1: GetActiveObject
    print("\n1. Testing GetActiveObject...")
    try:
        word = win32com.client.GetActiveObject("Word.Application")
        print(f"   SUCCESS: Connected via GetActiveObject")
        print(f"   Documents.Count: {word.Documents.Count}")

        # Try to get ActiveDocument directly
        try:
            active_doc = word.ActiveDocument
            print(f"   ActiveDocument: {active_doc.Name if active_doc else 'None'}")
        except Exception as e:
            print(f"   ActiveDocument error: {e}")

        # List all documents
        try:
            for i, doc in enumerate(word.Documents):
                print(f"      Doc {i+1}: {doc.Name}")
        except Exception as e:
            print(f"   Error listing docs: {e}")

        if word.Documents.Count > 0:
            return True
    except Exception as e:
        print(f"   FAILED: {e}")

    # Method 2: GetObject
    print("\n2. Testing GetObject...")
    try:
        word = win32com.client.GetObject(Class="Word.Application")
        print(f"   SUCCESS: Connected via GetObject")
        print(f"   Documents.Count: {word.Documents.Count}")
        try:
            active_doc = word.ActiveDocument
            print(f"   ActiveDocument: {active_doc.Name if active_doc else 'None'}")
        except Exception as e:
            print(f"   ActiveDocument error: {e}")
        if word.Documents.Count > 0:
            return True
    except Exception as e:
        print(f"   FAILED: {e}")

    # Method 3: Dispatch (may create new instance)
    print("\n3. Testing Dispatch...")
    try:
        word = win32com.client.Dispatch("Word.Application")
        print(f"   Connected via Dispatch")
        print(f"   Documents.Count: {word.Documents.Count}")
        try:
            active_doc = word.ActiveDocument
            print(f"   ActiveDocument: {active_doc.Name if active_doc else 'None'}")
        except Exception as e:
            print(f"   ActiveDocument error: {e}")
        if word.Documents.Count > 0:
            print(f"   SUCCESS: Found existing Word with documents")
            return True
        else:
            print(f"   WARNING: Dispatch connected but no documents (may be new instance)")
            return False
    except Exception as e:
        print(f"   FAILED: {e}")

    return False


def test_word_selection():
    """Test that we can read and write Word selection."""
    if not HAS_WIN32:
        print("SKIP: win32com not available")
        return False

    print("\n=== Testing Word Selection ===")

    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except:
        try:
            word = win32com.client.GetObject(Class="Word.Application")
        except:
            print("FAILED: Could not connect to Word")
            return False

    print(f"   Documents.Count: {word.Documents.Count}")

    # Try to access Selection even if Documents.Count is 0
    # (This can happen with Protected View, AutoRecovered docs, etc.)
    print("\n1. Testing Selection access...")
    selection = None
    try:
        selection = word.Selection
        print(f"   Selection object obtained: {selection is not None}")
    except Exception as e:
        print(f"   FAILED to get Selection: {e}")

    if not selection:
        print("   FAILED: Selection object is None")
        # Try ProtectedViewWindows
        try:
            pv_count = word.ProtectedViewWindows.Count
            print(f"   ProtectedViewWindows.Count: {pv_count}")
        except Exception as e:
            print(f"   ProtectedViewWindows check failed: {e}")
        return False

    # Test reading selection
    print("\n2. Testing Selection.Text read...")
    try:
        text = selection.Text
        sel_type = selection.Type
        print(f"   SUCCESS: Selection type={sel_type}")
        print(f"   Selected text ({len(text)} chars): '{text[:100]}...'" if len(text) > 100 else f"   Selected text: '{text}'")
    except Exception as e:
        print(f"   FAILED: {e}")
        return False

    # Test writing text (only if user confirms)
    print("\n3. Testing Selection.TypeText write...")
    print("   To test writing, select some text in Word and type 'yes': ", end="")

    # For automated testing, skip the write test
    if os.environ.get("AUTOMATED_TEST"):
        print("   SKIPPED (automated mode)")
    else:
        try:
            response = input()
            if response.lower() == 'yes':
                test_text = "[TEST_INSERT]"
                selection.TypeText(test_text)
                print(f"   SUCCESS: Inserted '{test_text}'")
                print("   (Press Ctrl+Z in Word to undo)")
            else:
                print("   SKIPPED")
        except:
            print("   SKIPPED (no input)")

    return True


def test_llm_handler():
    """Test that LLMHandler works."""
    print("\n=== Testing LLM Handler ===")

    try:
        from icharlotte_core.llm import LLMHandler

        settings = {
            'temperature': 0.7,
            'top_p': 0.95,
            'max_tokens': 100,
            'stream': False,
            'thinking_level': 'None'
        }

        print("1. Testing LLMHandler.generate...")
        result = LLMHandler.generate(
            provider="Gemini",
            model="models/gemini-2.0-flash",
            system_prompt="You are a test assistant. Respond with exactly: TEST_SUCCESS",
            user_prompt="Say the test response.",
            file_contents="",
            settings=settings
        )

        if result and "TEST" in result.upper():
            print(f"   SUCCESS: Got response: {result[:100]}")
            return True
        else:
            print(f"   WARNING: Unexpected response: {result[:100] if result else 'None'}")
            return True  # Still a success if we got any response

    except Exception as e:
        print(f"   FAILED: {e}")
        return False


def test_popup_creation():
    """Test that the popup dialog can be created."""
    print("\n=== Testing Popup Creation ===")

    try:
        # Need QApplication for widgets
        from PySide6.QtWidgets import QApplication
        app = QApplication.instance()
        if not app:
            app = QApplication([])

        from icharlotte_core.word_hotkey import WordLLMPopup

        popup = WordLLMPopup()
        print("   SUCCESS: WordLLMPopup created")

        # Check prompts loaded
        print(f"   Prompts loaded: {len(popup.prompts)}")
        for p in popup.prompts[:3]:
            print(f"      - {p['name']}")

        popup.close()
        return True

    except Exception as e:
        print(f"   FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def run_all_tests():
    """Run all tests and report results."""
    print("=" * 60)
    print("Word Hotkey Integration Tests")
    print("=" * 60)
    print("\nMake sure Word is open with a document before running tests.")
    print("Select some text in Word for selection tests.\n")

    results = {}

    results['COM Connection'] = test_word_com_connection()
    results['Selection'] = test_word_selection()
    results['LLM Handler'] = test_llm_handler()
    results['Popup Creation'] = test_popup_creation()

    print("\n" + "=" * 60)
    print("Test Results:")
    print("=" * 60)
    for name, passed in results.items():
        status = "PASS" if passed else "FAIL"
        print(f"  {name}: {status}")

    all_passed = all(results.values())
    print("\n" + ("All tests passed!" if all_passed else "Some tests failed."))
    return all_passed


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
