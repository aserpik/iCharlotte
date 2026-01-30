"""
Enumerate all Word processes and try to connect to each one.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client
import win32gui
import win32process
import pythoncom
import subprocess

print("=== Word Process Analysis ===\n")

# 1. Find all WINWORD.EXE processes
print("1. Finding all WINWORD.EXE processes...")
try:
    result = subprocess.run(
        ['tasklist', '/FI', 'IMAGENAME eq WINWORD.EXE', '/FO', 'CSV'],
        capture_output=True, text=True
    )
    lines = result.stdout.strip().split('\n')
    if len(lines) > 1:
        print(f"   Found {len(lines)-1} Word process(es):")
        for line in lines[1:]:
            parts = line.replace('"', '').split(',')
            if len(parts) >= 2:
                print(f"      PID: {parts[1]}, Name: {parts[0]}")
    else:
        print("   No Word processes found")
except Exception as e:
    print(f"   Error: {e}")

# 2. Find all Word windows and their process IDs
print("\n2. Finding Word windows and their PIDs...")
word_windows = []

def enum_callback(hwnd, results):
    if win32gui.IsWindowVisible(hwnd):
        class_name = win32gui.GetClassName(hwnd)
        if class_name == "OpusApp":
            title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            results.append((hwnd, pid, title))
    return True

win32gui.EnumWindows(enum_callback, word_windows)

if word_windows:
    print(f"   Found {len(word_windows)} Word window(s):")
    for hwnd, pid, title in word_windows:
        print(f"      PID: {pid}, Title: {title}")
else:
    print("   No Word windows found")

# 3. Try GetActiveObject and see which instance it connects to
print("\n3. Testing GetActiveObject...")
try:
    pythoncom.CoInitialize()
    word = win32com.client.GetActiveObject("Word.Application")
    print(f"   Connected to Word")
    print(f"   Documents.Count: {word.Documents.Count}")

    # Try to get the process ID of this COM instance
    # This is tricky - COM doesn't directly expose PID
    # But we can check the window handle
    try:
        if word.Documents.Count > 0:
            hwnd = word.ActiveWindow.Hwnd
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            print(f"   Active Window PID: {pid}")
    except Exception as e:
        print(f"   Could not get PID: {e}")

except Exception as e:
    print(f"   GetActiveObject failed: {e}")

# 4. Show unique PIDs
print("\n4. Summary:")
window_pids = set(pid for _, pid, _ in word_windows)
print(f"   Unique PIDs from visible windows: {window_pids}")
print(f"   Total visible Word windows: {len(word_windows)}")

print("\n=== Recommendation ===")
if len(word_windows) > 0:
    print("There are visible Word windows, but COM may be connecting to a")
    print("hidden/background Word process. Try:")
    print("  1. Open Task Manager")
    print("  2. Go to 'Details' tab")
    print("  3. Sort by 'Name' and find all WINWORD.EXE processes")
    print("  4. End any processes that don't match the PIDs above")
    print(f"     (Keep PIDs: {window_pids})")
