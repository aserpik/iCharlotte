"""Test Win+V using SendInput (more reliable for RegisterHotKey)."""
import subprocess, time, os, ctypes
from ctypes import wintypes

# SendInput structures
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]
    _fields_ = [
        ("type", wintypes.DWORD),
        ("_input", _INPUT),
    ]

def send_input(*inputs):
    n = len(inputs)
    arr = (INPUT * n)(*inputs)
    ctypes.windll.user32.SendInput(n, arr, ctypes.sizeof(INPUT))

def make_key(vk, flags=0):
    inp = INPUT()
    inp.type = INPUT_KEYBOARD
    inp._input.ki.wVk = vk
    inp._input.ki.dwFlags = flags
    return inp

def focus_word():
    """Focus Word via FindWindow + SetForegroundWindow."""
    import win32gui
    hwnd = ctypes.windll.user32.FindWindowW("OpusApp", None)
    if not hwnd:
        print("Word window not found")
        return False

    # Attach input to steal focus
    fg = ctypes.windll.user32.GetForegroundWindow()
    fg_tid = ctypes.windll.user32.GetWindowThreadProcessId(fg, None)
    my_tid = ctypes.windll.kernel32.GetCurrentThreadId()

    ctypes.windll.user32.AllowSetForegroundWindow(-1)
    if fg_tid != my_tid:
        ctypes.windll.user32.AttachThreadInput(fg_tid, my_tid, True)

    ctypes.windll.user32.SetForegroundWindow(hwnd)
    ctypes.windll.user32.BringWindowToTop(hwnd)

    if fg_tid != my_tid:
        ctypes.windll.user32.AttachThreadInput(fg_tid, my_tid, False)

    time.sleep(0.5)

    fg = ctypes.windll.user32.GetForegroundWindow()
    sb = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(fg, sb, 256)
    print(f"Foreground class: {sb.value}")
    return sb.value == "OpusApp"

def send_winv():
    """Send Win+V using SendInput."""
    VK_LWIN = 0x5B
    VK_V = 0x56

    send_input(
        make_key(VK_LWIN),      # Win down
        make_key(VK_V),          # V down
    )
    time.sleep(0.05)
    send_input(
        make_key(VK_V, KEYEVENTF_KEYUP),      # V up
        make_key(VK_LWIN, KEYEVENTF_KEYUP),    # Win up
    )

def read_recent_log(n=10):
    log_path = os.path.join(os.path.dirname(__file__), 'logs', 'icharlotte_2026-03-06.log')
    if os.path.exists(log_path):
        with open(log_path) as f:
            lines = f.readlines()
        print(f"(total log lines: {len(lines)})")
        for line in lines[-n:]:
            print(line.rstrip())

if __name__ == '__main__':
    print("Step 1: Focus Word...")
    ok = focus_word()
    if not ok:
        print("WARNING: Could not focus Word")

    print("\nStep 2: Send Win+V via SendInput...")
    send_winv()
    print("Win+V sent")

    time.sleep(3)

    print("\n--- Recent log entries ---")
    read_recent_log()
