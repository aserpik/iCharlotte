#!/usr/bin/env python3
"""
Patch Claude Code exe to fix Chrome MCP extension on Windows.

The socket-path discovery function ZKI() only scans for Unix .sock files
and never returns the Windows named pipe path. This patch adds an early
return on win32 so it returns the correct named pipe.

See: https://github.com/anthropics/claude-code/issues/21371
"""
import os
import sys
import shutil

exe_path = os.path.expanduser("~/.local/bin/claude.exe")
backup_path = exe_path + ".pre-chrome-patch"

# Restore mode
if "--restore" in sys.argv:
    if not os.path.exists(backup_path):
        print("No backup found at:", backup_path)
        sys.exit(1)
    # claude.exe may be running, rename it first
    running_path = exe_path + ".running"
    if os.path.exists(running_path):
        os.remove(running_path)
    try:
        os.rename(exe_path, running_path)
    except PermissionError:
        pass
    shutil.copy2(backup_path, exe_path)
    print("Restored claude.exe from backup.")
    sys.exit(0)

if not os.path.exists(exe_path):
    print("claude.exe not found at:", exe_path)
    sys.exit(1)

print(f"Reading: {exe_path}")
with open(exe_path, "rb") as f:
    data = f.read()

original_size = len(data)
print(f"Original size: {original_size:,} bytes")

# ── Step 1: Find the BROKEN function ZKI ──
# The broken function looks like:
# function ZKI(){let H=[],$=ZkH();try{let I=PKI.readdirSync($);for(let f of I)if(f.endsWith(".sock"))H.push(mY.join($,f))}catch{}let A=`claude-mcp-browser-bridge-${dzA()}`,L=mY.join(gy.tmpdir(),A),D=`/tmp/${A}`;if(!H.includes(L))H.push(L);if(L!==D&&!H.includes(D))H.push(D);return H}

# Search for the unique pattern in the broken function
search_pattern = b'function ZKI(){let H=[],$=ZkH();try{let I=PKI.readdirSync($);for(let f of I)if(f.endsWith(".sock"))H.push(mY.join($,f))}catch{}let A=`claude-mcp-browser-bridge-${dzA()}`,L=mY.join(gy.tmpdir(),A),D=`/tmp/${A}`;if(!H.includes(L))H.push(L);if(L!==D&&!H.includes(D))H.push(D);return H}'

count = data.count(search_pattern)
print(f"Found broken function {count} time(s)")

if count == 0:
    # Check if already patched
    if data.count(b'function ZKI(){if(gy.platform()==="win32")') > 0:
        print("\nAlready patched! Nothing to do.")
        sys.exit(0)
    print("Could not find the broken function. The minified code may have changed.")
    sys.exit(1)

# ── Step 2: Build the replacement ──
# The pipe path uses the same format as the working QZ$ function.
# In the binary, the template literal for the pipe is:
#   \\\\.\\pipe\\claude-mcp-browser-bridge-${dzA()}
# (4 backslash bytes, dot, 2 backslash bytes, "pipe", 2 backslash bytes, then interpolation)

# We add the win32 check at the start and remove the /tmp duplicate path logic
# to keep the same byte length.

# The win32 early return:
win32_prefix = b'function ZKI(){if(gy.platform()==="win32")return[`\\\\\\\\.\\\\pipe\\\\claude-mcp-browser-bridge-${dzA()}`];let H=[],$=ZkH();try{let I=PKI.readdirSync($);for(let f of I)if(f.endsWith(".sock"))H.push(mY.join($,f))}catch{}let A=`claude-mcp-browser-bridge-${dzA()}`,L=mY.join(gy.tmpdir(),A),D=`/tmp/${A}`;if(!H.includes(L))H.push(L);if(L!==D&&!H.includes(D))H.push(D);return H}'

search_len = len(search_pattern)
replace_len = len(win32_prefix)
print(f"Search length:  {search_len}")
print(f"Replace length: {replace_len}")

if replace_len > search_len:
    # Need to trim some code to fit. Remove the /tmp duplicate check:
    # ,D=`/tmp/${A}`;if(L!==D&&!H.includes(D))H.push(D)
    # and pad with spaces before return

    # Build without the /tmp fallback
    replacement = b'function ZKI(){if(gy.platform()==="win32")return[`\\\\\\\\.\\\\pipe\\\\claude-mcp-browser-bridge-${dzA()}`];let H=[],$=ZkH();try{let I=PKI.readdirSync($);for(let f of I)if(f.endsWith(".sock"))H.push(mY.join($,f))}catch{}let A=`claude-mcp-browser-bridge-${dzA()}`,L=mY.join(gy.tmpdir(),A);if(!H.includes(L))H.push(L);'

    diff = search_len - len(replacement) - len(b'return H}')
    if diff < 0:
        # Still too long, also remove the tmpdir fallback
        replacement = b'function ZKI(){if(gy.platform()==="win32")return[`\\\\\\\\.\\\\pipe\\\\claude-mcp-browser-bridge-${dzA()}`];let H=[],$=ZkH();try{let I=PKI.readdirSync($);for(let f of I)if(f.endsWith(".sock"))H.push(mY.join($,f))}catch{}'
        diff = search_len - len(replacement) - len(b'return H}')

    if diff < 0:
        print(f"ERROR: Replacement still too long by {-diff} bytes")
        sys.exit(1)

    # Pad with spaces before return
    replacement = replacement + b' ' * diff + b'return H}'
    print(f"Final replacement length: {len(replacement)} (trimmed /tmp path, padded {diff} spaces)")
else:
    # Pad with spaces if shorter (unlikely but handle it)
    diff = search_len - replace_len
    # Insert spaces before the closing return H}
    replacement = win32_prefix[:-len(b'return H}')] + b' ' * diff + b'return H}'
    print(f"Final replacement length: {len(replacement)} (padded {diff} spaces)")

assert len(replacement) == search_len, f"Length mismatch: {len(replacement)} != {search_len}"

# ── Step 3: Create backup ──
if not os.path.exists(backup_path):
    shutil.copy2(exe_path, backup_path)
    print(f"Backup created: {backup_path}")
else:
    print(f"Backup already exists: {backup_path}")

# ── Step 4: Apply patch ──
patched = data.replace(search_pattern, replacement)
patched_count = original_size - len(patched) + len(patched)  # sanity

# Verify size unchanged
assert len(patched) == original_size, f"Size changed! {len(patched)} != {original_size}"

# Count patches applied
patches_applied = count  # we replaced all occurrences

# Since claude.exe may be running, rename it first
running_path = exe_path + ".running"
renamed = False
try:
    if os.path.exists(running_path):
        os.remove(running_path)
    os.rename(exe_path, running_path)
    renamed = True
    print("Renamed running exe to .running")
except PermissionError:
    print("Warning: Could not rename running exe, writing directly")

with open(exe_path, "wb") as f:
    f.write(patched)

print(f"\nPATCH APPLIED SUCCESSFULLY ({patches_applied} occurrence(s))")
print(f"File size: {len(patched):,} bytes (unchanged)")

# ── Step 5: Verify ──
with open(exe_path, "rb") as f:
    verify = f.read()

remaining = verify.count(search_pattern)
patched_found = verify.count(b'function ZKI(){if(gy.platform()==="win32")')
print(f"Verification: {patched_found} patched function(s) found, {remaining} unpatched remaining")

if remaining > 0:
    print("WARNING: Some occurrences were not patched!")
elif patched_found >= count:
    print("\nAll occurrences patched successfully!")
    print("\nNext steps:")
    print("  1. Restart Chrome completely")
    print("  2. Start a new Claude Code session")
    print("  3. Run /chrome to initialize the MCP")
