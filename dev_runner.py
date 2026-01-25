import sys
import subprocess
import os

if __name__ == "__main__":
    target_script = "iCharlotte.py"
    script_args = [target_script] + sys.argv[1:]
    
    print(f"[DevRunner] Starting: {target_script}")
    print(f"[DevRunner] Note: Auto-restart on .py changes is DISABLED. Use the Restart button in the app UI.")
    
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    
    try:
        # Start the process and wait for it
        # Note: iCharlotte uses os.execl to restart itself, which on Windows
        # creates a NEW process and the old one (this subprocess) will exit.
        process = subprocess.Popen([sys.executable] + script_args, env=env)
        process.wait()
    except KeyboardInterrupt:
        print("\n[DevRunner] Stopping...")
        if process:
            if os.name == 'nt':
                subprocess.run(["taskkill", "/F", "/T", "/PID", str(process.pid)],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                process.terminate()