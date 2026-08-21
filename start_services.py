#!/usr/bin/env python3
"""Starts both the static server (5175) and the save server (5176)."""
import os, subprocess, sys, time
from pathlib import Path

DIR = Path(__file__).parent
os.chdir(str(DIR))

# Start save server (tkinter dialog) on 5176
save_proc = subprocess.Popen(
    [sys.executable, str(DIR / "ys_desktop.py")],
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
)

# Start static server on 5175
static_proc = subprocess.Popen(
    [sys.executable, "-m", "http.server", "5175", "--bind", "127.0.0.1"],
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
)

print(f"Static: {static_proc.pid}, Save: {save_proc.pid}")
print("Both services running. Ctrl+C to stop.")

try:
    while True:
        time.sleep(10)
        # Check if both are alive
        if static_proc.poll() is not None:
            print("Static server died, restarting...")
            static_proc = subprocess.Popen(
                [sys.executable, "-m", "http.server", "5175", "--bind", "127.0.0.1"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        if save_proc.poll() is not None:
            print("Save server died, restarting...")
            save_proc = subprocess.Popen(
                [sys.executable, str(DIR / "ys_desktop.py")],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
except KeyboardInterrupt:
    static_proc.terminate()
    save_proc.terminate()
    print("Stopped.")
