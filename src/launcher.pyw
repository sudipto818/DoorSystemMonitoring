"""
launcher.pyw
─────────────────────────────────────────────────────
Single entry-point: starts display_app in the background,
then opens control_app.  Double-click this (or its .exe)
to launch the entire Door Sign system.
"""

import subprocess
import sys
import os

# When frozen as .exe by PyInstaller, sys.executable is the .exe path on disk.
# We need the FOLDER where the .exe lives (not the temp extraction dir).
if getattr(sys, 'frozen', False):
    # Frozen .exe — find sibling exes in the same folder
    base_dir = os.path.dirname(os.path.abspath(sys.executable))
    display_path = os.path.join(base_dir, "display_app.exe")
    control_path = os.path.join(base_dir, "control_app.exe")
else:
    # Running as .py script
    base_dir = os.path.dirname(os.path.abspath(__file__))
    display_path = os.path.join(base_dir, "display_app.py")
    control_path = os.path.join(base_dir, "control_app.py")

# Verify files exist before launching
if not os.path.exists(display_path):
    import ctypes
    ctypes.windll.user32.MessageBoxW(
        0,
        f"Cannot find:\n{display_path}\n\nMake sure all 3 .exe files are in the same folder.",
        "Door Sign Launcher — Error",
        0x10  # MB_ICONERROR
    )
    sys.exit(1)

if not os.path.exists(control_path):
    import ctypes
    ctypes.windll.user32.MessageBoxW(
        0,
        f"Cannot find:\n{control_path}\n\nMake sure all 3 .exe files are in the same folder.",
        "Door Sign Launcher — Error",
        0x10
    )
    sys.exit(1)

# Launch display in background (no console window)
if getattr(sys, 'frozen', False):
    subprocess.Popen(
        [display_path],
        cwd=base_dir,
        creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS
    )
    # Launch control panel
    subprocess.Popen(
        [control_path],
        cwd=base_dir
    )
else:
    pythonw = sys.executable.replace("python.exe", "pythonw.exe")
    subprocess.Popen(
        [pythonw, display_path],
        cwd=base_dir,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    subprocess.Popen(
        [sys.executable, control_path],
        cwd=base_dir
    )
