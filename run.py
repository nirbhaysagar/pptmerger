#!/usr/bin/env python3
"""
run.py - One-Click Launcher for PPT Unify
Ensures environment and dependencies, starts the server, and opens the web application in your browser.
"""

import os
import sys
import time
import subprocess
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if sys.platform == "win32":
    VENV_PYTHON = os.path.join(BASE_DIR, ".venv", "Scripts", "python.exe")
else:
    VENV_PYTHON = os.path.join(BASE_DIR, ".venv", "bin", "python")


def check_and_relaunch():
    # If running with system python and .venv python exists, re-exec with venv python
    if os.path.exists(VENV_PYTHON) and sys.executable != os.path.abspath(VENV_PYTHON):
        print(f"[PPT Unify] Launching with virtual environment: {VENV_PYTHON}")
        os.execv(VENV_PYTHON, [VENV_PYTHON, __file__] + sys.argv[1:])


def main():
    check_and_relaunch()

    # Import app
    sys.path.insert(0, BASE_DIR)
    try:
        from app import run_server
    except ImportError as e:
        print(f"[Error] Failed to import application: {e}")
        print("Please ensure required dependencies are installed: pip install python-pptx reportlab pypdf pillow")
        sys.exit(1)

    port = 5005
    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        port = int(sys.argv[1])

    url = f"http://127.0.0.1:{port}"

    # Open browser after a slight delay
    def open_browser():
        time.sleep(1.2)
        print(f"[PPT Unify] Opening web browser at {url} ...")
        webbrowser.open(url)

    import threading
    t = threading.Thread(target=open_browser, daemon=True)
    t.start()

    run_server(port=port)


if __name__ == "__main__":
    main()
