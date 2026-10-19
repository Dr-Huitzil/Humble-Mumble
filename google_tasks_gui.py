"""
Humble Mumble - Google Tasks Studio Launcher
=============================================
Launches the Google Tasks synchronization desktop GUI.
Can also be launched via:
    humble-tasks
    python -m humble_mumble.google_tasks.google_tasks_gui
"""
import os
import sys

# Ensure src/ is on sys.path for direct execution from git checkout
_SRC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src")
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from humble_mumble.google_tasks.google_tasks_gui import main

if __name__ == "__main__":
    main()
