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

if __name__ == "__main__":
    try:
        from humble_mumble.google_tasks.google_tasks_gui import main
        main()
    except Exception as e:
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Google Tasks Studio Launcher Error",
                f"Failed to launch Google Tasks Sync Studio:\n\n{e}\n\n"
                "Please verify dependencies are installed:\n"
                "pip install -r requirements.txt"
            )
        except Exception:
            print(f"Error launching Google Tasks Studio: {e}", file=sys.stderr)
        sys.exit(1)
