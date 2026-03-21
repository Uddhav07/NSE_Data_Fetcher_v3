"""NSE Data Fetcher v3 — Desktop application entry point.

Launch with:
    python app.py          (development)
    NSE Data Fetcher.exe   (frozen / installed)
"""

import sys
import os

# When frozen by PyInstaller, ensure the base directory is the working dir
# so that relative paths in config (e.g. "data/...") resolve correctly.
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))


def _fatal_error(title: str, msg: str) -> None:
    """Show a native error dialog even if tkinter only partially loaded."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, msg)
        root.destroy()
    except Exception:
        pass
    # Also write to a crash file so the user always has a record
    try:
        crash_dir = (os.path.dirname(sys.executable)
                     if getattr(sys, "frozen", False) else ".")
        crash_path = os.path.join(crash_dir, "CRASH_LOG.txt")
        with open(crash_path, "w", encoding="utf-8") as f:
            f.write(f"{title}\n\n{msg}\n")
    except Exception:
        pass


if __name__ == "__main__":
    try:
        from scripts.gui import main
        main()
    except Exception as exc:
        import traceback
        _fatal_error(
            "NSE Data Fetcher — Fatal Error",
            f"The application encountered an unexpected error and must close.\n\n"
            f"{type(exc).__name__}: {exc}\n\n"
            f"Details:\n{traceback.format_exc()}\n\n"
            f"If this keeps happening, delete config/config.json and try again.",
        )
        sys.exit(1)
