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

from scripts.gui import main

if __name__ == "__main__":
    main()
