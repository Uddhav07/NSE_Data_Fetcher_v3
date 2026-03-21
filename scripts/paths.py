"""Centralized path resolution for both development and frozen (PyInstaller) modes."""

import os
import sys
from pathlib import Path


def is_frozen() -> bool:
    """True when running as a PyInstaller bundle."""
    return getattr(sys, "frozen", False)


def get_base_dir() -> Path:
    """Return the root directory for writable data (config, data, logs, archive).

    - Frozen: directory containing the .exe
    - Dev: project root (parent of scripts/)
    """
    if is_frozen():
        return Path(os.path.dirname(sys.executable))
    return Path(__file__).resolve().parent.parent


def get_bundle_dir() -> Path:
    """Return the directory where bundled read-only resources live.

    - Frozen: PyInstaller's _MEIPASS temp folder
    - Dev: same as base_dir (project root)
    """
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)))
    return get_base_dir()


def get_config_path() -> Path:
    """Path to config.json (writable copy alongside exe or in project root)."""
    return get_base_dir() / "config" / "config.json"


def get_data_dir() -> Path:
    return get_base_dir() / "data"


def get_logs_dir() -> Path:
    return get_base_dir() / "logs"


def get_archive_dir() -> Path:
    return get_base_dir() / "archive"


def ensure_dirs() -> None:
    """Create required writable directories if they don't exist."""
    for d in (get_data_dir(), get_logs_dir(), get_archive_dir(),
              get_config_path().parent):
        d.mkdir(parents=True, exist_ok=True)


def resolve_relative(path_str: str) -> Path:
    """Resolve a path string relative to the base directory."""
    p = Path(path_str)
    if p.is_absolute():
        return p
    return get_base_dir() / p
