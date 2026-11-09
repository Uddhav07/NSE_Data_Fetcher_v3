"""Centralized path resolution."""

from __future__ import annotations

from pathlib import Path


def get_base_dir() -> Path:
    """Project root (parent of scripts/)."""
    return Path(__file__).resolve().parent.parent


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
