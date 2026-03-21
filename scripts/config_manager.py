"""Configuration manager with validation and defaults."""

import json
import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import logging

from .paths import get_config_path, get_base_dir

logger = logging.getLogger(__name__)

# Supported tickers for quick reference
KNOWN_TICKERS = {
    "^NSEI": "Nifty 50",
    "^NSEBANK": "Bank Nifty",
    "^BSESN": "Sensex",
}


@dataclass
class AppConfig:
    """Application configuration with defaults and validation."""

    ticker: str = "^NSEI"
    excel_file: str = "data/stock_market_NSE50.xlsx"
    start_date: str = ""  # Computed dynamically: Jan 1st of current year
    log_level: str = "INFO"
    max_retries: int = 3
    request_timeout: int = 15
    open_excel_after_run: bool = True
    backup_before_update: bool = True

    def __post_init__(self) -> None:
        if not self.start_date:
            self.start_date = f"{datetime.now().year}-01-01"

    def validate(self) -> list[str]:
        """Return a list of validation errors (empty if valid)."""
        errors: list[str] = []

        # Validate start_date format
        try:
            datetime.strptime(self.start_date, "%Y-%m-%d")
        except ValueError:
            errors.append(f"Invalid start_date format: '{self.start_date}'. Use YYYY-MM-DD.")

        # Validate ticker is non-empty
        if not self.ticker.strip():
            errors.append("Ticker cannot be empty.")

        # Validate log level
        valid_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}
        if self.log_level.upper() not in valid_levels:
            errors.append(f"Invalid log_level: '{self.log_level}'. Must be one of {valid_levels}.")

        if self.max_retries < 0:
            errors.append("max_retries must be >= 0.")

        if self.request_timeout < 1:
            errors.append("request_timeout must be >= 1.")

        return errors


def load_config(config_path: str | None = None) -> AppConfig:
    """Load configuration from JSON file, falling back to defaults."""
    if config_path is None:
        config_path = str(get_config_path())
    config = AppConfig()

    if not os.path.exists(config_path):
        logger.warning("Config file '%s' not found — using defaults.", config_path)
        return config

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError) as exc:
        logger.error("Failed to read config file '%s': %s — using defaults.", config_path, exc)
        return config

    # Map known keys from JSON into the dataclass
    known_fields = {f.name for f in config.__dataclass_fields__.values()}
    for key, value in data.items():
        if key in known_fields:
            # Type-check: don't write garbage types into fields
            expected = type(getattr(config, key))
            if isinstance(value, expected) or (expected is int and isinstance(value, (int, float))):
                if expected is int and isinstance(value, float):
                    value = int(value)
                setattr(config, key, value)
            else:
                logger.warning(
                    "Config key '%s': expected %s, got %s — using default.",
                    key, expected.__name__, type(value).__name__,
                )
        elif key not in ("description", "notes"):
            logger.warning("Unknown config key ignored: '%s'", key)

    # Validate
    errors = config.validate()
    if errors:
        for err in errors:
            logger.error("Config validation error: %s", err)
        raise ValueError(f"Configuration has {len(errors)} error(s): {'; '.join(errors)}")

    return config


def save_default_config(config_path: str | None = None) -> None:
    """Write a default config file if none exists."""
    if config_path is None:
        config_path = str(get_config_path())
    path = Path(config_path)
    if path.exists():
        return

    path.parent.mkdir(parents=True, exist_ok=True)
    default = {
        "ticker": "^NSEI",
        "excel_file": "data/stock_market_NSE50.xlsx",
        "start_date": f"{datetime.now().year}-01-01",
        "log_level": "INFO",
        "max_retries": 3,
        "request_timeout": 15,
        "open_excel_after_run": True,
        "backup_before_update": True,
        "description": "Configuration for NSE Data Fetcher v3",
        "notes": [
            "ticker: Yahoo Finance symbol (^NSEI = Nifty 50, ^NSEBANK = Bank Nifty)",
            "excel_file: Relative path for the output Excel file",
            "start_date: First date to fetch historical data (YYYY-MM-DD). Defaults to Jan 1 of current year.",
            "log_level: DEBUG | INFO | WARNING | ERROR | CRITICAL",
            "max_retries: Number of retry attempts for failed HTTP requests",
            "request_timeout: HTTP request timeout in seconds",
            "open_excel_after_run: Open Excel file automatically after manual runs",
            "backup_before_update: Create a backup of Excel before updating",
        ],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(default, f, indent=4)
    logger.info("Created default config at %s", config_path)
