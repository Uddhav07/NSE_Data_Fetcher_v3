"""Main entry point for NSE Data Fetcher v3.

Usage:
    python -m scripts.main                  # normal incremental update
    python -m scripts.main --full-refresh   # re-fetch from start_date
    python -m scripts.main --start 2026-06-01
    python -m scripts.main --no-futures     # skip futures scraping
"""

import argparse
import logging
import os
import sys
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import __version__
from .config_manager import AppConfig, load_config, save_default_config
from .fetcher import fetch_ohlc_data
from .futures import fetch_futures
from .excel_writer import (
    backup_workbook,
    create_workbook,
    get_last_date,
    update_workbook,
)

logger = logging.getLogger("nse_fetcher")


# ── Health checks ─────────────────────────────────────────────────────────────

def _check_environment() -> list[str]:
    """Validate runtime environment. Returns list of issues (empty = OK)."""
    issues: list[str] = []

    # Python version
    if sys.version_info < (3, 10):
        issues.append(
            f"Python 3.10+ required (found {sys.version_info.major}.{sys.version_info.minor})"
        )

    # Required packages
    required = ["yfinance", "pandas", "openpyxl", "requests"]
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            issues.append(f"Missing package: {pkg}. Run: pip install -r requirements.txt")

    # Network check (lightweight)
    try:
        import requests as _req
        _req.head("https://www.google.com", timeout=5)
    except Exception:
        issues.append("No internet connection detected. Data fetch will fail.")

    return issues


# ── Logging setup ─────────────────────────────────────────────────────────────

def _setup_logging(config: AppConfig) -> None:
    """Configure root logger with console + rotating file handlers."""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    level = getattr(logging, config.log_level.upper(), logging.INFO)

    fmt = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Console handler
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(level)
    console.setFormatter(fmt)

    # File handler (rotating, 5 MB × 3 backups)
    file_handler = RotatingFileHandler(
        str(log_dir / "nse_fetcher.log"),
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)  # always capture debug in file
    file_handler.setFormatter(fmt)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    # Prevent duplicate handlers on repeated calls
    if not root.handlers:
        root.addHandler(console)
        root.addHandler(file_handler)


# ── CLI ───────────────────────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=f"NSE Data Fetcher v{__version__} — Nifty 50 OHLC + Futures tracker",
    )
    parser.add_argument(
        "--config", default="config/config.json",
        help="Path to config JSON file (default: config/config.json)",
    )
    parser.add_argument(
        "--start", metavar="YYYY-MM-DD",
        help="Override start date for data fetch",
    )
    parser.add_argument(
        "--full-refresh", action="store_true",
        help="Ignore existing data and re-fetch from start_date",
    )
    parser.add_argument(
        "--no-futures", action="store_true",
        help="Skip futures data scraping",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}",
    )
    return parser.parse_args()


# ── Main logic ────────────────────────────────────────────────────────────────

def run() -> None:
    args = _parse_args()

    # Ensure default config exists
    save_default_config(args.config)

    config = load_config(args.config)
    _setup_logging(config)

    logger.info("=" * 60)
    logger.info("NSE Data Fetcher v%s — %s", __version__, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("=" * 60)

    # Environment health check
    issues = _check_environment()
    for issue in issues:
        logger.warning("ENV CHECK: %s", issue)
    if any("Missing package" in i for i in issues):
        logger.error("Critical packages missing. Run setup.bat first.")
        sys.exit(1)

    excel_path = config.excel_file

    # Determine start date
    if args.full_refresh or not os.path.exists(excel_path):
        if args.full_refresh and os.path.exists(excel_path):
            logger.info("Full refresh requested — recreating workbook.")
            backup_workbook(excel_path)
            try:
                os.remove(excel_path)
            except PermissionError:
                logger.error(
                    "Cannot delete '%s' — the file is locked. "
                    "Please close it in Excel and try again.",
                    excel_path,
                )
                sys.exit(1)

        if not os.path.exists(excel_path):
            logger.info("Excel file not found — creating new workbook.")
            create_workbook(excel_path)

        start = args.start or config.start_date
        logger.info("Fetching historical data from %s...", start)
    else:
        # Incremental: pick up from day after last entry
        last = get_last_date(excel_path)
        if last is None:
            start = args.start or config.start_date
            logger.info("Excel is empty. Fetching from %s...", start)
        else:
            start = (last + timedelta(days=1)).strftime("%Y-%m-%d")
            logger.info("Last date in Excel: %s — fetching from %s...", last.strftime("%Y-%m-%d"), start)

    # Backup before update
    if config.backup_before_update and os.path.exists(excel_path):
        backup_workbook(excel_path)

    # Fetch OHLC
    data = fetch_ohlc_data(config, start)

    if data is None or data.empty:
        logger.info("No new data available. You're up to date!")
        logger.info("=" * 60)
        # Still open Excel if it exists so user can review
        if config.open_excel_after_run and os.path.exists(excel_path):
            _open_excel(excel_path)
        return

    # Fetch futures (only if not disabled)
    futures_price, futures_expiry = None, None
    if not args.no_futures:
        futures_price, futures_expiry = fetch_futures(config)

    # Write to Excel
    added, skipped = update_workbook(
        excel_path, data,
        futures_price=futures_price,
        futures_expiry=futures_expiry,
    )

    logger.info("=" * 60)
    logger.info("Update complete!  Added: %d  |  Skipped: %d", added, skipped)
    logger.info("=" * 60)

    # Auto-open Excel
    if config.open_excel_after_run and os.path.exists(excel_path):
        _open_excel(excel_path)


def _open_excel(path: str) -> None:
    """Open the Excel file in the default application."""
    try:
        os.startfile(os.path.abspath(path))
        logger.info("Opened %s", path)
    except Exception as exc:
        logger.warning("Could not open Excel file: %s", exc)


def main() -> None:
    """Safe entry point that catches unhandled exceptions."""
    try:
        run()
    except KeyboardInterrupt:
        print("\nAborted by user.")
        sys.exit(130)
    except Exception as exc:
        logger.critical("Fatal error: %s", exc, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
