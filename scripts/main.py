"""Main entry point for NSE Data Fetcher v3.

Usage:
    python -m scripts.main                  # normal incremental update
    python -m scripts.main --full-refresh   # re-fetch from start_date
    python -m scripts.main --start 2025-06-01
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

    excel_path = config.excel_file

    # Determine start date
    if args.full_refresh or not os.path.exists(excel_path):
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
