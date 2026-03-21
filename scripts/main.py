"""Main entry point for NSE Data Fetcher v3.

Usage (CLI):
    python -m scripts.main                  # normal incremental update
    python -m scripts.main --full-refresh   # re-fetch from start_date
    python -m scripts.main --start 2026-06-01
    python -m scripts.main --no-futures     # skip futures scraping

Usage (GUI):
    python app.py                           # launch desktop GUI
"""

import argparse
import logging
import os
import sys
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional, Callable

from . import __version__
from .paths import get_logs_dir, resolve_relative, ensure_dirs, get_base_dir
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

def setup_logging(config: AppConfig, extra_handler: Optional[logging.Handler] = None) -> None:
    """Configure root logger with console + rotating file handlers.

    Parameters
    ----------
    config : AppConfig
    extra_handler : optional additional handler (e.g. GUI text widget handler)
    """
    log_dir = get_logs_dir()
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
        if extra_handler:
            extra_handler.setFormatter(fmt)
            extra_handler.setLevel(level)
            root.addHandler(extra_handler)


# ── CLI ───────────────────────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=f"NSE Data Fetcher v{__version__} — Nifty 50 OHLC + Futures tracker",
    )
    parser.add_argument(
        "--config",
        help="Path to config JSON file (default: auto-detect)",
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


# ── Core fetch logic (called by both CLI and GUI) ────────────────────────────

class FetchResult:
    """Outcome of a fetch operation."""
    __slots__ = ("added", "skipped", "error", "excel_path")

    def __init__(self, added: int = 0, skipped: int = 0,
                 error: str = "", excel_path: str = ""):
        self.added = added
        self.skipped = skipped
        self.error = error
        self.excel_path = excel_path

    @property
    def ok(self) -> bool:
        return not self.error


def run_fetch(
    *,
    config_path: Optional[str] = None,
    start_override: Optional[str] = None,
    full_refresh: bool = False,
    no_futures: bool = False,
    open_excel: Optional[bool] = None,
) -> FetchResult:
    """Execute the data-fetch pipeline. Returns a FetchResult.

    This is the main workhorse used by **both** the CLI ``run()`` and the GUI.
    It never calls ``sys.exit()`` — errors are reported via ``FetchResult.error``.
    """
    # Ensure writable directories exist
    ensure_dirs()

    # Config
    save_default_config(config_path)
    config = load_config(config_path)

    logger.info("=" * 60)
    logger.info("NSE Data Fetcher v%s — %s", __version__,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("=" * 60)

    # Environment health check
    issues = _check_environment()
    for issue in issues:
        logger.warning("ENV CHECK: %s", issue)
    if any("Missing package" in i for i in issues):
        return FetchResult(error="Critical packages missing.")

    excel_path = str(resolve_relative(config.excel_file))

    # Determine start date
    if full_refresh or not os.path.exists(excel_path):
        if full_refresh and os.path.exists(excel_path):
            logger.info("Full refresh requested — recreating workbook.")
            backup_workbook(excel_path)
            try:
                os.remove(excel_path)
            except PermissionError:
                return FetchResult(
                    error=f"Cannot delete '{excel_path}' — the file is locked. "
                          "Please close it in Excel and try again.",
                    excel_path=excel_path,
                )

        if not os.path.exists(excel_path):
            logger.info("Excel file not found — creating new workbook.")
            create_workbook(excel_path)

        start = start_override or config.start_date
        logger.info("Fetching historical data from %s...", start)
    else:
        last = get_last_date(excel_path)
        if last is None:
            start = start_override or config.start_date
            logger.info("Excel is empty. Fetching from %s...", start)
        else:
            start = (last + timedelta(days=1)).strftime("%Y-%m-%d")
            logger.info("Last date in Excel: %s — fetching from %s...",
                        last.strftime("%Y-%m-%d"), start)

    # Backup before update
    if config.backup_before_update and os.path.exists(excel_path):
        backup_workbook(excel_path)

    # Fetch OHLC
    data = fetch_ohlc_data(config, start)

    if data is None or data.empty:
        logger.info("No new data available. You're up to date!")
        logger.info("=" * 60)
        should_open = open_excel if open_excel is not None else config.open_excel_after_run
        if should_open and os.path.exists(excel_path):
            _open_excel(excel_path)
        return FetchResult(excel_path=excel_path)

    # Fetch futures
    futures_price, futures_expiry = None, None
    if not no_futures:
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

    should_open = open_excel if open_excel is not None else config.open_excel_after_run
    if should_open and os.path.exists(excel_path):
        _open_excel(excel_path)

    return FetchResult(added=added, skipped=skipped, excel_path=excel_path)


# ── CLI entry ─────────────────────────────────────────────────────────────────

def run() -> None:
    """CLI entry point — parses args and delegates to run_fetch."""
    args = _parse_args()

    ensure_dirs()
    save_default_config(args.config)
    config = load_config(args.config)
    setup_logging(config)

    result = run_fetch(
        config_path=args.config,
        start_override=args.start,
        full_refresh=args.full_refresh,
        no_futures=args.no_futures,
    )

    if not result.ok:
        logger.error(result.error)
        sys.exit(1)


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
