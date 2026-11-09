"""Main entry point for NSE Data Fetcher v3.

Usage (CLI):
    python -m scripts.main                  # update: new days + any gaps/empty cells
    python -m scripts.main --full-refresh   # re-fetch from start_date
    python -m scripts.main --start 2026-06-01
    python -m scripts.main --no-futures     # skip futures scraping
    python -m scripts.main --no-rsi         # skip intraday RSI
    python -m scripts.main --refresh-recent # re-pull latest 3 rows (fixes values saved before close)

"""

from __future__ import annotations

import argparse
import logging
import os
import subprocess
import sys
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from typing import Optional

from . import __version__
from .bootstrap import ensure_dependencies

ensure_dependencies()  # installs/relaunches in the project venv if packages are missing

from .market import session_in_progress, today_str  # noqa: E402
from .paths import get_logs_dir, resolve_relative, ensure_dirs
from .config_manager import AppConfig, load_config, save_default_config
from .fetcher import fetch_ohlc_data
from .futures import fetch_futures
from .rsi import fetch_rsi
from .excel_writer import (
    backup_workbook,
    create_workbook,
    get_last_date,
    update_workbook,
)

# Windows consoles may use a legacy code page: never crash on characters like "—"
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

logger = logging.getLogger("nse_fetcher")


# ── Health checks ─────────────────────────────────────────────────────────────

def _check_environment() -> list[str]:
    """Validate runtime environment. Returns list of issues (empty = OK)."""
    issues: list[str] = []

    # Python version
    if sys.version_info < (3, 9):
        issues.append(
            f"Python 3.9+ required (found {sys.version_info.major}.{sys.version_info.minor})"
        )

    # Required packages
    required = ["yfinance", "pandas", "openpyxl", "requests"]
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            issues.append(f"Missing package: {pkg}. Run: pip install -r requirements.txt")

    # Network check (lightweight — test actual data source)
    try:
        import requests as _req
        _req.head("https://query1.finance.yahoo.com", timeout=5)
    except Exception:
        issues.append("Cannot reach Yahoo Finance. Check your internet connection.")

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
        "--no-rsi", action="store_true",
        help="Skip intraday RSI columns (R-U)",
    )
    parser.add_argument(
        "--refresh-recent", nargs="?", const=3, type=int, default=0, metavar="N",
        help="Overwrite Open/High/Low/Close and RSI of the latest N rows (default 3) with fresh "
             "data, even if the script wrote them earlier. Use once to correct values saved "
             "before a session was final.",
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
    no_rsi: bool = False,
    refresh_recent: int = 0,
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
    if any("Cannot reach" in i for i in issues):
        logger.warning("Network issue detected — will attempt fetch anyway.")

    excel_path = str(resolve_relative(config.excel_file))

    existed = os.path.exists(excel_path)

    # Prepare workbook / decide how far back to fetch
    if full_refresh and existed:
        logger.info("Full refresh requested — recreating workbook.")
        backup_workbook(excel_path)
        existed = False
        try:
            os.remove(excel_path)
        except PermissionError:
            return FetchResult(
                error=f"Cannot delete '{excel_path}' — the file is locked. "
                      "Please close it in Excel and try again.",
                excel_path=excel_path,
            )

    configured_start = start_override or config.start_date
    if not os.path.exists(excel_path):
        logger.info("Excel file not found — creating new workbook.")
        create_workbook(excel_path)
        start = configured_start
    else:
        last = get_last_date(excel_path)
        if last is None:
            start = configured_start
        else:
            # Re-check everything from the configured start date so any missing
            # days / empty cells are filled; filled cells are never overwritten.
            day_after = (last + timedelta(days=1)).strftime("%Y-%m-%d")
            start = min(configured_start, day_after)
            logger.info("Last date in Excel: %s — checking from %s for gaps and new days.",
                        last.strftime("%Y-%m-%d"), start)
    logger.info("Fetching data from %s...", start)

    # Backup before update (only if there is something worth saving)
    if config.backup_before_update and existed:
        backup_workbook(excel_path)

    # Fetch OHLC
    data = fetch_ohlc_data(config, start)
    if data is None:
        data = _empty_ohlc()

    # While the session is open, today's candle and futures price are not final.
    live = session_in_progress()
    if live:
        today = today_str()
        data = data[data.index.strftime("%Y-%m-%d") != today]
        logger.info("Market session still open (before 15:35 IST) — today's row and futures "
                    "are skipped. Run again after 15:35 IST.")

    futures_price, futures_expiry = None, None
    if not no_futures and not live:
        futures_price, futures_expiry = fetch_futures(config)
    rsi_map = {} if no_rsi else fetch_rsi(config)

    try:
        added, skipped = update_workbook(
            excel_path, data,
            futures_price=futures_price,
            futures_expiry=futures_expiry,
            rsi=rsi_map,
            refresh_recent=refresh_recent,
        )
    except PermissionError as exc:
        return FetchResult(error=str(exc), excel_path=excel_path)

    logger.info("=" * 60)
    logger.info("Update complete!  New rows: %d  |  Already in sheet: %d", added, skipped)
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
        no_rsi=args.no_rsi,
        refresh_recent=args.refresh_recent,
    )

    if not result.ok:
        logger.error(result.error)
        sys.exit(1)


def _empty_ohlc():
    import pandas as pd
    return pd.DataFrame(columns=["Open", "High", "Low", "Close"])


def _open_excel(path: str) -> None:
    """Open the Excel file in the default application."""
    try:
        target = os.path.abspath(path)
        if sys.platform == "darwin":
            subprocess.run(["open", target], check=True)
        elif sys.platform.startswith("win"):
            os.startfile(target)  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", target], check=True)
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
