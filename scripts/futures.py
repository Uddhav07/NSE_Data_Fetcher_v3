"""Nifty futures data scraper with retry and fallback."""

import json
import logging
import re
import time
from datetime import datetime
from typing import Optional

import requests

from .config_manager import AppConfig

logger = logging.getLogger(__name__)

# Type alias for the return value: (price, expiry_label)
FuturesResult = tuple[Optional[float], Optional[str]]

_MONTH_MAP = {
    "JAN": "Jan", "FEB": "Feb", "MAR": "Mar", "APR": "Apr",
    "MAY": "May", "JUN": "Jun", "JUL": "Jul", "AUG": "Aug",
    "SEP": "Sep", "OCT": "Oct", "NOV": "Nov", "DEC": "Dec",
}

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}


def _parse_expiry_from_symbol(symbol: str) -> Optional[str]:
    """Parse expiry info from a Groww futures symbol like NIFTY25NOVFUT."""
    for code, name in _MONTH_MAP.items():
        if code in symbol:
            match = re.search(rf"NIFTY(\d{{2}}){code}", symbol)
            if match:
                year = "20" + match.group(1)
                return f"NIFTY {name} {year} Fut"
    return None


def _fetch_from_groww(config: AppConfig) -> FuturesResult:
    """Scrape current Nifty futures price and expiry from Groww."""
    url = "https://groww.in/futures/nifty"
    try:
        resp = requests.get(url, headers=_HEADERS, timeout=config.request_timeout)
        resp.raise_for_status()
    except requests.RequestException as exc:
        logger.warning("Groww request failed: %s", exc)
        return None, None

    # Extract __NEXT_DATA__ JSON blob
    match = re.search(
        r'<script id="__NEXT_DATA__"[^>]*>({.*?})</script>',
        resp.text,
        re.DOTALL,
    )
    if not match:
        logger.warning("Could not locate __NEXT_DATA__ in Groww response.")
        return None, None

    try:
        data = json.loads(match.group(1))
    except json.JSONDecodeError:
        logger.warning("Failed to parse Groww __NEXT_DATA__ JSON.")
        return None, None

    try:
        contract = data["props"]["pageProps"]["contractData"]
        live = contract["livePrice"]
        price = live.get("ltp")
        symbol = live.get("symbol", "")
    except (KeyError, TypeError):
        logger.warning("Unexpected Groww data structure.")
        return None, None

    # Derive expiry label
    expiry = _parse_expiry_from_symbol(symbol)
    if expiry is None:
        # Fallback: look in contractDetails
        try:
            raw = contract["contractDetails"]["expiry"]
            dt = datetime.strptime(raw, "%Y-%m-%d")
            expiry = f"NIFTY {dt.strftime('%b')} {dt.year} Fut"
        except (KeyError, TypeError, ValueError):
            expiry = None

    return price, expiry


def fetch_futures(config: AppConfig) -> FuturesResult:
    """Fetch current Nifty futures price and expiry with retries.

    Returns (price, expiry_label) or (None, None) on failure.
    """
    for attempt in range(1, config.max_retries + 1):
        logger.info("Fetching futures data (attempt %d/%d)...", attempt, config.max_retries)
        price, expiry = _fetch_from_groww(config)
        if price is not None:
            logger.info("Futures: %.2f — %s", price, expiry or "unknown expiry")
            return price, expiry
        if attempt < config.max_retries:
            wait = 2 ** attempt
            logger.info("Retrying futures fetch in %d seconds...", wait)
            time.sleep(wait)

    logger.warning("Could not fetch futures data after %d attempts.", config.max_retries)
    return None, None
