"""Yahoo Finance data fetcher with retry logic and validation."""

import logging
import time
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import yfinance as yf

from .config_manager import AppConfig

logger = logging.getLogger(__name__)


def fetch_ohlc_data(
    config: AppConfig,
    start_date: str | datetime,
    end_date: Optional[str | datetime] = None,
) -> Optional[pd.DataFrame]:
    """Fetch OHLC data from Yahoo Finance with retries.

    Returns a DataFrame with columns [Open, High, Low, Close] indexed by Date,
    or None if no data could be fetched.
    """
    if end_date is None:
        end_date = datetime.now()

    for attempt in range(1, config.max_retries + 1):
        try:
            logger.info(
                "Fetching %s data from %s to %s (attempt %d/%d)",
                config.ticker, start_date, end_date, attempt, config.max_retries,
            )
            ticker = yf.Ticker(config.ticker)
            data = ticker.history(start=start_date, end=end_date)

            if data.empty:
                logger.info("No data returned — market may be closed or no trading occurred.")
                return None

            # Keep only OHLC columns
            data = data[["Open", "High", "Low", "Close"]].copy()
            data.index.name = "Date"

            # Validate: drop rows where all values are NaN or zero
            data = data.dropna(how="all")
            data = data[(data != 0).any(axis=1)]

            if data.empty:
                logger.warning("All fetched rows were invalid (NaN/zero). No usable data.")
                return None

            # Round to 2 decimal places
            data = data.round(2)

            logger.info("Successfully fetched %d records.", len(data))
            return data

        except Exception as exc:
            logger.warning("Attempt %d failed: %s", attempt, exc)
            if attempt < config.max_retries:
                wait = 2 ** attempt
                logger.info("Retrying in %d seconds...", wait)
                time.sleep(wait)

    logger.error("All %d fetch attempts failed for %s.", config.max_retries, config.ticker)
    return None
