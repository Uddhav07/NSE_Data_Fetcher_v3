"""Intraday RSI (5/15/30/45-minute candles) from Yahoo Finance.

For each trading date we report the RSI of the **last candle of the day**.
Yahoo only serves ~60 days of 5m/15m/30m history, so older dates get no value.
45-minute candles are built from 15-minute candles (Yahoo has no native 45m).
"""
from __future__ import annotations

import logging
import time
from typing import Optional

import pandas as pd
import yfinance as yf

from .config_manager import AppConfig

logger = logging.getLogger(__name__)

RSI_PERIOD = 14
INTERVALS = (5, 15, 30, 45)

# date "YYYY-MM-DD" -> {minutes: rsi}
RsiMap = dict


def wilder_rsi(close: pd.Series, period: int = RSI_PERIOD) -> pd.Series:
    """Wilder's RSI (EMA with alpha=1/period)."""
    delta = close.diff()
    gain = delta.clip(lower=0).ewm(alpha=1 / period, adjust=False, min_periods=period).mean()
    loss = (-delta.clip(upper=0)).ewm(alpha=1 / period, adjust=False, min_periods=period).mean()
    rs = gain / loss
    rsi = 100 - 100 / (1 + rs)
    rsi = rsi.where(loss != 0, 100.0)  # no losses -> RSI 100
    return rsi


def _to_45m(close_15m: pd.Series) -> pd.Series:
    """Build 45-minute closes per trading day, anchored at the first candle (09:15)."""
    parts = []
    for _, day in close_15m.groupby(close_15m.index.date):
        parts.append(day.resample("45min", origin="start").last().dropna())
    return pd.concat(parts) if parts else close_15m.iloc[0:0]


def _last_rsi_per_day(close: pd.Series) -> dict:
    rsi = wilder_rsi(close).dropna()
    out = {}
    for d, s in rsi.groupby(rsi.index.date):
        out[d.strftime("%Y-%m-%d")] = round(float(s.iloc[-1]), 2)
    return out


def compute_rsi_map(closes: dict) -> RsiMap:
    """closes: {5: Series, 15: Series, 30: Series} of intraday closes. Adds 45m."""
    series = dict(closes)
    if 15 in closes and len(closes[15]):
        series[45] = _to_45m(closes[15])
    result: RsiMap = {}
    for minutes, s in series.items():
        if s is None or len(s) <= RSI_PERIOD:
            continue
        for date_str, val in _last_rsi_per_day(s).items():
            result.setdefault(date_str, {})[minutes] = val
    return result


def _download(config: AppConfig, interval: str) -> Optional[pd.Series]:
    for attempt in range(1, config.max_retries + 1):
        try:
            df = yf.Ticker(config.ticker).history(period="60d", interval=interval)
            if df.empty:
                return None
            close = df["Close"].dropna()
            # Only candles that start before the 15:30 IST close (ends the session at 15:30)
            mins = close.index.hour * 60 + close.index.minute
            return close[mins < 15 * 60 + 30]
        except Exception as exc:  # network / yahoo errors
            logger.warning("RSI %s fetch attempt %d failed: %s", interval, attempt, exc)
            if attempt < config.max_retries:
                time.sleep(2 ** attempt)
    return None


def fetch_rsi(config: AppConfig) -> RsiMap:
    """Download 5/15/30m candles and compute last-candle-of-day RSI. {} on failure."""
    closes = {}
    for minutes in (5, 15, 30):
        logger.info("Fetching %dm candles for RSI...", minutes)
        s = _download(config, f"{minutes}m")
        if s is not None:
            closes[minutes] = s
    if not closes:
        logger.warning("No intraday data available — RSI columns left blank.")
        return {}
    result = compute_rsi_map(closes)
    # Don't record today's RSI until the session is over (candles still forming)
    now_ist = pd.Timestamp.now(tz="Asia/Kolkata")
    if now_ist.hour * 60 + now_ist.minute < 15 * 60 + 35:
        if result.pop(now_ist.strftime("%Y-%m-%d"), None) is not None:
            logger.info("Market still open — today's RSI skipped; re-run after 15:35 IST.")
    logger.info("RSI computed for %d trading day(s).", len(result))
    return result
