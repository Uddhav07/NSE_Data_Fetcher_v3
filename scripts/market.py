"""NSE trading-session helpers (IST)."""
from __future__ import annotations

import pandas as pd

SESSION_END_MINUTES = 15 * 60 + 35  # 15:35 IST: session closed and final candles settled


def now_ist() -> pd.Timestamp:
    return pd.Timestamp.now(tz="Asia/Kolkata")


def session_in_progress() -> bool:
    """True on a weekday before 15:35 IST — today's candles/futures are not final yet."""
    n = now_ist()
    return n.weekday() < 5 and n.hour * 60 + n.minute < SESSION_END_MINUTES


def today_str() -> str:
    return now_ist().strftime("%Y-%m-%d")
