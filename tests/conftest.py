"""Shared test helpers. Tests never touch the network: fetchers are mocked."""
from __future__ import annotations

import pandas as pd
import pytest


def make_ohlc(dates, base: float = 25000.0) -> pd.DataFrame:
    """Deterministic OHLC frame with distinct float values per row (Yahoo-like index name)."""
    idx = pd.DatetimeIndex(dates, name="Date")
    rows = []
    for i in range(len(idx)):
        o = base + i * 37.25
        rows.append({"Open": o, "High": o + 120.5, "Low": o - 95.75, "Close": o + (i % 3 - 1) * 42.1})
    return pd.DataFrame(rows, index=idx, columns=["Open", "High", "Low", "Close"])


@pytest.fixture
def ohlc():
    return make_ohlc
