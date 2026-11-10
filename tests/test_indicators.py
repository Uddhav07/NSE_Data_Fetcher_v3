"""Pure calculations: RSI, 45-minute candles, intraday filtering, Moon/nakshatra, NSE session."""
from __future__ import annotations

from datetime import date

import pandas as pd
import pytest

import scripts.market as market
import scripts.rsi as rsi_mod
from scripts.moon_calc import _julian_day, moon_tropical_longitude, sidereal_moon_longitude
from scripts.nakshatra import NAKSHATRAS, moon_nakshatra
from scripts.rsi import RSI_PERIOD, _to_45m, compute_rsi_map, wilder_rsi

IST = "Asia/Kolkata"


# ── RSI ───────────────────────────────────────────────────────────────────────

def _reference_rsi(closes, period=RSI_PERIOD):
    """Independent loop implementation of the same Wilder smoothing (seeded with the first change)."""
    a = 1 / period
    out, g_avg, l_avg = [None], None, None
    for i in range(1, len(closes)):
        ch = closes[i] - closes[i - 1]
        g, lo = max(ch, 0.0), max(-ch, 0.0)
        g_avg = g if g_avg is None else (1 - a) * g_avg + a * g
        l_avg = lo if l_avg is None else (1 - a) * l_avg + a * lo
        if i < period:
            out.append(None)
        elif l_avg == 0:
            out.append(100.0)
        else:
            out.append(100 - 100 / (1 + g_avg / l_avg))
    return out


def test_wilder_rsi_matches_reference_loop():
    closes = [100, 101.5, 100.2, 102.8, 103.1, 101.9, 104.4, 105.0, 103.3, 106.1,
              107.2, 105.8, 104.9, 106.6, 108.0, 107.1, 109.3, 108.2, 110.5, 109.9, 108.4]
    got = wilder_rsi(pd.Series(closes, dtype=float)).tolist()
    want = _reference_rsi(closes)
    for i, (g, w) in enumerate(zip(got, want)):
        if w is None:
            assert pd.isna(g), i                        # first RSI_PERIOD values are blank
        else:
            assert g == pytest.approx(w, abs=1e-9), i
    assert pd.notna(got[RSI_PERIOD])


def test_wilder_rsi_extremes():
    up = pd.Series([float(x) for x in range(30)])
    down = pd.Series([float(x) for x in range(30, 0, -1)])
    assert wilder_rsi(up).iloc[-1] == 100.0
    assert wilder_rsi(down).iloc[-1] == pytest.approx(0.0)


def _day_15m(day: str, n: int = 25, start: float = 100.0) -> pd.Series:
    """One NSE session of 15m candles: 09:15 … 15:15 (25 candles)."""
    idx = pd.date_range(f"{day} 09:15", periods=n, freq="15min", tz=IST)
    return pd.Series([start + i * 0.5 + (i % 4) for i in range(n)], index=idx)


def test_45m_candles_anchor_at_0915_and_end_with_last_15m_close():
    s = _day_15m("2026-01-05")
    m45 = _to_45m(s)
    assert [t.strftime("%H:%M") for t in m45.index[:3]] == ["09:15", "10:00", "10:45"]
    assert len(m45) == 9                               # 8 full buckets + 15:00 bucket (15:00, 15:15)
    assert m45.iloc[-1] == s.iloc[-1]                  # 45m close = last 15m close in the bucket
    assert m45.iloc[0] == s.iloc[2]


def test_45m_buckets_never_cross_days():
    s = pd.concat([_day_15m("2026-01-05"), _day_15m("2026-01-06", start=200)])
    m45 = _to_45m(s)
    assert {t.strftime("%Y-%m-%d") for t in m45.index} == {"2026-01-05", "2026-01-06"}
    assert m45[m45.index.strftime("%Y-%m-%d") == "2026-01-06"].index[0].strftime("%H:%M") == "09:15"


def test_rsi_map_is_last_candle_of_each_day():
    s = pd.concat([_day_15m("2026-01-05"), _day_15m("2026-01-06", start=110)])
    m = compute_rsi_map({15: s})
    full = wilder_rsi(s)
    assert m["2026-01-06"][15] == round(float(full.iloc[-1]), 2)
    assert set(m["2026-01-06"]) == {15, 45}            # 45m derived from 15m


def test_rsi_map_skips_series_too_short():
    s = _day_15m("2026-01-05", n=RSI_PERIOD)           # not more than RSI_PERIOD candles
    assert compute_rsi_map({5: s}) == {}


class _FakeTicker:
    def __init__(self, df):
        self._df = df

    def history(self, **kw):
        return self._df


def test_download_drops_candles_from_1530(monkeypatch):
    idx = pd.DatetimeIndex(["2026-01-05 15:15", "2026-01-05 15:25", "2026-01-05 15:30"]).tz_localize(IST)
    df = pd.DataFrame({"Close": [1.0, 2.0, 3.0]}, index=idx)
    monkeypatch.setattr(rsi_mod.yf, "Ticker", lambda t: _FakeTicker(df))
    from scripts.config_manager import AppConfig
    s = rsi_mod._download(AppConfig(max_retries=1), "5m")
    assert s.tolist() == [1.0, 2.0]


def test_fetch_rsi_drops_today_while_session_open(monkeypatch):
    s = pd.concat([_day_15m("2026-01-05"), _day_15m("2026-01-06", start=110)])
    monkeypatch.setattr(rsi_mod, "_download", lambda config, interval, period="60d": s)
    monkeypatch.setattr(rsi_mod, "today_str", lambda: "2026-01-06")
    from scripts.config_manager import AppConfig
    monkeypatch.setattr(rsi_mod, "session_in_progress", lambda: True)
    assert "2026-01-06" not in rsi_mod.fetch_rsi(AppConfig(max_retries=1))
    monkeypatch.setattr(rsi_mod, "session_in_progress", lambda: False)
    got = rsi_mod.fetch_rsi(AppConfig(max_retries=1))
    assert set(got["2026-01-06"]) == {5, 15, 30, 45, 60}


# ── Moon / nakshatra (references: Swiss Ephemeris 2.10, Lahiri, 03:45 UT; Meeus ex. 47.a) ──

SWISS_EPHEMERIS = {   # date -> sidereal Moon longitude at 09:15 IST
    (2020, 2, 29): 15.9242, (2024, 1, 1): 133.6514, (2025, 3, 14): 148.1407,
    (2026, 1, 6): 118.2479, (2026, 11, 9): 200.9559, (2030, 8, 15): 320.0078,
}


def test_julian_day_reference_points():
    assert _julian_day(2000, 1, 1, 12.0) == 2451545.0             # J2000.0
    assert _julian_day(1957, 10, 4, 0.81 * 24) == pytest.approx(2436116.31)   # Meeus ex. 7.a


def test_moon_longitude_meeus_example_47a():
    jd_ut = _julian_day(1992, 4, 12, 0.0) - 69.0 / 86400.0         # 1992-04-12 0h TD
    assert moon_tropical_longitude(jd_ut) == pytest.approx(133.162655, abs=0.01)


@pytest.mark.parametrize("ymd,lon", sorted(SWISS_EPHEMERIS.items()))
def test_sidereal_moon_matches_swiss_ephemeris(ymd, lon):
    assert sidereal_moon_longitude(date(*ymd)) == pytest.approx(lon, abs=0.01)


@pytest.mark.parametrize("ymd,name", [
    ((2020, 2, 29), "Bharani"), ((2024, 1, 1), "Purva Phalguni"), ((2025, 3, 14), "Uttara Phalguni"),
    ((2026, 1, 6), "Ashlesha"), ((2026, 11, 9), "Vishakha"),       # all ≥0.3° from a boundary
])
def test_nakshatra_names(ymd, name):
    assert moon_nakshatra(date(*ymd)) == name


def test_nakshatra_table():
    assert len(NAKSHATRAS) == 27 and len(set(NAKSHATRAS)) == 27
    assert NAKSHATRAS[0] == "Ashwini" and NAKSHATRAS[-1] == "Revati"


# ── NSE session (IST) ──────────────────────────────────────────────────────────

@pytest.mark.parametrize("ts,live", [
    ("2026-01-05 09:00", True),      # Monday before open: today's data not final
    ("2026-01-05 15:34", True),
    ("2026-01-05 15:35", False),     # settled
    ("2026-01-05 23:59", False),
    ("2026-01-10 11:00", False),     # Saturday
    ("2026-01-11 11:00", False),     # Sunday
])
def test_session_in_progress(monkeypatch, ts, live):
    monkeypatch.setattr(market, "now_ist", lambda: pd.Timestamp(ts, tz=IST))
    assert market.session_in_progress() is live


def test_today_str_uses_ist(monkeypatch):
    monkeypatch.setattr(market, "now_ist", lambda: pd.Timestamp("2026-01-06 00:30", tz=IST))
    assert market.today_str() == "2026-01-06"
