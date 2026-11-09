"""Pure-Python Moon position (Meeus, 'Astronomical Algorithms' ch. 47, truncated) + Lahiri ayanamsa.

No third-party packages, so the Nakshatra column works on every platform.
Accuracy ~0.01 degree for the Moon, far finer than a nakshatra (13.33 degrees).
"""
from __future__ import annotations

import math
from datetime import date, datetime

# (D, M, M', F, sum_l in 1e-6 degree)
_TERMS = [
    (0, 0, 1, 0, 6288774), (2, 0, -1, 0, 1274027), (2, 0, 0, 0, 658314), (0, 0, 2, 0, 213618),
    (0, 1, 0, 0, -185116), (0, 0, 0, 2, -114332), (2, 0, -2, 0, 58793), (2, -1, -1, 0, 57066),
    (2, 0, 1, 0, 53322), (2, -1, 0, 0, 45758), (0, 1, -1, 0, -40923), (1, 0, 0, 0, -34720),
    (0, 1, 1, 0, -30383), (2, 0, 0, -2, 15327), (0, 0, 1, 2, -12528), (0, 0, 1, -2, 10980),
    (4, 0, -1, 0, 10675), (0, 0, 3, 0, 10034), (4, 0, -2, 0, 8548), (2, 1, -1, 0, -7888),
    (2, 1, 0, 0, -6766), (1, 0, -1, 0, -5163), (1, 1, 0, 0, 4987), (2, -1, 1, 0, 4036),
    (2, 0, 2, 0, 3994), (4, 0, 0, 0, 3861), (2, 0, -3, 0, 3665), (0, 1, -2, 0, -2689),
    (2, 0, -1, 2, -2602), (2, -1, -2, 0, 2390), (1, 0, 1, 0, -2348), (2, -2, 0, 0, 2236),
    (0, 1, 2, 0, -2120), (0, 2, 0, 0, -2069), (2, -2, -1, 0, 2048), (2, 0, 1, -2, -1773),
    (2, 0, 0, 2, -1595), (4, -1, -1, 0, 1215), (0, 0, 2, 2, -1110), (3, 0, -1, 0, -892),
    (2, 1, 1, 0, -810), (4, -1, -2, 0, 759), (0, 2, -1, 0, -713), (2, 2, -1, 0, -700),
    (2, 1, -2, 0, 691), (2, -1, 0, -2, 596), (4, 0, 1, 0, 549), (0, 0, 4, 0, 537),
    (4, -1, 0, 0, 520), (1, 0, -2, 0, -487), (2, 1, 0, -2, -399), (0, 0, 2, -2, -381),
    (1, 1, 1, 0, 351), (3, 0, -2, 0, -340), (4, 0, -3, 0, 330), (2, -1, 2, 0, 327),
    (0, 2, 1, 0, -323), (1, 1, -1, 0, 299), (2, 0, 3, 0, 294),
]


def _julian_day(y: int, m: int, d: int, hour_ut: float) -> float:
    if m <= 2:
        y -= 1
        m += 12
    a = y // 100
    b = 2 - a + a // 4
    return int(365.25 * (y + 4716)) + int(30.6001 * (m + 1)) + d + b - 1524.5 + hour_ut / 24.0


def moon_tropical_longitude(jd_ut: float) -> float:
    """Apparent geocentric ecliptic longitude of the Moon (degrees, of date)."""
    jd_tt = jd_ut + 69.0 / 86400.0
    t = (jd_tt - 2451545.0) / 36525.0
    r = math.radians
    lp = (218.3164477 + 481267.88123421 * t - 0.0015786 * t**2 + t**3 / 538841 - t**4 / 65194000)
    dd = (297.8501921 + 445267.1114034 * t - 0.0018819 * t**2 + t**3 / 545868 - t**4 / 113065000)
    m = 357.5291092 + 35999.0502909 * t - 0.0001536 * t**2 + t**3 / 24490000
    mp = (134.9633964 + 477198.8675055 * t + 0.0087414 * t**2 + t**3 / 69699 - t**4 / 14712000)
    f = (93.2720950 + 483202.0175233 * t - 0.0036539 * t**2 - t**3 / 3526000 + t**4 / 863310000)
    a1 = 119.75 + 131.849 * t
    a2 = 53.09 + 479264.290 * t
    e = 1 - 0.002516 * t - 0.0000074 * t**2
    sl = 0.0
    for cd, cm, cmp_, cf, coef in _TERMS:
        arg = r(cd * dd + cm * m + cmp_ * mp + cf * f)
        sl += coef * (e ** abs(cm)) * math.sin(arg)
    sl += 3958 * math.sin(r(a1)) + 1962 * math.sin(r(lp - f)) + 318 * math.sin(r(a2))
    lon = lp + sl / 1e6
    omega = 125.04452 - 1934.136261 * t          # nutation in longitude (dominant term)
    lon += (-17.20 * math.sin(r(omega))) / 3600.0
    return lon % 360.0


def lahiri_ayanamsa(jd_ut: float) -> float:
    """Lahiri (Chitrapaksha) ayanamsa in degrees."""
    t = (jd_ut - 2451545.0) / 36525.0
    return 23.857092 + (5029.0966 * t + 1.11113 * t * t) / 3600.0


def sidereal_moon_longitude(day: date | datetime, hour_ut: float = 3.75) -> float:
    jd = _julian_day(day.year, day.month, day.day, hour_ut)
    return (moon_tropical_longitude(jd) - lahiri_ayanamsa(jd)) % 360.0
