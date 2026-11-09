"""Moon nakshatra (sidereal, Lahiri ayanamsa) at 09:15 IST on a given date.

Pure Python (see moon_calc.py) — no extra packages, identical on every OS.
"""
from __future__ import annotations

from datetime import date, datetime

from .moon_calc import sidereal_moon_longitude

NAKSHATRAS = [
    "Ashwini", "Bharani", "Krittika", "Rohini", "Mrigashira", "Ardra",
    "Punarvasu", "Pushya", "Ashlesha", "Magha", "Purva Phalguni",
    "Uttara Phalguni", "Hasta", "Chitra", "Swati", "Vishakha", "Anuradha",
    "Jyeshtha", "Mula", "Purva Ashadha", "Uttara Ashadha", "Shravana",
    "Dhanishta", "Shatabhisha", "Purva Bhadrapada", "Uttara Bhadrapada",
    "Revati",
]
_SPAN = 360.0 / 27  # 13°20'


def moon_nakshatra(day: date | datetime) -> str:
    """Nakshatra occupied by the Moon at 09:15 IST (03:45 UT) on `day`."""
    return NAKSHATRAS[int(sidereal_moon_longitude(day, 3.75) // _SPAN) % 27]
