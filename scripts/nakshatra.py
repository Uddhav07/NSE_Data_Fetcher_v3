"""Moon nakshatra (sidereal, Lahiri ayanamsa) at 09:15 IST on a given date."""
from __future__ import annotations

from datetime import date, datetime

import swisseph as swe

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
    swe.set_sid_mode(swe.SIDM_LAHIRI)
    jd = swe.julday(day.year, day.month, day.day, 3.75)  # 09:15 IST = 03:45 UT
    lon, _ = swe.calc_ut(jd, swe.MOON, swe.FLG_SWIEPH | swe.FLG_SIDEREAL)[:2]
    return NAKSHATRAS[int((lon[0] if hasattr(lon, "__len__") else lon) // _SPAN) % 27]
