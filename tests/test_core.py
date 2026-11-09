from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from scripts.excel_writer import create_workbook, update_workbook, get_last_date
from scripts.futures import _parse_expiry_from_symbol
from scripts.config_manager import AppConfig


def _df(dates):
    idx = pd.DatetimeIndex(dates, name="Date")
    n = len(idx)
    return pd.DataFrame({"Open": range(100, 100 + n), "High": range(110, 110 + n),
                         "Low": range(90, 90 + n), "Close": range(105, 105 + n)}, index=idx)


def test_expiry():
    assert _parse_expiry_from_symbol("NIFTY25NOVFUT") == "NIFTY Nov 2025 Fut"
    assert _parse_expiry_from_symbol("garbage") is None


def test_config_validate():
    assert AppConfig().validate() == []
    assert AppConfig(start_date="bad").validate()


def test_append_dupes_weekly_and_futures(tmp_path):
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    # 7 weekdays with a gap (holiday on 2026-01-07)
    dates = ["2026-01-01", "2026-01-02", "2026-01-05", "2026-01-06", "2026-01-08", "2026-01-09", "2026-01-12"]
    assert update_workbook(p, _df(dates), 25000.0, "NIFTY Jan 2026 Fut") == (7, 0)
    assert get_last_date(p) == datetime(2026, 1, 12)
    ws = load_workbook(p).active
    assert ws["H6"].value is None and ws["H7"].value == "=F7-F2"
    assert ws["H8"].value == "=F8-F3"       # 5 trading rows back, skipping the holiday gap
    assert ws["N8"].value == 25000.0 and ws["M8"].value == "=N8-F8"
    assert ws["N7"].value is None
    # duplicates skipped
    assert update_workbook(p, _df(dates)) == (0, 7)
    # futures-only update stamps the last dated row, even with a stray lower row
    ws = load_workbook(p); w = ws.active; w["P20"] = "note"; ws.save(p)
    assert get_last_date(p) == datetime(2026, 1, 12)
    update_workbook(p, _df([]), 25100.0, "NIFTY Jan 2026 Fut")
    assert load_workbook(p).active["N8"].value == 25000.0   # existing futures never overwritten


def test_date_total_and_rsi_columns(tmp_path):
    from scripts.rsi import compute_rsi_map
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    rsi = {"2026-11-09": {5: 55.5, 15: 60.1, 30: 48.0, 45: 51.2}}
    update_workbook(p, _df(["2026-11-09"]), rsi=rsi)
    ws = load_workbook(p).active
    assert ws["Q1"].value.startswith("Date") and ws["U1"].value.startswith("RSI")
    assert ws["Q2"].value == '=DAY(B2)&", "&(1+MOD(YEAR(B2)*10000+MONTH(B2)*100+DAY(B2)-1,9))'
    # digital root of 20261109 -> 2+0+2+6+1+1+0+9 = 21 -> 3
    assert 1 + (20261109 - 1) % 9 == 3
    assert [ws.cell(2, c).value for c in (18, 19, 20, 21)] == [55.5, 60.1, 48.0, 51.2]

    # synthetic 15m closes for 3 days -> every interval produces a last-of-day value
    idx = pd.date_range("2026-11-05 09:15", periods=25, freq="15min", tz="Asia/Kolkata")
    idx = idx.append(pd.date_range("2026-11-06 09:15", periods=25, freq="15min", tz="Asia/Kolkata"))
    close = pd.Series([100 + (i % 7) + i * 0.1 for i in range(len(idx))], index=idx)
    m = compute_rsi_map({15: close})
    assert set(m["2026-11-06"]) == {15, 45}
    assert 0 <= m["2026-11-06"][15] <= 100


def test_user_edits_survive_update(tmp_path):
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    update_workbook(p, _df(["2026-01-05", "2026-01-06"]))
    wb = load_workbook(p); ws = wb.active
    ws["P2"] = "my note"; ws["F2"] = 999.0; ws["R2"] = 1.0
    ws["A3"].fill = __import__("openpyxl").styles.PatternFill("solid", start_color="FF0000", end_color="FF0000")
    ws["Q3"] = "custom"
    ws["H1"] = "Wk chg (mine)"
    wb.save(p)
    # new run: one new date + RSI for existing dates
    update_workbook(p, _df(["2026-01-07"]), 25000.0, "X",
                    rsi={"2026-01-05": {5: 70.0, 15: 71.0, 30: 72.0, 45: 73.0}})
    ws = load_workbook(p).active
    assert ws["P2"].value == "my note" and ws["F2"].value == 999.0
    assert ws["R2"].value == 1.0 and ws["S2"].value == 71.0   # blank filled, filled kept
    assert ws["A3"].fill.start_color.rgb.endswith("FF0000")   # Tuesday row keeps user fill
    assert ws["Q3"].value == "custom" and ws["H1"].value == "Wk chg (mine)"
    assert ws.max_row == 4


def test_nakshatra_and_cell_only_rules(tmp_path):
    from scripts.nakshatra import moon_nakshatra
    from datetime import date
    assert moon_nakshatra(date(2024, 1, 1)) == "Purva Phalguni"
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    update_workbook(p, _df(["2026-01-06", "2026-01-07"]))
    ws = load_workbook(p).active
    assert ws["V1"].value.startswith("Moon") and ws["V2"].value and ws["V3"].value
    rules = {str(cf.sqref): [r.formula for r in cf.rules] for cf in ws.conditional_formatting}
    assert rules["R2:U1048576"] == [["AND(ISNUMBER(R2),R2<30)"], ["AND(ISNUMBER(R2),R2>=80)"], ["AND(ISNUMBER(R2),R2>=70,R2<80)"]]
    assert rules["Q2:Q1048576"] == [['OR(Q2="3, 6",Q2="6, 3")']]
    n = len(ws.conditional_formatting)
    update_workbook(p, _df([]))                        # rerun: no duplicate rules
    assert len(load_workbook(p).active.conditional_formatting) == n
