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
    assert ws["U8"].value == 25000.0 and ws["M8"].value == "=U8-F8"
    assert ws["U7"].value is None
    # duplicates skipped
    assert update_workbook(p, _df(dates)) == (0, 7)
    # futures-only update stamps the last dated row, even with a stray lower row
    ws = load_workbook(p); w = ws.active; w["W20"] = "note"; ws.save(p)
    assert get_last_date(p) == datetime(2026, 1, 12)
    update_workbook(p, _df([]), 25100.0, "NIFTY Jan 2026 Fut")
    assert load_workbook(p).active["U8"].value == 25000.0   # existing futures never overwritten


def test_date_total_and_rsi_columns(tmp_path):
    from scripts.rsi import compute_rsi_map
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    rsi = {"2026-11-09": {5: 55.5, 15: 60.1, 30: 48.0, 45: 51.2, 60: 66.6}}
    update_workbook(p, _df(["2026-11-09"]), rsi=rsi)
    ws = load_workbook(p).active
    assert ws["N1"].value.startswith("Date") and ws["R1"].value.startswith("RSI")
    assert ws["N2"].value == '=DAY(B2)&", "&(1+MOD(YEAR(B2)*10000+MONTH(B2)*100+DAY(B2)-1,9))'
    # digital root of 20261109 -> 2+0+2+6+1+1+0+9 = 21 -> 3
    assert 1 + (20261109 - 1) % 9 == 3
    assert [ws.cell(2, c).value for c in (15, 16, 17, 18, 19)] == [55.5, 60.1, 48.0, 51.2, 66.6]

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
    ws["W2"] = "my note"; ws["F2"] = 999.0; ws["O2"] = 1.0
    ws["A3"].fill = __import__("openpyxl").styles.PatternFill("solid", start_color="FF0000", end_color="FF0000")
    ws["N3"] = "custom"
    ws["H1"] = "Wk chg (mine)"
    wb.save(p)
    # new run: one new date + RSI for existing dates
    update_workbook(p, _df(["2026-01-07"]), 25000.0, "X",
                    rsi={"2026-01-05": {5: 70.0, 15: 71.0, 30: 72.0, 45: 73.0}})
    ws = load_workbook(p).active
    assert ws["W2"].value == "my note" and ws["F2"].value == 999.0
    assert ws["O2"].value == 1.0 and ws["P2"].value == 71.0   # blank filled, filled kept
    assert ws["A3"].fill.start_color.rgb.endswith("FF0000")   # Tuesday row keeps user fill
    assert ws["N3"].value == "custom" and ws["H1"].value == "Wk chg (mine)"
    assert ws.max_row == 4


def test_nakshatra_and_cell_only_rules(tmp_path):
    from scripts.nakshatra import moon_nakshatra
    from datetime import date
    assert moon_nakshatra(date(2024, 1, 1)) == "Purva Phalguni"
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    update_workbook(p, _df(["2026-01-06", "2026-01-07"]))
    ws = load_workbook(p).active
    assert ws["T1"].value.startswith("Moon") and ws["T2"].value and ws["T3"].value
    rules = {str(cf.sqref): [r.formula for r in cf.rules] for cf in ws.conditional_formatting}
    assert rules["O2:S1048576"] == [["AND(ISNUMBER(O2),O2<30)"], ["AND(ISNUMBER(O2),O2>=80)"], ["AND(ISNUMBER(O2),O2>=70,O2<80)"]]
    assert rules["N2:N1048576"] == [['OR(N2="3, 6",N2="6, 3")']]
    n = len(ws.conditional_formatting)
    update_workbook(p, _df([]))                        # rerun: no duplicate rules
    assert len(load_workbook(p).active.conditional_formatting) == n


def test_gap_backfill_inserts_in_order_and_keeps_edits(tmp_path):
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    d = ["2026-01-02", "2026-01-05", "2026-01-06", "2026-01-07", "2026-01-08",
         "2026-01-09", "2026-01-12", "2026-01-13"]
    # first run: everything except 01-06 and 01-09 (gaps), and 01-02 is also missing
    have = [x for x in d if x not in ("2026-01-06", "2026-01-09", "2026-01-02")]
    update_workbook(p, _df(have))
    wb = load_workbook(p); ws = wb.active
    ws["W3"] = "my note on 01-07"          # user note on an existing row
    ws["F3"] = 12345.0                     # user edited value on an existing row
    wb.save(p)

    # second run: full data incl. the missing dates (earlier, middle) -> backfill
    full = _df(d)
    added, skipped = update_workbook(p, full)
    ws = load_workbook(p).active
    dates = [ws.cell(r, 2).value.strftime("%Y-%m-%d") for r in range(2, ws.max_row + 1)]
    assert dates == d and added == 3
    # edits moved with their rows and were not overwritten
    r = dates.index("2026-01-07") + 2
    assert ws.cell(r, 23).value == "my note on 01-07" and ws.cell(r, 6).value == 12345.0
    # every formula points at its own row (no stale references after insertion)
    for row in range(2, ws.max_row + 1):
        assert ws.cell(row, 1).value == f'=TEXT(B{row},"dddd")'
        if row > 2:
            assert ws.cell(row, 7).value == f"=F{row}-F{row - 1}"
        if row > 6:
            assert ws.cell(row, 8).value == f"=F{row}-F{row - 5}"
    # idempotent: running again changes nothing
    snap = [[c.value for c in row] for row in ws.iter_rows()]
    assert update_workbook(p, full)[0] == 0
    ws2 = load_workbook(p).active
    assert snap == [[c.value for c in row] for row in ws2.iter_rows()]


def test_bootstrap_detects_missing_and_ok():
    from scripts import bootstrap
    assert bootstrap._missing(("pandas", "definitely_not_a_pkg")) == ["definitely_not_a_pkg"]
    assert bootstrap._missing(("os",)) == []


def _build_c(tmp_path):
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    rsi = {"2026-01-05": {5: 61.0, 15: 62.0, 30: 63.0, 45: 64.0, 60: 65.0}}
    update_workbook(p, _df(["2026-01-05", "2026-01-06"]), 25000.0, "NIFTY Jan 2026 Fut", rsi=rsi)
    wb = load_workbook(p); ws = wb.active
    ws["W2"] = "my note"; ws["W3"] = "second note"
    from openpyxl.comments import Comment
    ws["U3"].comment = Comment("hello", "me")
    wb.save(p)
    return p


def _check_c(p):
    ws = load_workbook(p).active
    heads = [ws.cell(1, c).value for c in range(1, 24)]
    assert heads[14:19] == ["RSI 5m", "RSI 15m", "RSI 30m", "RSI 45m", "RSI 60m"]
    assert ws["T1"].value.startswith("Moon") and ws["U1"].value.startswith("Future as")
    assert ws["V1"].value == "Future Expiry" and ws["W1"].value == "Comments"
    assert ws["W2"].value == "my note" and ws["W3"].value == "second note"      # notes intact
    assert [ws.cell(2, c).value for c in range(15, 20)] == [61.0, 62.0, 63.0, 64.0, 65.0]
    assert ws["T2"].value                                                       # nakshatra
    assert ws["U3"].value == 25000.0 and ws["V3"].value == "NIFTY Jan 2026 Fut"
    assert ws["U3"].comment is not None and ws["U3"].comment.text == "hello"    # cell comment moved
    assert ws["M3"].value == "=U3-F3"                                           # diff re-pointed


def test_layout_B_to_C_adds_rsi60_without_losing_anything(tmp_path):
    p = _build_c(tmp_path)
    wb = load_workbook(p); ws = wb.active
    # turn it back into layout B (no RSI 60m column; everything after shifts left)
    for r in range(1, 4):
        for c in range(20, 24):
            ws.cell(r, c - 1).value = ws.cell(r, c).value
        ws.cell(r, 23).value = None
    from openpyxl.comments import Comment
    cm = ws["U3"].comment                       # the Future cell's comment sits one column left in layout B
    ws["U3"].comment = None
    ws["T3"].comment = Comment(cm.text, cm.author)
    ws["M3"] = "=T3-F3"
    wb.save(p)
    # next run brings RSI 60m data -> the new (empty) column is filled, nothing else changes
    update_workbook(p, _df(["2026-01-05", "2026-01-06"]), rsi={"2026-01-05": {60: 65.0, 5: 99.0}})
    _check_c(p)
    # idempotent
    snap = [[c.value for c in row] for row in load_workbook(p).active.iter_rows()]
    update_workbook(p, _df(["2026-01-05", "2026-01-06"]))
    assert snap == [[c.value for c in row] for row in load_workbook(p).active.iter_rows()]


def test_layout_A_to_C_oldest_workbook(tmp_path):
    from scripts.excel_writer import _OLD_TO_NEW
    p = _build_c(tmp_path)
    wb = load_workbook(p); ws = wb.active
    # layout C -> B -> A (reverse both migrations by hand)
    for r in range(1, 4):                      # C -> B: drop RSI 60m column
        for c in range(19, 23):
            ws.cell(r, c).value = ws.cell(r, c + 1).value
        ws.cell(r, 23).value = None
    new_to_old = {n: o for o, n in _OLD_TO_NEW.items()}
    snap = {(r, c): ws.cell(r, c).value for r in range(1, 4) for c in range(14, 23)}
    for (r, c), v in snap.items():             # B -> A
        ws.cell(r, new_to_old[c]).value = v
    ws["M3"] = "=N3-F3"
    wb.save(p)
    update_workbook(p, _df(["2026-01-05", "2026-01-06"]))
    ws2 = load_workbook(p).active
    # the 60m value did not exist in layout A, so it is legitimately blank (filled by the next RSI fetch)
    assert ws2["S2"].value is None
    assert ws2["W2"].value == "my note" and ws2["W3"].value == "second note"
    assert ws2["M3"].value == "=U3-F3" and [ws2.cell(2, c).value for c in range(15, 19)] == [61.0, 62.0, 63.0, 64.0]


def test_formatted_as_table_for_filtering(tmp_path):
    p = str(tmp_path / "t.xlsx")
    create_workbook(p)
    update_workbook(p, _df(["2026-01-05", "2026-01-06", "2026-01-07"]))
    ws = load_workbook(p).active
    assert list(ws.tables) == ["NSE_Table"]
    t = ws.tables["NSE_Table"]
    assert t.ref == "A1:W4"
    # header names are single-line, unique strings (required by Excel tables)
    heads = [ws.cell(1, c).value for c in range(1, 24)]
    assert all(isinstance(h, str) and "\n" not in h for h in heads) and len(set(heads)) == 23
    # a missing earlier day is inserted -> table grows; rerun keeps exactly one table
    update_workbook(p, _df(["2026-01-02", "2026-01-05", "2026-01-06", "2026-01-07"]))
    ws = load_workbook(p).active
    assert list(ws.tables) == ["NSE_Table"] and ws.tables["NSE_Table"].ref == "A1:W5"
    update_workbook(p, _df(["2026-01-02"]))
    assert load_workbook(p).active.tables["NSE_Table"].ref == "A1:W5"
