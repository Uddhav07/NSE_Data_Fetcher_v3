"""Characterization ("golden") test: locks in the exact workbook the script produces.

A fixed two-run scenario (first fill, user edits, then a run that back-fills a gap,
adds a new Tuesday, refreshes recent values and stamps futures) is serialised cell by
cell — value, formula, number format, fill, font, alignment, border, comment — plus
conditional formats, table, freeze pane, view, column widths and the refresh ledger.

Any change in output fails this test. If a change is intended, regenerate with
    UPDATE_GOLDEN=1 venv/bin/python -m pytest tests/test_golden.py
and review the JSON diff in git before committing.
"""
from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from scripts.excel_writer import create_workbook, update_workbook

from conftest import make_ohlc

GOLDEN = Path(__file__).parent / "golden" / "workbook_scenario.json"

ALL_DATES = ["2026-01-01", "2026-01-02", "2026-01-05", "2026-01-06", "2026-01-07",
             "2026-01-08", "2026-01-09", "2026-01-12", "2026-01-13"]
FIRST_RUN = [d for d in ALL_DATES if d not in ("2026-01-07", "2026-01-13")]

# RSI values chosen to cross every colour band (<30, 70-79.99, >=80) and leave gaps
RSI_RUN1 = {
    "2026-01-01": {5: 25.0, 15: 45.5, 30: 50.0, 45: 55.0, 60: 60.0},
    "2026-01-02": {5: 70.0, 15: 79.99, 30: 80.0, 45: 85.25, 60: 29.99},
    "2026-01-05": {5: 50.5, 15: 51.5},
    "2026-01-06": {5: 65.0, 15: 66.0, 30: 67.0, 45: 68.0, 60: 69.0},
    "2026-01-08": {5: 40.0, 60: 41.0},
    "2026-01-09": {5: 71.0, 15: 72.0, 30: 73.0, 45: 74.0, 60: 75.0},
    "2026-01-12": {5: 33.0, 15: 34.0, 30: 35.0, 45: 36.0, 60: 37.0},
}


def _color(c):
    if c is None:
        return None
    if c.type == "rgb":
        return c.rgb
    return f"{c.type}:{c.value}:{c.tint}"


def _val(v):
    if isinstance(v, datetime):
        return "dt:" + v.isoformat()
    return v


def snapshot(path: str) -> dict:
    ws = load_workbook(path).active
    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=23):
        for c in row:
            entry = {
                "v": _val(c.value),
                "fmt": c.number_format,
                "fill": [c.fill.fill_type, _color(c.fill.fgColor)],
                "font": [bool(c.font.bold), _color(c.font.color), c.font.sz],
                "align": [c.alignment.horizontal, c.alignment.wrap_text],
                "border_b": c.border.bottom.style if c.border and c.border.bottom else None,
            }
            if c.comment:
                entry["comment"] = c.comment.text
            cells[c.coordinate] = entry
    cf = []
    for rng in ws.conditional_formatting:
        for r in rng.rules:
            cf.append([str(rng.sqref), r.type, list(r.formula or []), r.priority,
                       _color(r.dxf.fill.fgColor) if r.dxf and r.dxf.fill else None,
                       _color(r.dxf.font.color) if r.dxf and r.dxf.font else None])
    view = ws.sheet_view
    ledger = Path(path).with_suffix(".state.json")
    return {
        "title": ws.title,
        "max_row": ws.max_row,
        "cells": cells,
        "conditional_formats": cf,
        "tables": {n: ws.tables[n].ref for n in ws.tables},
        "freeze": ws.freeze_panes,
        "top_left": view.pane.topLeftCell if view.pane is not None else view.topLeftCell,
        "selection": [[s.pane, s.activeCell, str(s.sqref)] for s in view.selection],
        "widths": {k: d.width for k, d in sorted(ws.column_dimensions.items())},
        "hidden_rows": sorted(r for r, d in ws.row_dimensions.items() if d.hidden),
        "ledger": json.loads(ledger.read_text()) if ledger.exists() else None,
    }


def run_scenario(tmp_path) -> dict:
    p = str(tmp_path / "golden.xlsx")
    create_workbook(p)
    update_workbook(p, make_ohlc(FIRST_RUN), 25000.5, "NIFTY Jan 2026 Fut", rsi=RSI_RUN1)

    # the user works in the sheet between runs
    wb = load_workbook(p)
    ws = wb.active
    ws["W4"] = "my note on 01-05"
    ws["F3"] = 24999.99                                   # edited Close (old row)
    ws["N5"].fill = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")  # Tuesday 01-06
    ws["P8"] = 12.34                                      # edited RSI 15m on 01-12 (recent row)
    wb.save(p)

    # next run: full range, final values for the latest sessions, one more day
    data = make_ohlc(ALL_DATES)
    data.loc["2026-01-12", "Close"] += 11.11              # data source revised the close
    rsi2 = dict(RSI_RUN1)
    rsi2["2026-01-12"] = {5: 81.0, 15: 82.0, 30: 83.0, 45: 84.0, 60: 85.0}
    rsi2["2026-01-13"] = {5: 55.0, 15: 56.0, 30: 57.0, 45: 58.0, 60: 59.0}
    rsi2["2026-01-07"] = {60: 44.4}
    update_workbook(p, data, 25100.25, "NIFTY Jan 2026 Fut", rsi=rsi2)
    return snapshot(p)


def test_workbook_matches_golden(tmp_path):
    got = json.loads(json.dumps(run_scenario(tmp_path)))     # normalise tuples/keys like the file
    if os.environ.get("UPDATE_GOLDEN") or not GOLDEN.exists():
        GOLDEN.parent.mkdir(exist_ok=True)
        GOLDEN.write_text(json.dumps(got, indent=1, sort_keys=True) + "\n")
    want = json.loads(GOLDEN.read_text())
    for key in want:
        if key == "cells":
            diff = {k: (want["cells"].get(k), got["cells"].get(k))
                    for k in set(want["cells"]) | set(got["cells"])
                    if want["cells"].get(k) != got["cells"].get(k)}
            assert not diff, f"{len(diff)} cell(s) differ, e.g. {sorted(diff.items())[:3]}"
        else:
            assert got[key] == want[key], key
    assert set(got) == set(want)


def test_golden_scenario_sanity():
    """Guard the golden file itself: key facts that must hold whatever the snapshot says."""
    g = json.loads(GOLDEN.read_text())
    c = g["cells"]
    dates = [c[f"B{r}"]["v"][3:13] for r in range(2, g["max_row"] + 1)]
    assert dates == ALL_DATES                                   # gap inserted in order
    assert c["W4"]["v"] == "my note on 01-05"                   # user note kept
    assert c["F3"]["v"] == 24999.99                             # user edit kept
    assert c["N5"]["fill"] == ["solid", "00FFFF00"]             # user fill kept (Tuesday row)
    assert c["P9"]["v"] == 12.34                                # user RSI edit kept (01-12 moved to row 9)
    assert c["O9"]["v"] == 81.0                                 # script value refreshed
    assert c["U10"]["v"] == 25100.25 and c["M10"]["v"] == "=U10-F10"   # futures on newest row
    assert c["U9"]["v"] == 25000.5                              # run 1's stamp on 01-12 kept
    assert all(c[f"U{r}"]["v"] is None for r in range(2, 9))    # never on older rows
    assert c["S6"]["v"] == 44.4                                 # back-filled gap row got RSI
    for col in "ABCDEFGHIJKLMNOPQRSTUVW":                       # new Tuesday: full-row fill
        assert c[f"{col}10"]["fill"] == ["solid", "00DDEBF7"], col
    assert g["tables"] == {"NSE_Table": "A1:W10"}
