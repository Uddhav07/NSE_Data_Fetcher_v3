"""Excel workbook creation, formatting, and update logic."""

from __future__ import annotations

import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .nakshatra import moon_nakshatra
from .paths import get_archive_dir

logger = logging.getLogger(__name__)

# ── Styles ────────────────────────────────────────────────────────────────────

_BLUE_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
_WHITE_BOLD = Font(bold=True, color="FFFFFF", size=10)
_YELLOW_FILL = PatternFill(start_color="BF8F00", end_color="BF8F00", fill_type="solid")
_TUESDAY_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
_GREEN_FONT = Font(color="006100")
_RED_FONT = Font(color="9C0006")
_GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
_RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
_RSI70_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")   # light green
_RSI70_FONT = Font(color="006100", bold=True)
_RSI80_FILL = PatternFill(start_color="00682F", end_color="00682F", fill_type="solid")   # dark green
_RSI80_FONT = Font(color="FFFFFF", bold=True)
_RSI30_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")   # red
_RSI30_FONT = Font(color="FFFFFF", bold=True)
_TOTAL_FILL = PatternFill(start_color="D9B3FF", end_color="D9B3FF", fill_type="solid")   # purple
_TOTAL_FONT = Font(color="4B0082", bold=True)
_THIN_BORDER = Border(
    bottom=Side(style="thin", color="D9D9D9"),
)
_CENTER = Alignment(horizontal="center", vertical="center")
_LEFT = Alignment(horizontal="left", vertical="center", indent=0)
_CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Column spec: (header_text, width, header_group)
# header_group: 'blue', 'yellow', 'plain'
COLUMNS = [
    ("Day",                          13.0,  "plain"),   # A
    ("Date",                         11.0,  "blue"),    # B
    ("Open",                         10.0,  "blue"),    # C
    ("High",                         13.0,  "blue"),    # D
    ("Low",                          13.0,  "blue"),    # E
    ("Close",                        13.0,  "blue"),    # F
    ("Daily Change",                  9.5,  "blue"),    # G
    ("Weekly Change",                13.0,  "blue"),    # H
    ("Close Above prev High",      12.0,  "yellow"),  # I
    ("Close Below prev Low",       12.0,  "yellow"),  # J
    ("HH / HL",                     10.5,  "yellow"),  # K
    ("LL / LH",                     10.5,  "yellow"),  # L
    ("Diff in Future",              9.5,  "plain"),   # M  (= T - Close)
    ("Date Total",                  10.0,  "plain"),   # N  "day, digit-root of full date"
    ("RSI 5m",                       8.0,  "plain"),   # O
    ("RSI 15m",                      8.0,  "plain"),   # P
    ("RSI 30m",                      8.0,  "plain"),   # Q
    ("RSI 45m",                      8.0,  "plain"),   # R
    ("Moon Nakshatra (9:15am)",     18.0,  "plain"),   # S
    ("Future as on Date",          13.0,  "plain"),   # T
    ("Future Expiry",              13.0,  "plain"),   # U
    ("Comments",                    30.0,  "plain"),   # V — user notes (never overwritten)
]
_DATE_TOTAL_COL = 14
_RSI_COLS = {5: 15, 15: 16, 30: 17, 45: 18}   # minutes -> column index
_NAKSHATRA_COL = 19
_DIFF_COL = 13
_FUT_COL = 20
_EXP_COL = 21
_COMMENTS_COL = 22


def _apply_header_style(ws) -> None:
    """Apply header styles and column widths."""
    for col_idx, (text, width, group) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=text)
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width + 2  # room for the filter button

        if group == "blue":
            cell.fill = _BLUE_FILL
            cell.font = _WHITE_BOLD
            cell.alignment = _CENTER_WRAP
        elif group == "yellow":
            cell.fill = _YELLOW_FILL
            cell.font = _WHITE_BOLD
            cell.alignment = _CENTER_WRAP
        else:
            cell.font = Font(bold=True, size=10)
            cell.alignment = _CENTER_WRAP


def _ensure_headers(ws) -> None:
    """Add any missing header cells (e.g. new columns on an older workbook).

    Existing header cells are never touched, so user edits to headers stay.
    """
    for col_idx, (text, width, group) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        if cell.value in (None, ""):
            cell.value = text
            cell.font = Font(bold=True, size=10)
            cell.alignment = _CENTER_WRAP
            letter = get_column_letter(col_idx)
            if ws.column_dimensions[letter].width in (None, 13.0) or letter not in ws.column_dimensions:
                ws.column_dimensions[letter].width = width + 2  # room for the filter button


# Old layout (<= v3.2): N Future, O Expiry, P Comments, Q Date Total, R-U RSI, V Nakshatra.
# New layout: N Date Total, O-R RSI, S Nakshatra, T Future, U Expiry, V Comments.
_OLD_TO_NEW = {14: 20, 15: 21, 16: 22, 17: 14, 18: 15, 19: 16, 20: 17, 21: 18, 22: 19}


def _migrate_layout(ws) -> bool:
    """Move columns N:V from the old order to the new one, keeping every value,
    cell style, cell comment and your Comments notes. Returns True if it migrated."""
    if not str(ws.cell(1, 16).value or "").strip().startswith("Comments"):
        return False
    snap = {}
    for r in range(1, ws.max_row + 1):
        for old in _OLD_TO_NEW:
            c = ws.cell(r, old)
            cm = c.comment
            snap[(r, old)] = (c.value, copy(c._style),
                              (cm.text, cm.author, cm.width, cm.height) if cm else None)
            c.comment = None
    for (r, old), (value, style, cm) in snap.items():
        tgt = ws.cell(r, _OLD_TO_NEW[old])
        tgt.value = value
        tgt._style = style
        if cm:
            tgt.comment = Comment(cm[0], cm[1], height=cm[3], width=cm[2])
    for r in range(2, ws.max_row + 1):  # "Diff in Future" now points at column T
        m = ws.cell(r, _DIFF_COL)
        if isinstance(m.value, str):
            m.value = re.sub(r"^=N(\d+)-F(\d+)$", r"=T\1-F\2", m.value)
    for col, (_, width, _g) in enumerate(COLUMNS, start=1):
        if col >= 13:
            ws.column_dimensions[get_column_letter(col)].width = width + 2  # room for the filter button
    logger.info("Column layout updated: Future / Expiry / Comments moved to the end.")
    return True


def _as_date(val):
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        try:
            return datetime.strptime(val[:10], "%Y-%m-%d")
        except ValueError:
            return None
    return None


def _fill_derived_columns(ws, rsi: Optional[dict]) -> None:
    """Fill Nakshatra + RSI cells that are EMPTY. Non-empty cells are never changed."""
    for row in range(2, ws.max_row + 1):
        dt = _as_date(ws.cell(row, 2).value)
        if dt is None:
            continue
        targets = []
        day_rsi = (rsi or {}).get(dt.strftime("%Y-%m-%d"), {})
        c = ws.cell(row, _NAKSHATRA_COL)
        if c.value in (None, ""):
            nk = moon_nakshatra(dt).strip()
            if nk:
                c.value = nk
                targets.append(c)
        for minutes, col in _RSI_COLS.items():
            c = ws.cell(row, col)
            if c.value in (None, "") and minutes in day_rsi:
                c.value = day_rsi[minutes]
                c.number_format = "0.00"
                targets.append(c)
        for c in targets:
            c.alignment = _LEFT if c.column == _NAKSHATRA_COL else _CENTER
            c.border = _THIN_BORDER
            if dt.weekday() == 1 and c.fill.fill_type is None:
                c.fill = _TUESDAY_FILL


def _scroll_to_latest(ws, last_row: int) -> None:
    """Scroll so the newest row sits near the bottom of the window and select it."""
    top = max(2, last_row - 25)
    view = ws.sheet_view
    if view.pane is not None:
        view.pane.topLeftCell = f"A{top}"
    else:
        view.topLeftCell = f"A{top}"
    for sel in view.selection:
        if sel.pane in (None, "bottomLeft"):
            sel.activeCell = f"B{last_row}"
            sel.sqref = f"B{last_row}"


def create_workbook(excel_path: str) -> None:
    """Create a new, empty Excel workbook with formatted headers."""
    path = Path(excel_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "NSE_Data"

    _apply_header_style(ws)

    # Freeze the top row so headers stay visible when scrolling
    ws.freeze_panes = "A2"

    wb.save(str(path))
    logger.info("Created new workbook: %s", excel_path)


def backup_workbook(excel_path: str) -> Optional[str]:
    """Create a timestamped backup copy in archive/. Returns backup path or None."""
    src = Path(excel_path)
    if not src.exists():
        return None
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = get_archive_dir() / f"{src.stem}_{ts}{src.suffix}"
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(src), str(dest))
    logger.info("Backup saved: %s", dest)
    return str(dest)


def get_last_date(excel_path: str) -> Optional[datetime]:
    """Read the last date entry from column B of the workbook."""
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        ws = wb.active
        val = None
        for (cell,) in ws.iter_rows(min_col=2, max_col=2, min_row=2):
            if isinstance(cell.value, (datetime, str)) and cell.value != "":
                val = cell.value
        wb.close()
        if isinstance(val, datetime):
            return val
        if isinstance(val, str):
            return pd.to_datetime(val)
        return None
    except Exception as exc:
        logger.error("Error reading last date: %s", exc)
        return None


def _add_conditional_formatting(ws) -> None:
    """Add sheet-level conditional formatting rules for change columns.

    Rules skip Tuesday rows (WEEKDAY($B2,2)=2) so the Tuesday highlight
    fill is never overridden by the green/red formatting.
    """
    from openpyxl.formatting.rule import FormulaRule

    # Drop our own previous RSI / Date Total rules so colour changes apply
    _ours = {("R2:U1048576", "80"), ("R2:U1048576", "70"),
             ("R2:U1048576", "AND(ISNUMBER(R2),R2<30)"),
             ("R2:U1048576", "AND(ISNUMBER(R2),R2>=80)"),
             ("R2:U1048576", "AND(ISNUMBER(R2),R2>=70,R2<80)"),
             ("Q2:Q1048576", 'OR(Q2="3, 6",Q2="6, 3")'),
             ("O2:R1048576", "AND(ISNUMBER(O2),O2<30)"),
             ("O2:R1048576", "AND(ISNUMBER(O2),O2>=80)"),
             ("O2:R1048576", "AND(ISNUMBER(O2),O2>=70,O2<80)"),
             ("N2:N1048576", 'OR(N2="3, 6",N2="6, 3")')}
    for cf in list(ws.conditional_formatting):
        cf.rules = [r for r in cf.rules
                    if (str(cf.sqref), (r.formula or [""])[0]) not in _ours]
        if not cf.rules:
            del ws.conditional_formatting._cf_rules[cf]

    # Keep rules already in the sheet (ours or the user's); add only missing ones
    existing = set()
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            existing.add((str(cf.sqref), tuple(rule.formula or ())))
    _orig_add = ws.conditional_formatting.add

    class _Guard:
        @staticmethod
        def add(rng, rule):
            if (rng, tuple(rule.formula or ())) not in existing:
                _orig_add(rng, rule)

    guard = _Guard

    # RSI (R:U) — cell-only colours; <30 red, 70-79.99 light green, 80+ dark green (mutually exclusive)
    guard.add("O2:R1048576", FormulaRule(formula=["AND(ISNUMBER(O2),O2<30)"],
              fill=_RSI30_FILL, font=_RSI30_FONT))
    guard.add("O2:R1048576", FormulaRule(formula=["AND(ISNUMBER(O2),O2>=80)"],
              fill=_RSI80_FILL, font=_RSI80_FONT))
    guard.add("O2:R1048576", FormulaRule(formula=["AND(ISNUMBER(O2),O2>=70,O2<80)"],
              fill=_RSI70_FILL, font=_RSI70_FONT))
    # Date Total (Q) — only the cell, for "3, 6" and "6, 3"
    guard.add("N2:N1048576", FormulaRule(formula=['OR(N2="3, 6",N2="6, 3")'],
              fill=_TOTAL_FILL, font=_TOTAL_FONT))

    # Daily Change (G) — green if > 0, red if < 0 (skip Tuesdays)
    guard.add(
        "G2:G1048576",
        FormulaRule(formula=['AND(G2>0,WEEKDAY($B2,2)<>2)'], fill=_GREEN_FILL, font=_GREEN_FONT),
    )
    guard.add(
        "G2:G1048576",
        FormulaRule(formula=['AND(G2<0,WEEKDAY($B2,2)<>2)'], fill=_RED_FILL, font=_RED_FONT),
    )
    # Weekly Change (H)
    guard.add(
        "H2:H1048576",
        FormulaRule(formula=['AND(H2>0,WEEKDAY($B2,2)<>2)'], fill=_GREEN_FILL, font=_GREEN_FONT),
    )
    guard.add(
        "H2:H1048576",
        FormulaRule(formula=['AND(H2<0,WEEKDAY($B2,2)<>2)'], fill=_RED_FILL, font=_RED_FONT),
    )
    # Difference in future (M)
    guard.add(
        "M2:M1048576",
        FormulaRule(formula=['AND(M2>0,WEEKDAY($B2,2)<>2)'], fill=_GREEN_FILL, font=_GREEN_FONT),
    )
    guard.add(
        "M2:M1048576",
        FormulaRule(formula=['AND(M2<0,WEEKDAY($B2,2)<>2)'], fill=_RED_FILL, font=_RED_FONT),
    )


_FIRST_REF = re.compile(r"[A-Z]{1,2}(\d+)")


def _is_stale(value, row: int) -> bool:
    """A formula whose first cell reference points at a different row (rows were shifted)."""
    if not (isinstance(value, str) and value.startswith("=")):
        return False
    m = _FIRST_REF.search(value)
    return bool(m) and int(m.group(1)) != row


def _canonical(row: int, prev: Optional[int], ref5: Optional[int]) -> dict:
    """Our formulas for one row: {column: formula}."""
    f = {
        1: f'=TEXT(B{row},"dddd")',
        _DATE_TOTAL_COL: (f'=DAY(B{row})&", "&(1+MOD(YEAR(B{row})*10000+'
                          f'MONTH(B{row})*100+DAY(B{row})-1,9))'),
    }
    if prev:
        f[7] = f"=F{row}-F{prev}"
        f[9] = f'=IF(F{row}>D{prev},"High","-")'
        f[10] = f'=IF(F{row}<E{prev},"Low","-")'
        f[11] = f'=IF(D{row}>D{prev},"Higher High","LH")'
        f[12] = f'=IF(E{row}<E{prev},"L Low","Higher Low")'
    if ref5:
        f[8] = f"=F{row}-F{ref5}"
    return f


def _dated_rows(ws) -> list:
    """[(date_str, row)] for every row with a real date in column B, in sheet order."""
    out = []
    for row in range(2, ws.max_row + 1):
        dt = _as_date(ws.cell(row, 2).value)
        if dt is not None:
            out.append((dt.strftime("%Y-%m-%d"), row))
    return out


def _ensure_formulas(ws) -> None:
    """Fill EMPTY formula cells and repair formulas shifted by row insertion.

    A formula is only rewritten when it still points at its old row number
    (i.e. the row moved). Formulas you edited in place are left alone.
    """
    dated = _dated_rows(ws)
    for idx, (_, row) in enumerate(dated):
        prev = dated[idx - 1][1] if idx >= 1 else None
        ref5 = dated[idx - 5][1] if idx >= 5 else None
        for col, formula in _canonical(row, prev, ref5).items():
            cell = ws.cell(row, col)
            if cell.value in (None, ""):
                cell.value = formula
                cell.alignment = _CENTER
                if col in (7, 8):
                    cell.number_format = "#,##0.00"
            elif _is_stale(cell.value, row):
                cell.value = formula
        # futures difference (M) is ours too: repair only when shifted
        m = ws.cell(row, _DIFF_COL)
        if _is_stale(m.value, row):
            m.value = f"=T{row}-F{row}"


def _write_new_row(ws, row: int, date_idx, row_data) -> None:
    """Write a freshly inserted row's values and styling (formulas come later)."""
    ws.cell(row, 2, date_idx.to_pydatetime().replace(tzinfo=None))
    ws.cell(row, 2).number_format = "YYYY-MM-DD"
    ws.cell(row, 2).alignment = _CENTER
    for off, field in enumerate(["Open", "High", "Low", "Close"]):
        cell = ws.cell(row, 3 + off, round(float(row_data[field]), 2))
        cell.number_format = "#,##0.00"
        cell.alignment = _CENTER
    for col in range(1, len(COLUMNS) + 1):
        ws.cell(row, col).border = _THIN_BORDER
    if date_idx.weekday() == 1:  # Tuesday highlight — only on rows we create
        for col in range(1, len(COLUMNS) + 1):
            ws.cell(row, col).fill = _TUESDAY_FILL


_TABLE_NAME = "NSE_Table"


def _normalize_headers(ws) -> None:
    """Header text on one line (Excel table column names must match the cells exactly)."""
    for col in range(1, len(COLUMNS) + 1):
        c = ws.cell(1, col)
        if isinstance(c.value, str):
            c.value = " ".join(c.value.split())


def _apply_table(ws) -> None:
    """Format A1:V<last> as an Excel Table so every column has filter/sort buttons.

    Re-created on each run so its range always covers the current rows.
    """
    last_row = 1
    for row in range(ws.max_row, 1, -1):
        if any(ws.cell(row, c).value not in (None, "") for c in range(1, len(COLUMNS) + 1)):
            last_row = row
            break
    if last_row < 2:
        return
    headers = [ws.cell(1, c).value for c in range(1, len(COLUMNS) + 1)]
    if any(not isinstance(h, str) or not h.strip() for h in headers) or len(set(headers)) != len(headers):
        logger.warning("Header row has blank or duplicate names — sheet not formatted as a table.")
        return
    others = [n for n in ws.tables if n != _TABLE_NAME]
    if others:
        logger.warning("Sheet already contains another table (%s) — leaving it alone, "
                       "not creating the NSE table.", ", ".join(others))
        return
    if _TABLE_NAME in ws.tables:
        del ws.tables[_TABLE_NAME]
    ref = f"A1:{get_column_letter(len(COLUMNS))}{last_row}"
    table = Table(displayName=_TABLE_NAME, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False,
                                          showLastColumn=False, showRowStripes=False,
                                          showColumnStripes=False)
    ws.add_table(table)


def update_workbook(
    excel_path: str,
    data: pd.DataFrame,
    futures_price: Optional[float] = None,
    futures_expiry: Optional[str] = None,
    rsi: Optional[dict] = None,
) -> tuple[int, int]:
    """Bring the workbook up to date without touching what is already there.

    * Dates in ``data`` that are missing from the sheet are inserted in
      chronological order (new days at the end, older gaps in the middle).
    * Empty cells (OHLC, formulas, Date Total, RSI, Nakshatra, futures) are filled.
    * Cells that already have a value — including your edits — are never changed.
    * Formulas that move because a row was inserted are re-pointed.

    Returns (added_count, already_present_count).
    """
    try:
        wb = load_workbook(excel_path)
    except PermissionError:
        msg = (f"Cannot open '{excel_path}' — the file is locked. "
               "Please close it in Excel and try again.")
        logger.error(msg)
        raise PermissionError(msg)
    ws = wb.active
    _migrate_layout(ws)
    _ensure_headers(ws)
    _normalize_headers(ws)

    existing = dict(_dated_rows(ws))  # date_str -> row
    added = skipped = 0
    shifted = False  # True if a missing day was inserted above existing rows

    # 1) existing dates: fill only EMPTY OHLC cells
    for date_idx, row_data in data.iterrows():
        ds = date_idx.strftime("%Y-%m-%d")
        if ds in existing:
            skipped += 1
            for off, field in enumerate(["Open", "High", "Low", "Close"]):
                cell = ws.cell(existing[ds], 3 + off)
                if cell.value in (None, ""):
                    cell.value = round(float(row_data[field]), 2)
                    cell.number_format = "#,##0.00"
                    cell.alignment = _CENTER

    # 2) missing dates: insert in order (oldest first)
    for date_idx, row_data in data.sort_index().iterrows():
        ds = date_idx.strftime("%Y-%m-%d")
        if ds in existing:
            continue
        dated = _dated_rows(ws)
        earlier = [r for d, r in dated if d < ds]
        pos = (max(earlier) + 1) if earlier else 2
        if dated and pos <= max(r for _, r in dated):
            shifted = True
        ws.insert_rows(pos)
        _write_new_row(ws, pos, date_idx, row_data)
        existing = dict(_dated_rows(ws))
        added += 1

    if shifted:
        # insert_rows moves cells but not row visibility: drop any stale filter-hidden rows
        for dim in ws.row_dimensions.values():
            dim.hidden = False
        logger.info("Rows were inserted in the middle of the sheet — filters were cleared; "
                    "re-apply them in Excel if you need them.")

    # 3) formulas: fill empty ones, re-point shifted ones
    _ensure_formulas(ws)

    dated = _dated_rows(ws)
    newest_row = max((r for _, r in dated), default=None)

    # 4) futures — newest dated row, only into empty cells
    if futures_price is not None and newest_row:
        r = newest_row
        if ws.cell(r, _FUT_COL).value in (None, ""):
            ws.cell(r, _FUT_COL, futures_price)
            ws.cell(r, _FUT_COL).number_format = "#,##0.00"
            ws.cell(r, _FUT_COL).alignment = _CENTER
            if futures_expiry and ws.cell(r, _EXP_COL).value in (None, ""):
                ws.cell(r, _EXP_COL, futures_expiry)
                ws.cell(r, _EXP_COL).alignment = _CENTER
            if ws.cell(r, _DIFF_COL).value in (None, ""):
                ws.cell(r, _DIFF_COL, f"=T{r}-F{r}")
                ws.cell(r, _DIFF_COL).number_format = "#,##0.00"
                ws.cell(r, _DIFF_COL).alignment = _CENTER
        else:
            logger.info("Futures already present on row %d — left unchanged.", r)

    # 5) Nakshatra / RSI — only into empty cells
    _fill_derived_columns(ws, rsi)

    _add_conditional_formatting(ws)
    _apply_table(ws)
    ws.freeze_panes = "A2"
    if newest_row:
        _scroll_to_latest(ws, newest_row)

    try:
        wb.save(excel_path)
    except PermissionError:
        msg = (f"Cannot save '{excel_path}' — the file is locked. "
               "Please close it in Excel and try again.")
        logger.error(msg)
        raise PermissionError(msg)
    finally:
        wb.close()

    if added:
        logger.info("Added %d new row(s) to %s.", added, excel_path)
    else:
        logger.info("No new rows to add; empty cells were filled where data was available.")
    return added, skipped
