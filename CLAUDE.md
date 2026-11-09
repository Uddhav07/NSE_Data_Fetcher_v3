# NSE Data Fetcher v3 — project instructions

**What it does:** appends Nifty 50 daily OHLC (Yahoo `^NSEI`), futures (Groww), intraday RSI at the 15:30 close and Moon nakshatra to `data/stock_market_NSE50.xlsx` with formula columns and
formatting. See `docs/ARCHITECTURE.md` (how it works) and `docs/PLAN.md` (what we're doing).

## Scope
- Script/CLI use only (macOS/Linux, Python 3.9+). No Windows GUI, `.bat`, PyInstaller, Inno Setup.

## Environment
- `scripts/bootstrap.py` auto-creates `venv/` and installs `requirements.txt` if packages are missing (never rely on a global Python). No compiled/optional packages: the Moon is calculated in pure Python (`moon_calc.py`, validated against Swiss Ephemeris: <0.01°).
- Never hard-code a user folder in docs/code; the user chooses the project location.

## Run
```
./setup.sh && ./run.sh        # also: --full-refresh --start DATE, --no-futures, --no-rsi
venv/bin/python -m pytest tests
```
Config: `config/config.json`. Logs: `logs/nse_fetcher.log`. Backups: `archive/`.

## Rules when editing
- Keep `run_fetch()` returning `FetchResult` and never calling `sys.exit` (only `run()` does).
- Derived Excel columns stay **formulas**; column P (Comments) is user-owned — never write it.
- Never overwrite existing dates; every write is preceded by a backup in `archive/`.
- Futures are written only on the newest row; failure of futures must not block OHLC.
- Excel layout (columns A–V, order, header colors) is a contract with the user's existing sheet — don't reorder.
- Code is truth where README disagrees (weekly change = 5 trading rows back; Tuesday fill A–P).
- The data range is an Excel Table (`NSE_Table`): header names must stay unique single-line strings matching the cells; the table ref is rebuilt each run (`_apply_table`). Never sort rows (formulas use the previous row).
- Runs re-check from start_date: insert missing dates in order, repair shifted formulas, fill empty cells. Fill only EMPTY cells (exception: for the latest 3 sessions, OHLC/RSI cells the script wrote — tracked in `<workbook>.state.json` — may be refreshed ONLY if still equal to the ledger value; never touch user-edited cells; `--refresh-recent` is the explicit override); never overwrite user values/formatting. Highlights are **cell-only**, never whole rows.
- Column order: M Diff in Future, N Date Total, O–S RSI (5/15/30/45/60m at 15:30 IST), T Moon Nakshatra (09:15 IST), U Future, V Expiry, W Comments (always last). Old layouts are migrated by `_migrate_layout`.
- Don't hit Yahoo/Groww in tests; mock `fetch_ohlc_data`/`fetch_futures`.
- Confirm before deleting files or `--full-refresh` against the user's real workbook.
