# NSE Data Fetcher v3 — project instructions

**What it does:** appends Nifty 50 daily OHLC (Yahoo `^NSEI`), futures (Groww), intraday RSI at the 15:30 close and Moon nakshatra to `data/stock_market_NSE50.xlsx` with formula columns and
formatting. See `docs/ARCHITECTURE.md` (how it works) and `docs/PLAN.md` (what we're doing).

## Scope
- Script/CLI use only (macOS/Linux, Python 3.9+). No Windows GUI, `.bat`, PyInstaller, Inno Setup.

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
- Fill only EMPTY cells; never overwrite user values/formatting. Highlights are **cell-only**, never whole rows.
- Columns: Q Date Total, R–U RSI (5/15/30/45m at 15:30 IST), V Moon Nakshatra (09:15 IST).
- Don't hit Yahoo/Groww in tests; mock `fetch_ohlc_data`/`fetch_futures`.
- Confirm before deleting files or `--full-refresh` against the user's real workbook.
