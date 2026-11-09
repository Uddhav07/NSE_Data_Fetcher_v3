# NSE Data Fetcher v3.2 — Architecture

## Use case
Daily journal of the **Nifty 50** (`^NSEI`) in one Excel file. Each run appends the
new trading days' OHLC, adds price-action formulas (daily/weekly change, close vs
previous high/low, HH/HL, LL/LH), and stamps the current **Nifty futures** price and
expiry on the latest row so the user can see futures premium/discount. Tuesdays are
highlighted (weekly-expiry day, a trader's convention). Column P "Comments" is the
user's own notes and is never overwritten.

## Pipeline (`scripts/main.py:run_fetch`)

```
ensure_dirs → save_default_config → load_config(+validate)
  → env health check (py>=3.10, packages, Yahoo reachable)
  → decide start date:
       no xlsx / --full-refresh → create_workbook, start = --start or config.start_date
       xlsx exists              → start = last date in column B + 1 day
  → backup_workbook to archive/ (timestamped)
  → fetch_ohlc_data  (yfinance, retries 2^n backoff)  → empty ⇒ "up to date", stop
  → fetch_futures    (Groww scrape, retries)          → skipped with --no-futures
  → update_workbook  (append rows, formulas, formatting, save)
  → optionally open Excel
```
`run_fetch` never calls `sys.exit`; it returns `FetchResult(added, skipped, error, excel_path)`.
The CLI (`run`) and the GUI both call it — **the GUI is not needed for our use**.

## Modules

| File | Role |
|---|---|
| `scripts/main.py` | argparse CLI, logging setup (console + 5 MB×3 rotating file), health check, orchestration |
| `scripts/config_manager.py` | `AppConfig` dataclass, type-checked JSON load, `validate()`, writes default config |
| `scripts/paths.py` | base/data/logs/archive/config dirs; frozen (PyInstaller) vs dev |
| `scripts/fetcher.py` | `yf.Ticker(t).history(start, end=tomorrow)`; keep OHLC, drop NaN/zero rows, round 2dp |
| `scripts/futures.py` | GET `groww.in/futures/nifty`, parse `__NEXT_DATA__` → `contractData.livePrice.{ltp,symbol}`; expiry label from symbol `NIFTY25NOVFUT` → "NIFTY Nov 2025 Fut" |
| `scripts/excel_writer.py` | workbook creation, row append, formulas, Tuesday fill, conditional formats |
| `scripts/rsi.py` | 5/15/30m candles (period 60d) → Wilder RSI(14) → last value per date; 45m built from 15m; today skipped before 15:35 IST |
| `scripts/nakshatra.py` | Moon sidereal longitude (pure-Python Meeus series in `moon_calc.py` + Lahiri ayanamsa) at 03:45 UT (09:15 IST) → one of 27 nakshatras; verified against Swiss Ephemeris over 3,650 days (max 0.008°, 0 mismatches) |

## Excel layout (sheet `NSE_Data`, header row frozen)

| Col | Header | Content |
|---|---|---|
| A | Day | `=TEXT(B,"dddd")` |
| B | Date | real datetime, `YYYY-MM-DD` |
| C–F | Open High Low Close | values, `#,##0.00` |
| G | Daily Change | `=F_n - F_(n-1)` (not on first row) |
| H | Weekly Change | `=F_n - F_(row of 5th prior trading date)` (needs ≥5 prior rows) |
| I | Close Above prev High | `=IF(F_n>D_prev,"High","-")` |
| J | Close Below prev Low | `=IF(F_n<E_prev,"Low","-")` |
| K | HH / HL | `=IF(D_n>D_prev,"Higher High","LH")` |
| L | LL / LH | `=IF(E_n<E_prev,"L Low","Higher Low")` |
| M | Diff in Future | `=T-F` (only on futures row) |
| N | Date Total | `=DAY(B)&", "&(1+MOD(yyyymmdd-1,9))` |
| O–R | RSI 5m/15m/30m/45m | static numbers at 15:30 IST close |
| S | Moon Nakshatra | static text, left-aligned |
| T | Future as on Date | scraped LTP (latest row only) |
| U | Future Expiry | e.g. "NIFTY Nov 2026 Fut" |
| V | Comments | user-owned, always the last column |

Header colors: B–H blue `366092`, I–L gold `BF8F00`, others plain. Tuesday rows: whole
row A–P filled (theme 9, tint 0.4). Conditional formats on G, H, M: green >0 / red <0,
skipped on Tuesdays via `WEEKDAY($B2,2)<>2`. Duplicate dates are skipped, never overwritten.

## Key design properties to preserve
1. **Incremental & idempotent** — last date in col B drives the next fetch; date→row map prevents dupes.
2. **Formulas, not values** for derived columns, so the user can edit/extend in Excel.
3. **Backup before every write** to `archive/`.
4. **Futures only on the newest row** (v2.1 bug stamped one price on every row).
5. **Locked-file handling** — PermissionError → friendly "close Excel" message.
6. **Retry with exponential backoff** on both network sources.

## Backfill & bootstrap (v3.3)
`run_fetch` fetches from `min(start_date, last_date+1)`. `update_workbook` inserts missing dates in chronological order (`insert_rows`), re-points formulas whose first reference no longer matches their row (`_ensure_formulas`), and fills empty cells only. `bootstrap.ensure_dependencies()` runs before pandas is imported and, if needed, builds `venv/`, installs requirements and re-executes inside it. Rows/futures for today are skipped while the session is open (`market.session_in_progress`).

## Edit-safety contract (v3.2)
Rows are only appended; Q–V (and futures M–O) are filled only into **empty** cells; headers are added only if missing; Tuesday fill only on new rows; conditional formats are added only if absent (our own RSI/Date Total rules are replaced on update so colour changes apply). Colours are always **cell-only**.

## Issues found in the original — status (all addressed in 3.2.0)
1. `os.startfile` in `main.py:_open_excel` is **Windows-only** → use `open` on macOS (`subprocess.run(["open", path])`).
2. Futures are only written if OHLC returned new rows; docstring/README claim "weekends/holidays" works — it doesn't (early return on empty data).
3. README says weekly change = "~7 calendar days back" and Tuesday fill "A–H"; code does **5 trading rows back** and fills **A–P**. Treat the code as truth.
4. `get_last_date` uses `ws.max_row` → a stray comment/format in a lower row can break the last date. Better: scan col B upward for the last real date.
5. Groww scraping depends on a private `__NEXT_DATA__` structure — most fragile part; keep a clean fallback (skip futures + log) and isolate it behind `fetch_futures()`.
6. `K` writes "LH" and `L` writes "Higher Low" as the else-branches — labels are asymmetric (kept for fidelity with v2.1 workflows).
7. Tuesday highlight uses openpyxl theme color 9 — renders differently across themes; use a fixed RGB instead.
8. Unused imports / dead `get_base_dir` import; `yfinance` `end=tomorrow` returns today's partial bar if run during market hours — decide whether to drop incomplete day.
9. Config `start_date` default is computed once; fine.
