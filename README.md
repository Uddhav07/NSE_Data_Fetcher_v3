# NSE Data Fetcher v3.2 — script edition

A daily Nifty 50 journal in one Excel file. Each run appends the new trading days'
OHLC (Yahoo `^NSEI`), price-action formulas, the current Nifty futures price/expiry
(Groww), intraday RSI at the 15:30 close, and the Moon's nakshatra at 09:15 IST.

Runs from the terminal on macOS/Linux (Python 3.9+). No GUI, no installer.
Rewritten from [v2.1](https://github.com/pmmittal-byte/NSE_Data_Fetcher_v2.1); v3.1 shipped a Windows app, which this edition drops.

## Quick start
```bash
git clone https://github.com/Uddhav07/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
./setup.sh                                    # one-time: venv + dependencies
./run.sh --full-refresh --start 2026-01-01    # one-time: build the sheet from Jan 2026
./run.sh                                      # every day after 15:35 IST
```
**macOS double-click:** use `setup.command` / `run.command` in Finder (a `.sh` file opens in an editor instead of running). If macOS blocks it the first time: right-click → Open.

**Windows (Python 3.9+, tick "Add to PATH" when installing):** double-click `setup.bat` once, then `run.bat`
(or `run.bat --full-refresh --start 2026-01-01` from a terminal for the first build). Equivalent manual commands:
```bat
python -m venv venv
venv\Scripts\python.exe -m pip install -r requirements.txt
venv\Scripts\python.exe -m scripts.main --full-refresh --start 2026-01-01
venv\Scripts\python.exe -m scripts.main
```
The Windows scripts are untested on a real Windows machine; the Python code is cross-platform and opens Excel via `os.startfile`.

Close the workbook in Excel before running (a locked file gives a friendly error).
The sheet opens scrolled to the latest date.

## Options
```
./run.sh                               incremental update from the last date in the sheet
./run.sh --full-refresh --start DATE   rebuild (old file is backed up to archive/ first)
./run.sh --no-futures                  skip Groww futures
./run.sh --no-rsi                      skip intraday RSI
./run.sh --config PATH                 use another config.json
```

## Excel layout (sheet `NSE_Data`)
| Col | Content |
|---|---|
| A–B | Day, Date |
| C–F | Open, High, Low, Close |
| G–H | Daily change; Weekly change (vs 5 trading days back, holiday-aware) |
| I–L | Close above prev High / below prev Low, HH-HL, LL-LH (formulas) |
| M–O | Futures diff, futures price, expiry (newest row only) |
| P | **Comments** — yours, never written by the script |
| Q | **Date Total** — `day, digit-root of full date`; 2026-11-09 → `9, 3` |
| R–U | **RSI(14)** on 5/15/30/45-min candles, as of the 15:30 IST close |
| V | **Moon Nakshatra** at 09:15 IST (sidereal, Lahiri), left-aligned |

Cell-only colours: RSI below 30 red, RSI 70–79.99 light green, RSI 80+ dark green, Date Total `3, 6` / `6, 3` purple.
Daily/weekly/futures change columns go green/red. Tuesday rows get a light-blue row fill when created.

Notes on the intraday data: Yahoo keeps only ~60 days of 5/15/30-min history, so older rows have
no RSI. 45-min candles are built from 15-min ones. Today's RSI is filled only on a run after 15:35 IST.

## Your edits are safe
The script only **adds rows** and **fills empty cells**. It never rewrites a cell that has a value,
nor your fills, fonts, headers, notes or your own conditional-format rules. Futures are written once per
row. Limitation: charts/images/pivots added to the sheet are dropped when openpyxl saves.

## Configuration — `config/config.json`
`ticker` (`^NSEI`, `^NSEBANK`, `^BSESN`) · `excel_file` · `start_date` · `log_level` ·
`max_retries` · `request_timeout` · `open_excel_after_run` · `backup_before_update`

## Project layout
```
setup.sh/run.sh (+ .command wrappers) · setup.bat/run.bat (Windows)
scripts/  main.py (CLI+orchestration) · fetcher.py (Yahoo OHLC) · futures.py (Groww)
          rsi.py (intraday RSI) · nakshatra.py (Moon) · excel_writer.py · config_manager.py · paths.py
config/   config.json        tests/  test_core.py (pytest, no network)
data/ archive/ logs/         generated at runtime (git-ignored)
docs/     ARCHITECTURE.md · PLAN.md        CHANGELOG.md
```
Run tests: `venv/bin/pip install -r requirements-dev.txt && venv/bin/python -m pytest tests`

## Troubleshooting
| Problem | Fix |
|---|---|
| "file is locked" | Close the workbook in Excel |
| "No new data" | Market closed or already up to date (futures/RSI still refresh) |
| Futures blank | Groww page changed or down; OHLC is still saved |
| RSI blank on old rows | Yahoo intraday history limit (~60 days) |

Logs: `logs/nse_fetcher.log`. Provided as-is for personal and educational use; not financial advice.
