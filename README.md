# NSE Data Fetcher v3.2 — script edition

A daily Nifty 50 journal in one Excel file. Each run appends the new trading days'
OHLC (Yahoo `^NSEI`), price-action formulas, the current Nifty futures price/expiry
(Groww), intraday RSI at the 15:30 close, and the Moon's nakshatra at 09:15 IST.

Runs from the terminal on macOS/Linux (Python 3.9+). No GUI, no installer.
Rewritten from [v2.1](https://github.com/pmmittal-byte/NSE_Data_Fetcher_v2.1); v3.1 shipped a Windows app, which this edition drops.

## Quick start (macOS / Linux)
Open a terminal **in the folder where you want the project** (you choose the location), then:
```bash
git clone https://github.com/Uddhav07/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
./setup.sh                                    # one-time: venv + dependencies
./run.sh --full-refresh --start 2026-01-01    # one-time: build the sheet from Jan 2026
./run.sh                                      # every day after 15:35 IST
```
Close the workbook in Excel before running (a locked file gives a friendly error).
The sheet opens scrolled to the latest date.

**macOS double-click:** use `setup.command` / `run.command` in Finder (a `.sh` file opens in an editor instead of running). If macOS blocks it the first time: right-click → Open.

**What a run does:** adds every missing trading day (new days at the end, older gaps inserted in date order),
fills any empty cells (OHLC, formulas, Date Total, RSI, Nakshatra, futures) and **never changes a cell that
already has a value** — your edits stay. Before 15:35 IST on a trading day, today's row and futures price are
skipped (not final yet); just run again after close.

**"No module named pandas" can't happen:** if a required package is missing from the Python you used
(system Python, IDE interpreter...), the script creates `venv/` in the project, installs the packages, and
re-runs itself inside it automatically.

## Windows — step by step (PowerShell, copy & paste)
Needs Python 3.9+ (tick **"Add python.exe to PATH"** in the installer) and Git. **You choose where the project lives:**
in File Explorer open the folder you want (e.g. `D:\Trading`), click the address bar, type `powershell`, press Enter.
Everything runs inside a project venv; the commands call `venv\Scripts\python.exe` directly, so no activation is needed.
Shortcut: double-click `setup.bat` once, then `run.bat` each day (untested on a real Windows machine).

```powershell
python --version
git clone https://github.com/Uddhav07/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
```
One-time setup (venv + packages) and the first build from Jan 2026:
```powershell
python -m venv venv
venv\Scripts\python.exe -m pip install --upgrade pip
venv\Scripts\python.exe -m pip install -r requirements.txt
venv\Scripts\python.exe -m pip install -r requirements-optional.txt
venv\Scripts\python.exe -m scripts.main --full-refresh --start 2026-01-01
```
(`requirements-optional.txt` is only for the Moon Nakshatra column; if it fails to install everything else still works.)

Every day after 15:35 IST, with the Excel file closed — open PowerShell **in the project folder** (same address-bar trick) and run:
```powershell
venv\Scripts\python.exe -m scripts.main
```
Options: `--no-futures`, `--no-rsi`. Update the project: `git pull` then
`venv\Scripts\python.exe -m pip install -r requirements.txt`.

Optional — activate the venv to type plain `python`:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned   # only if activation is blocked
venv\Scripts\Activate.ps1
python -m scripts.main
deactivate
```
In Command Prompt (cmd) activate with `venv\Scripts\activate.bat`.

## Options
```
./run.sh                               add missing days + fill empty cells (checks from start_date)
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
The script only **adds missing rows** and **fills empty cells**. It never rewrites a cell that has a value,
nor your fills, fonts, headers, notes or your own conditional-format rules. Futures are written once per
row. Limitation: charts/images/pivots added to the sheet are dropped when openpyxl saves.

## Configuration — `config/config.json`
`ticker` (`^NSEI`, `^NSEBANK`, `^BSESN`) · `excel_file` (relative to the project, or an absolute path anywhere you like) · `start_date` · `log_level` ·
`max_retries` · `request_timeout` · `open_excel_after_run` · `backup_before_update`

## Project layout
```
setup.sh/run.sh (+ .command wrappers) · setup.bat/run.bat (Windows)
scripts/  main.py (CLI+orchestration) · fetcher.py (Yahoo OHLC) · futures.py (Groww)
          rsi.py (intraday RSI) · nakshatra.py (Moon) · excel_writer.py · config_manager.py · paths.py
          bootstrap.py (auto-install deps) · market.py (session hours)
config/   config.json        tests/  test_core.py (pytest, no network)
data/ archive/ logs/         generated at runtime (git-ignored)
docs/     ARCHITECTURE.md · PLAN.md        CHANGELOG.md
```
Run tests: `venv/bin/pip install -r requirements-dev.txt && venv/bin/python -m pytest tests`

## Troubleshooting
| Problem | Fix |
|---|---|
| "file is locked" | Close the workbook in Excel |
| `No module named ...` | Should self-heal; if not: `venv/bin/python -m pip install -r requirements.txt` (Windows: `venv\\Scripts\\python.exe ...`) |
| "No new data" | Market closed or already up to date (futures/RSI still refresh) |
| Futures blank | Groww page changed or down; OHLC is still saved |
| RSI blank on old rows | Yahoo intraday history limit (~60 days) |

Logs: `logs/nse_fetcher.log`. Provided as-is for personal and educational use; not financial advice.
